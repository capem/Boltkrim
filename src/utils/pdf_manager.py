import os
import shutil
import tempfile
import fitz  # PyMuPDF
from PIL import Image
import io
import time
import socket
from datetime import datetime
from .excel_manager import is_path_available  # Reuse the network check function
from .template_manager import TemplateManager

class PDFManager:
    def __init__(self):
        self.current_file_index = -1
        self.current_file_list = []
        self.cached_pdf = None
        self.cached_pdf_path = None
        self._cached_source_folder = None
        self._last_refresh_time = 0
        self._refresh_interval = 5  # Refresh file list every 5 seconds
        self._network_timeout = 5  # 5 seconds timeout for network operations
        self._max_retries = 3  # Maximum number of retries for file operations
        self._retry_delay = 1  # Initial retry delay in seconds
        self.current_rotation = 0  # Track current rotation (0, 90, 180, 270)
        self.template_manager = TemplateManager()

    def generate_output_path(self, template, data):
        """Generate output path using template and data."""
        try:
            # Add current date to data for date-based operations ONLY if not already present
            if 'DATE FACTURE' not in data:
                data['DATE FACTURE'] = datetime.now()
            
            # Sanitize filter values to handle path characters
            invalid_chars = r'<>:"/\|?*'
            sanitized_data = data.copy()
            
            # Only sanitize filter values, not the processed_folder or date
            for key in data:
                if key.startswith('filter') and isinstance(data[key], str):
                    # Replace invalid characters with underscores
                    value = data[key]
                    for char in invalid_chars:
                        value = value.replace(char, '_')
                    sanitized_data[key] = value
            
            # Ensure we have all required fields with proper names
            if 'filter_1' in sanitized_data and 'filter1' not in sanitized_data:
                sanitized_data['filter1'] = sanitized_data['filter_1']
            if 'filter_2' in sanitized_data and 'filter2' not in sanitized_data:
                sanitized_data['filter2'] = sanitized_data['filter_2']
            
            # Process the template
            filepath = self.template_manager.process_template(template, sanitized_data)
            
            # Ensure the path is properly formatted
            filepath = os.path.normpath(filepath)
            
            # Create directory structure if it doesn't exist
            directory = os.path.dirname(filepath)
            if directory and not os.path.exists(directory):
                os.makedirs(directory, exist_ok=True)
            
            return filepath
        except Exception as e:
            raise Exception(f"Error generating output path: {str(e)}")

    def process_pdf(self, current_pdf, template_data, processed_folder, output_template):
        """Process a PDF file using template-based naming."""
        if not os.path.exists(current_pdf):
            raise Exception("Source PDF file not found")
            
        if not os.path.exists(processed_folder):
            os.makedirs(processed_folder)

        # Add processed_folder to template data
        template_data['processed_folder'] = processed_folder

        # Generate new filepath using template
        try:
            new_filepath = self.generate_output_path(output_template, template_data)
        except Exception as e:
            raise Exception(f"Error generating output path: {str(e)}")

        # Ensure we're not holding the file open
        self.ensure_pdf_not_cached(current_pdf)

        retry_count = 0
        delay = self._retry_delay

        while retry_count < self._max_retries:
            try:
                # Create temporary directory for atomic operations
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_pdf = os.path.join(temp_dir, "original.pdf")
                    rotated_pdf = os.path.join(temp_dir, "rotated.pdf")
                    
                    # Try to copy with multiple retries
                    copy_success = False
                    for _ in range(3):
                        try:
                            shutil.copy2(current_pdf, temp_pdf)
                            copy_success = True
                            break
                        except PermissionError:
                            time.sleep(delay)
                    
                    if not copy_success:
                        raise Exception("Failed to create backup copy after multiple attempts")

                    # Apply rotation if needed
                    if self.current_rotation != 0:
                        doc = fitz.open(temp_pdf)
                        page = doc[0]  # Assuming single page PDFs
                        page.set_rotation(self.current_rotation)
                        doc.save(rotated_pdf)
                        doc.close()
                        temp_pdf = rotated_pdf

                    try:
                        # Ensure target directory exists
                        os.makedirs(os.path.dirname(new_filepath), exist_ok=True)
                        
                        # Try to move the file
                        if os.path.exists(new_filepath):
                            os.remove(new_filepath)
                        
                        # Use windows-specific move operation if available
                        try:
                            import win32file
                            win32file.MoveFileEx(
                                temp_pdf,
                                new_filepath,
                                win32file.MOVEFILE_REPLACE_EXISTING | 
                                win32file.MOVEFILE_COPY_ALLOWED
                            )
                            # Explicitly remove the source file after successful move
                            os.remove(current_pdf)
                        except ImportError:
                            # Fallback to os.replace if win32file is not available
                            shutil.copy2(temp_pdf, new_filepath)
                            os.remove(current_pdf)
                        
                        # Reset rotation after successful processing
                        self.current_rotation = 0
                        return True

                    except Exception as move_error:
                        # Restore file from backup if operation fails
                        if os.path.exists(new_filepath):
                            try:
                                os.remove(new_filepath)
                            except:
                                pass
                        shutil.copy2(temp_pdf, current_pdf)
                        raise move_error

            except (PermissionError, OSError) as e:
                retry_count += 1
                if retry_count >= self._max_retries:
                    raise Exception(f"Failed to process PDF after {self._max_retries} attempts: {str(e)}")
                
                # Exponential backoff
                time.sleep(delay)
                delay *= 2
            except Exception as e:
                # Don't retry other types of errors
                raise Exception(f"Error processing PDF: {str(e)}")

        raise Exception("Failed to process PDF after exhausting all retries")

    def get_next_pdf(self, source_folder):
        """Get the next PDF file from the source folder."""
        try:
            if not is_path_available(source_folder):
                raise Exception("Network path is not available")
                
            current_time = time.time()
            
            # Only refresh file list if source folder changed or refresh interval passed
            if (source_folder != self._cached_source_folder or 
                current_time - self._last_refresh_time > self._refresh_interval):
                
                # Get list of PDF files
                try:
                    original_timeout = socket.getdefaulttimeout()
                    socket.setdefaulttimeout(self._network_timeout)
                    try:
                        pdf_files = sorted([
                            f for f in os.listdir(source_folder)
                            if f.lower().endswith('.pdf')
                        ])
                        self.current_file_list = pdf_files
                        self._cached_source_folder = source_folder
                        self._last_refresh_time = current_time
                    finally:
                        socket.setdefaulttimeout(original_timeout)
                except Exception as e:
                    if isinstance(e, socket.timeout):
                        raise Exception("Network timeout while accessing PDF folder")
                    raise Exception(f"Error reading source folder: {str(e)}")
            
            # If no PDF files found
            if not self.current_file_list:
                return None
                
            # Move to next file
            self.current_file_index += 1
            
            # If we've reached the end, start over
            if self.current_file_index >= len(self.current_file_list):
                self.current_file_index = 0
                
            # Return full path of next PDF
            if self.current_file_list:
                next_pdf = os.path.join(source_folder, self.current_file_list[self.current_file_index])
                # Clear cache if different file
                if next_pdf != self.cached_pdf_path:
                    self.clear_cache()
                return next_pdf
                
            return None
        except Exception as e:
            self.clear_cache()  # Clear cache on error
            raise e
            
    def clear_cache(self):
        """Clear the cached PDF document."""
        if self.cached_pdf:
            self.cached_pdf.close()
            self.cached_pdf = None
            self.cached_pdf_path = None
                
    def ensure_pdf_not_cached(self, pdf_path):
        """Ensure that the PDF is not cached if it matches the given path."""
        if self.cached_pdf_path == pdf_path:
            self.clear_cache()

    def rotate_page(self, clockwise=True):
        """Rotate the current PDF page clockwise or counterclockwise by 90 degrees."""
        if clockwise:
            self.current_rotation = (self.current_rotation + 90) % 360
        else:
            self.current_rotation = (self.current_rotation - 90) % 360
            
    def get_rotation(self):
        """Get the current rotation angle."""
        return self.current_rotation
        
    def render_pdf_page(self, pdf_path, page_num=0, zoom=1.0):
        """Render a PDF page as a PhotoImage."""
        try:
            if not is_path_available(pdf_path):
                raise Exception("Network path is not available")
                
            # Use cached document or open new one
            if pdf_path != self.cached_pdf_path:
                self.clear_cache()
                original_timeout = socket.getdefaulttimeout()
                socket.setdefaulttimeout(self._network_timeout)
                try:
                    self.cached_pdf = fitz.open(pdf_path)
                    self.cached_pdf_path = pdf_path
                finally:
                    socket.setdefaulttimeout(original_timeout)
            
            # Get the specified page
            page = self.cached_pdf[page_num]
            
            # Calculate matrix with optimized settings for scanned PDFs
            base_dpi = 72  # Base DPI for PDF
            quality_multiplier = 2 if zoom > 1.0 else 1  # Higher quality when zoomed in
            
            # Create matrix with optimal settings for scanned documents and rotation
            zoom_matrix = fitz.Matrix(zoom * quality_multiplier, zoom * quality_multiplier)
            if self.current_rotation:
                zoom_matrix.prerotate(self.current_rotation)
            
            # Get page as a PNG image with optimized settings for scanned documents
            pix = page.get_pixmap(
                matrix=zoom_matrix,
                alpha=False,  # No alpha channel needed for scanned docs
                colorspace="rgb",  # Force RGB colorspace
            )
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))
            
            return image
            
        except Exception as e:
            self.clear_cache()  # Clear cache on error
            if isinstance(e, socket.timeout):
                raise Exception("Network timeout while accessing PDF file")
            raise Exception(f"Error rendering PDF: {str(e)}")
