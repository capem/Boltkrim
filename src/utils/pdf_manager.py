from os import path, makedirs, remove
from os import listdir as os_listdir
from shutil import copy2
from tempfile import TemporaryDirectory
from io import BytesIO
from time import sleep, time
import re
from socket import timeout as SocketTimeout, getdefaulttimeout, setdefaulttimeout
from fitz import open as fitz_open, Matrix
from PIL.Image import open as pil_open
from typing import List, Optional, Dict, Any, Tuple
from win32file import MoveFileEx, MOVEFILE_REPLACE_EXISTING, MOVEFILE_COPY_ALLOWED
from .excel_manager import is_path_available
from .template_manager import TemplateManager
from .models import PDFTask

class PDFManager:
    def __init__(self) -> None:
        self.current_file_index: int = -1
        self.current_file_list: List[str] = []
        self.cached_pdf: Optional[Any] = None
        self.cached_pdf_path: Optional[str] = None
        self._network_timeout: int = 5  # 5 seconds timeout for network operations
        self._max_retries: int = 3  # Maximum number of retries for file operations
        self._retry_delay: int = 1  # Initial retry delay in seconds
        self.current_rotation: int = 0  # Track current rotation (0, 90, 180, 270)
        self.template_manager = TemplateManager()

    def _get_next_version_number(self, filepath: str) -> Tuple[str, int]:
        """
        Get the next available version number for a file.
        Returns tuple of (versioned_filepath, version_number)
        """
        if not path.exists(filepath):
            return filepath, 0

        directory = path.dirname(filepath)
        filename = path.basename(filepath)
        name, ext = path.splitext(filename)

        # Check if filename already has a version number
        version_match = re.match(r'^(.+)_v(\d+)$', name)
        if version_match:
            base_name = version_match.group(1)
        else:
            base_name = name

        # Find the highest existing version
        version = 1
        while True:
            versioned_name = f"{base_name}_v{version}{ext}"
            versioned_path = path.join(directory, versioned_name)
            if not path.exists(versioned_path):
                return versioned_path, version
            version += 1

    def generate_output_path(self, template, data):
        """Generate output path using template and data."""
        try:
            # Sanitize filter values to handle path characters
            invalid_chars = r'<>:"/\|?*{}[]#%&$+!`=\';,@'
            sanitized_data = data.copy()

            # Only sanitize filter values, not the processed_folder
            for key in data:
                if key.startswith("filter") and isinstance(data[key], str):
                    # Replace invalid characters with underscores
                    value = data[key]
                    for char in invalid_chars:
                        value = value.replace(char, "_")
                    sanitized_data[key] = value

            # Process the template
            filepath = self.template_manager.process_template(template, sanitized_data)

            # Ensure the path is properly formatted
            filepath = path.normpath(filepath)

            # Create directory structure if it doesn't exist
            directory = path.dirname(filepath)
            if directory and not path.exists(directory):
                makedirs(directory, exist_ok=True)

            # Check if file exists and get versioned path if needed
            if path.exists(filepath):
                filepath, _ = self._get_next_version_number(filepath)

            return filepath
        except Exception as e:
            raise Exception(f"Error generating output path: {str(e)}")

    def process_pdf(
        self,
        task: PDFTask,
        template_data: Dict[str, Any],
        processed_folder: str,
        output_template: str,
    ) -> bool:
        """Process a PDF file using template-based naming."""
        if not path.exists(task.pdf_path):
            raise Exception("Source PDF file not found")

        if not path.exists(processed_folder):
            makedirs(processed_folder)

        # Add processed_folder to template data
        template_data["processed_folder"] = processed_folder

        # Generate new filepath using template
        try:
            new_filepath = self.generate_output_path(output_template, template_data)
            
            # Store the processed_pdf_location in the task object if provided
            if task is not None:
                task.processed_pdf_location = new_filepath
        except Exception as e:
            raise Exception(f"Error generating output path: {str(e)}")

        # Ensure we're not holding the file open
        if self.cached_pdf_path == task.pdf_path:
            self.clear_cache()

        retry_count = 0
        delay = self._retry_delay

        while retry_count < self._max_retries:
            try:
                # Create temporary directory for atomic operations
                with TemporaryDirectory() as temp_dir:
                    temp_pdf = path.join(temp_dir, "original.pdf")
                    rotated_pdf = path.join(temp_dir, "rotated.pdf")

                    # Try to copy with multiple retries
                    copy_success = False
                    for _ in range(3):
                        try:
                            copy2(task.pdf_path, temp_pdf)
                            copy_success = True
                            break
                        except PermissionError:
                            sleep(delay)

                    if not copy_success:
                        raise Exception(
                            "Failed to create backup copy after multiple attempts"
                        )

                    # Apply rotation if needed
                    if self.current_rotation != 0:
                        doc = fitz_open(temp_pdf)
                        page = doc[0]  # Assuming single page PDFs
                        page.set_rotation(self.current_rotation)
                        doc.save(rotated_pdf)
                        doc.close()
                        temp_pdf = rotated_pdf

                    try:
                        # Ensure target directory exists
                        makedirs(path.dirname(new_filepath), exist_ok=True)

                        # Try to move the file
                        if path.exists(new_filepath):
                            remove(new_filepath)

                        # Use windows-specific move operation
                        MoveFileEx(
                            temp_pdf,
                            new_filepath,
                            MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED,
                        )
                        # Explicitly remove the source file after successful move
                        remove(task.pdf_path)

                        # Reset rotation tracking
                        self.current_rotation = 0
                        return True

                    except Exception as move_error:
                        # Restore file from backup if operation fails
                        if path.exists(new_filepath):
                            try:
                                remove(new_filepath)
                            except:
                                pass
                        copy2(temp_pdf, task.pdf_path)
                        raise move_error

            except (PermissionError, OSError) as e:
                retry_count += 1
                if retry_count >= self._max_retries:
                    raise Exception(
                        f"Failed to process PDF after {self._max_retries} attempts: {str(e)}"
                    )

                # Exponential backoff
                sleep(delay)
                delay *= 2
            except Exception as e:
                # Don't retry other types of errors
                raise Exception(f"Error processing PDF: {str(e)}")

        raise Exception("Failed to process PDF after exhausting all retries")

    def get_next_pdf(self, source_folder: str) -> Optional[str]:
        """Get the next PDF file from the source folder."""
        try:
            if not is_path_available(source_folder):
                raise Exception("Network path is not available")

            # Get list of PDF files
            try:
                original_timeout = getdefaulttimeout()
                setdefaulttimeout(self._network_timeout)
                try:
                    pdf_files = sorted(
                        [
                            f
                            for f in os_listdir(source_folder)
                            if f.lower().endswith(".pdf")
                        ]
                    )
                    self.current_file_list = pdf_files
                finally:
                    setdefaulttimeout(original_timeout)
            except Exception as e:
                if isinstance(e, SocketTimeout):
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
                next_pdf = path.join(
                    source_folder, self.current_file_list[self.current_file_index]
                )
                # Clear cache if different file
                if next_pdf != self.cached_pdf_path:
                    self.clear_cache()
                    self.current_rotation = 0  # Reset rotation when moving to a new file
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

    def rotate_page(self, clockwise=True):
        """Rotate the current PDF page clockwise or counterclockwise by 90 degrees."""
        if clockwise:
            self.current_rotation = (self.current_rotation + 90) % 360
        else:
            self.current_rotation = (self.current_rotation - 90) % 360

    def get_rotation(self):
        """Get the current rotation angle."""
        return self.current_rotation

    def render_pdf_page(
        self, pdf_path: str, page_num: int = 0, zoom: float = 1.0
    ) -> Any:
        """Render a PDF page as a PhotoImage."""
        try:
            if not is_path_available(pdf_path):
                raise Exception("Network path is not available")

            # Use cached document or open new one
            if pdf_path != self.cached_pdf_path:
                self.clear_cache()
                original_timeout = getdefaulttimeout()
                setdefaulttimeout(self._network_timeout)
                try:
                    self.cached_pdf = fitz_open(pdf_path)
                    self.cached_pdf_path = pdf_path
                finally:
                    setdefaulttimeout(original_timeout)

            # Get the specified page
            page = self.cached_pdf[page_num]

            # Calculate matrix with optimized settings for scanned PDFs
            base_dpi = 72  # Base DPI for PDF
            target_dpi = 144  # Target DPI for better quality
            dpi_scale = target_dpi / base_dpi
            
            # Calculate quality multiplier based on zoom level
            # Smoothly increase quality as zoom increases
            if zoom <= 1.0:
                quality_multiplier = 1.0
            else:
                # Gradually increase quality up to 1.5x for zooms up to 2.0
                # Beyond 2.0x zoom, maintain 1.5x quality
                quality_multiplier = min(1.0 + (zoom - 1.0) * 0.5, 1.5)
            
            # Final zoom calculation incorporating DPI scaling and quality multiplier
            effective_zoom = zoom * (dpi_scale * quality_multiplier)

            # Create matrix with optimal settings for scanned documents and rotation
            zoom_matrix = Matrix(effective_zoom, effective_zoom)
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
            image = pil_open(BytesIO(img_data))

            return image

        except Exception as e:
            self.clear_cache()  # Clear cache on error
            if isinstance(e, SocketTimeout):
                raise Exception("Network timeout while accessing PDF file")
            raise Exception(f"Error rendering PDF: {str(e)}")

    def revert_pdf_location(
        self,
        task: PDFTask
    ) -> None:
        """Revert PDF file back to its original location.
        
        Args:
            task (PDFTask): The task object containing the current and original PDF locations and processed_pdf_location
            
        Raises:
            Exception: If reverting the PDF location fails.
        """
        try:
            if not task.original_pdf_location:
                raise Exception("No original PDF location stored in task")

            current_pdf_path = task.processed_pdf_location

            print(f"[DEBUG] Attempting to revert PDF location from '{current_pdf_path}' to '{task.original_pdf_location}'")
            
            # First, ensure the original directory exists
            try:
                original_dir = path.dirname(task.original_pdf_location)
                if not path.exists(original_dir):
                    print(f"[DEBUG] Creating original directory '{original_dir}'")
                    makedirs(original_dir, exist_ok=True)
            except Exception as e:
                print(f"[DEBUG] Failed to create original directory '{original_dir}': {str(e)}")
                raise
            
            # Move the file back to its original location
            try:
                if not path.exists(current_pdf_path):
                    raise Exception(f"Processed PDF file not found at '{current_pdf_path}'")

                # If original file exists, remove it first
                if path.exists(task.original_pdf_location):
                    remove(task.original_pdf_location)
                    print(f"[DEBUG] Removed existing file at original location '{task.original_pdf_location}'")

                copy2(current_pdf_path, task.original_pdf_location)
                print(f"[DEBUG] PDF file moved back to original location '{task.original_pdf_location}' successfully")
                
                # Remove the file from the processed folder
                remove(current_pdf_path)
                print(f"[DEBUG] Removed PDF file from processed location '{current_pdf_path}'")
                
                # Clear versioned filename from task
                task.versioned_filename = None
                
            except Exception as e:
                print(f"[DEBUG] Error reverting PDF location: {str(e)}")
                raise Exception(f"Failed to revert PDF location: {str(e)}")
                
        except Exception as e:
            print(f"[DEBUG] Error in revert_pdf_location: {str(e)}")
            raise

    def close_current_pdf(self) -> None:
        """Close any open PDF files to release system resources."""
        try:
            if hasattr(self, '_current_pdf') and self._current_pdf is not None:
                self._current_pdf.close()
                self._current_pdf = None
        except Exception as e:
            print(f"[DEBUG] Error closing PDF: {str(e)}")
