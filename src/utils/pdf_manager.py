import os
import shutil
import tempfile
import fitz  # PyMuPDF
from PIL import Image
import io

class PDFManager:
    def __init__(self):
        self.current_file_index = -1
        self.current_file_list = []
        self.cached_pdf = None
        self.cached_pdf_path = None
        
    def get_next_pdf(self, source_folder):
        """Get the next PDF file from the source folder."""
        if not os.path.exists(source_folder):
            return None
            
        # Get list of PDF files if not already loaded or if folder changed
        pdf_files = sorted([
            f for f in os.listdir(source_folder)
            if f.lower().endswith('.pdf')
        ])
        
        # If no PDF files found
        if not pdf_files:
            return None
            
        # Move to next file
        self.current_file_index += 1
        
        # If we've reached the end, start over
        if self.current_file_index >= len(pdf_files):
            self.current_file_index = 0
            
        # Return full path of next PDF
        if pdf_files:
            next_pdf = os.path.join(source_folder, pdf_files[self.current_file_index])
            # Clear cache if different file
            if next_pdf != self.cached_pdf_path:
                self.clear_cache()
            return next_pdf
            
        return None
        
    def clear_cache(self):
        """Clear the cached PDF document."""
        if self.cached_pdf:
            self.cached_pdf.close()
            self.cached_pdf = None
            self.cached_pdf_path = None
                
    def process_pdf(self, current_pdf, new_filepath, processed_folder):
        """Process a PDF file - move it to the processed folder with a new name."""
        if not os.path.exists(current_pdf):
            raise Exception("Source PDF file not found")
            
        if not os.path.exists(processed_folder):
            os.makedirs(processed_folder)
            
        # Create temporary directory for atomic operations
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create backup copy in temp directory
            temp_pdf = os.path.join(temp_dir, "original.pdf")
            shutil.copy2(current_pdf, temp_pdf)
            
            try:
                # Move and rename the PDF file
                os.replace(current_pdf, new_filepath)
                return True
            except Exception as e:
                # Restore file from backup if operation fails
                if os.path.exists(new_filepath):
                    os.remove(new_filepath)
                shutil.copy2(temp_pdf, current_pdf)
                raise Exception(f"Error processing PDF: {str(e)}")
                
    def render_pdf_page(self, pdf_path, page_num=0, zoom=1.0):
        """Render a PDF page as a PhotoImage."""
        try:
            # Use cached document or open new one
            if pdf_path != self.cached_pdf_path:
                self.clear_cache()
                self.cached_pdf = fitz.open(pdf_path)
                self.cached_pdf_path = pdf_path
            
            # Get the specified page
            page = self.cached_pdf[page_num]
            
            # Calculate matrix with optimized settings for scanned PDFs
            base_dpi = 72  # Base DPI for PDF
            quality_multiplier = 2 if zoom > 1.0 else 1  # Higher quality when zoomed in
            
            # Create matrix with optimal settings for scanned documents
            zoom_matrix = fitz.Matrix(zoom * quality_multiplier, zoom * quality_multiplier)
            
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
            raise Exception(f"Error rendering PDF: {str(e)}")
        finally:
            if 'pdf_document' in locals():
                pdf_document.close()
