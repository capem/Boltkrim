import os
import shutil
import tempfile
import fitz  # PyMuPDF
from PIL import Image
import io

class PDFManager:
    @staticmethod
    def get_next_pdf(source_folder):
        """Get the next PDF file from the source folder."""
        if not os.path.exists(source_folder):
            return None
            
        for file in os.listdir(source_folder):
            if file.lower().endswith('.pdf'):
                return os.path.join(source_folder, file)
        return None
        
    @staticmethod
    def process_pdf(current_pdf, new_filepath, processed_folder):
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
                
    @staticmethod
    def render_pdf_page(pdf_path, page_num=0, zoom=1.0):
        """Render a PDF page as a PhotoImage."""
        try:
            # Open the PDF file
            pdf_document = fitz.open(pdf_path)
            
            # Get the specified page
            page = pdf_document[page_num]
            
            # Get the page's dimensions
            zoom_matrix = fitz.Matrix(zoom, zoom)
            
            # Get page as a PNG image
            pix = page.get_pixmap(matrix=zoom_matrix)
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))
            
            return image
            
        except Exception as e:
            raise Exception(f"Error rendering PDF: {str(e)}")
        finally:
            if 'pdf_document' in locals():
                pdf_document.close()
