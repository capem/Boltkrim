import pandas as pd
import openpyxl
import os

class ExcelManager:
    def __init__(self):
        self.excel_data = None
        
    def load_excel_data(self, excel_file, sheet_name):
        """Load data from Excel file."""
        try:
            self.excel_data = pd.read_excel(
                excel_file,
                sheet_name=sheet_name
            )
            return True
        except Exception as e:
            raise Exception(f"Error loading Excel data: {str(e)}")
            
    def get_sheet_names(self, excel_file):
        """Get list of sheet names from Excel file."""
        try:
            xl = pd.ExcelFile(excel_file)
            return xl.sheet_names
        except Exception as e:
            raise Exception(f"Error reading Excel sheets: {str(e)}")
            
    def get_column_names(self):
        """Get list of column names from loaded Excel data."""
        if self.excel_data is None:
            return []
        return self.excel_data.columns.tolist()
        
    def update_pdf_link(self, excel_file, sheet_name, row_idx, pdf_path):
        """Update Excel with PDF link."""
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb[sheet_name]
            
            # Get the last column
            last_col = ws.max_column
            
            # Add header for link column if it doesn't exist
            if ws.cell(row=1, column=last_col).value != "PDF Link":
                last_col += 1
                ws.cell(row=1, column=last_col, value="PDF Link")
            
            # Create relative path for Excel link
            rel_path = os.path.relpath(
                pdf_path,
                os.path.dirname(excel_file)
            )
            
            # Add hyperlink
            ws.cell(row=row_idx + 2, column=last_col).hyperlink = rel_path
            
            # Save Excel file
            wb.save(excel_file)
            wb.close()
            
        except Exception as e:
            raise Exception(f"Error updating Excel with PDF link: {str(e)}")
            
    def find_matching_row(self, filter1_col, filter2_col, value1, value2):
        """Find row matching the filter values."""
        if self.excel_data is None:
            return None
            
        mask = (self.excel_data[filter1_col] == value1) & \
               (self.excel_data[filter2_col] == value2)
               
        if not mask.any():
            return None
            
        return self.excel_data[mask].iloc[0], mask.idxmax()
