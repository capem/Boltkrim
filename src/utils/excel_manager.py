import pandas as pd
import win32com.client
import shutil
import time
import os

class ExcelManager:
    def __init__(self):
        self.excel_data = None
        self.excel_app = None
        
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
            # Create backup of the Excel file
            backup_file = excel_file + '.bak'
            shutil.copy2(excel_file, backup_file)
            
            try:
                # Initialize Excel application
                if self.excel_app is None:
                    self.excel_app = win32com.client.Dispatch("Excel.Application")
                    self.excel_app.Visible = False
                    self.excel_app.DisplayAlerts = False
                
                # Open the workbook
                wb = self.excel_app.Workbooks.Open(excel_file)
                ws = wb.Worksheets(sheet_name)
                
                # Find the FACTURES column
                factures_col = None
                for col in range(1, ws.UsedRange.Columns.Count + 1):
                    if ws.Cells(1, col).Value == "FACTURES":
                        factures_col = col
                        break
                
                if factures_col is None:
                    raise Exception("FACTURES column not found in Excel sheet")
                
                # Create relative path for Excel link
                rel_path = os.path.relpath(
                    pdf_path,
                    os.path.dirname(excel_file)
                )
                
                # Get the cell and add hyperlink
                cell = ws.Cells(row_idx + 2, factures_col)
                original_value = cell.Value
                
                # Remove existing hyperlink if any
                if cell.Hyperlinks.Count > 0:
                    cell.Hyperlinks.Delete()
                
                # Add new hyperlink while preserving the cell value
                ws.Hyperlinks.Add(
                    Anchor=cell,
                    Address=rel_path,
                    TextToDisplay=original_value
                )
                
                # Save and close
                wb.Save()
                wb.Close()
                
                # Delete backup if everything succeeded
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                    
            except Exception as e:
                # Restore from backup if something went wrong
                if os.path.exists(backup_file):
                    shutil.copy2(backup_file, excel_file)
                raise e
                
            finally:
                # Clean up
                if os.path.exists(backup_file):
                    try:
                        os.remove(backup_file)
                    except:
                        pass
                        
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

    def __del__(self):
        """Cleanup Excel application on object destruction"""
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass
