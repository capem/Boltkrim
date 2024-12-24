from pandas import read_excel, ExcelFile, DataFrame, Series
from win32com.client import Dispatch
from shutil import copy2
from os import path, remove
from socket import socket, AF_INET, SOCK_STREAM, getdefaulttimeout, setdefaulttimeout, timeout
from typing import Optional, List, Tuple

def is_path_available(filepath: str, timeout: int = 2) -> bool:
    """Check if a network path is available with timeout.
    
    Args:
        filepath: The path to check
        timeout: Connection timeout in seconds
        
    Returns:
        bool: True if path is available, False otherwise
    """
    if not filepath.startswith('\\\\'): # Not a network path
        return path.exists(filepath)
        
    try:
        # Extract server name from UNC path
        server = filepath.split('\\')[2]
        # Try to connect to the server
        sock = socket(AF_INET, SOCK_STREAM)
        sock.settimeout(timeout)
        sock.connect((server, 445))  # 445 is the SMB port
        sock.close()
        return True
    except:
        return False

class ExcelManager:
    def __init__(self):
        self.excel_data: Optional[DataFrame] = None
        self.excel_app = None
        self._cached_file = None
        self._cached_sheet = None
        self._last_modified = None
        self._network_timeout = 5  # 5 seconds timeout for network operations
        
    def load_excel_data(self, excel_file: str, sheet_name: str) -> bool:
        """Load data from Excel file with caching."""
        try:
            # Check network path availability first
            if not is_path_available(excel_file):
                raise Exception("Network path is not available")
                
            # Check if we need to reload
            try:
                current_modified = path.getmtime(excel_file) if path.exists(excel_file) else None
            except OSError:
                current_modified = None  # Handle network errors for file stat
            
            if (self.excel_data is not None and 
                self._cached_file == excel_file and 
                self._cached_sheet == sheet_name and 
                self._last_modified == current_modified):
                return True  # Use cached data
                
            # Load new data with timeout
            original_timeout = getdefaulttimeout()
            setdefaulttimeout(self._network_timeout)
            try:
                self.excel_data = read_excel(
                    excel_file,
                    sheet_name=sheet_name
                )
            finally:
                setdefaulttimeout(original_timeout)
            
            # Update cache info
            self._cached_file = excel_file
            self._cached_sheet = sheet_name
            self._last_modified = current_modified
            
            return True
        except Exception as e:
            if isinstance(e, timeout):
                raise Exception("Network timeout while accessing Excel file")
            raise Exception(f"Error loading Excel data: {str(e)}")
            
    def get_sheet_names(self, excel_file: str) -> List[str]:
        """Get list of sheet names from Excel file."""
        try:
            if not is_path_available(excel_file):
                raise Exception("Network path is not available")
                
            original_timeout = getdefaulttimeout()
            setdefaulttimeout(self._network_timeout)
            try:
                xl = ExcelFile(excel_file)
                return xl.sheet_names
            finally:
                setdefaulttimeout(original_timeout)
        except Exception as e:
            if isinstance(e, timeout):
                raise Exception("Network timeout while accessing Excel file")
            raise Exception(f"Error reading Excel sheets: {str(e)}")
            
    def get_column_names(self) -> List[str]:
        """Get list of column names from loaded Excel data."""
        if self.excel_data is None:
            return []
        return list(self.excel_data.columns)
        
    def update_pdf_link(self, excel_file: str, sheet_name: str, row_idx: int, pdf_path: str) -> None:
        """Update Excel with PDF link."""
        try:
            # Create backup of the Excel file
            backup_file = excel_file + '.bak'
            copy2(excel_file, backup_file)
            
            try:
                # Initialize Excel application
                if self.excel_app is None:
                    self.excel_app = Dispatch("Excel.Application")
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
                rel_path = path.relpath(
                    pdf_path,
                    path.dirname(excel_file)
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
                if path.exists(backup_file):
                    remove(backup_file)
                    
            except Exception as e:
                # Restore from backup if something went wrong
                if path.exists(backup_file):
                    copy2(backup_file, excel_file)
                raise e
                
            finally:
                # Clean up
                if path.exists(backup_file):
                    try:
                        remove(backup_file)
                    except:
                        pass
                        
        except Exception as e:
            raise Exception(f"Error updating Excel with PDF link: {str(e)}")
            
    def find_matching_row(self, filter1_col: str, filter2_col: str, filter3_col: str, value1: str, value2: str, value3: str) -> Tuple[Optional[Series], Optional[int]]:
        """Find row matching the filter values."""
        if self.excel_data is None:
            return None, None
            
        mask = (self.excel_data[filter1_col] == value1) & \
               (self.excel_data[filter2_col] == value2) & \
               (self.excel_data[filter3_col] == value3)
               
        if not mask.any():
            return None, None
            
        return self.excel_data[mask].iloc[0], mask.idxmax()

    def __del__(self) -> None:
        """Cleanup Excel application on object destruction"""
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass
