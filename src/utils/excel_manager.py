from pandas import read_excel, ExcelFile, DataFrame, Series
from win32com.client import Dispatch, pywintypes
from shutil import copy2
from os import path, remove
from socket import socket, AF_INET, SOCK_STREAM, getdefaulttimeout, setdefaulttimeout, timeout
from typing import Optional, List, Tuple
from time import sleep
from random import uniform

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

def retry_with_backoff(func, max_attempts: int = 5, initial_delay: float = 1.0):
    """Decorator to retry a function with exponential backoff.
    
    Args:
        func: Function to retry
        max_attempts: Maximum number of retry attempts
        initial_delay: Initial delay between retries in seconds
    """
    def wrapper(*args, **kwargs):
        delay = initial_delay
        last_exception = None
        
        for attempt in range(max_attempts):
            try:
                return func(*args, **kwargs)
            except (pywintypes.com_error, IOError, PermissionError) as e:
                last_exception = e
                if attempt == max_attempts - 1:
                    raise
                
                # Add jitter to avoid thundering herd
                sleep_time = delay + uniform(0, 0.1 * delay)
                sleep(sleep_time)
                delay *= 2
                
        raise last_exception
    return wrapper

class ExcelManager:
    def __init__(self):
        self.excel_data: Optional[DataFrame] = None
        self.excel_app = None
        self._cached_file = None
        self._cached_sheet = None
        self._last_modified = None
        self._network_timeout = 5  # 5 seconds timeout for network operations
        
    @retry_with_backoff
    def load_excel_data(self, excel_file: str, sheet_name: str) -> bool:
        """Load data from Excel file with caching and retry logic."""
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
    
    @retry_with_backoff
    def update_pdf_link(self, excel_file: str, sheet_name: str, row_idx: int, pdf_path: str) -> bool:
        """Update Excel with PDF link. Returns True if update was successful, False if file was locked."""
        if not path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
            
        if not path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
            
        # Create backup of the Excel file
        backup_file = excel_file + '.bak'
        wb = None
        
        try:
            copy2(excel_file, backup_file)
            
            # Initialize Excel application
            if self.excel_app is None:
                self.excel_app = Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.excel_app.DisplayAlerts = False
            
            try:
                # Open the workbook with read/write access
                wb = self.excel_app.Workbooks.Open(
                    excel_file,
                    UpdateLinks=False,
                    ReadOnly=False
                )
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
                    TextToDisplay=original_value or path.basename(pdf_path)
                )
                
                # Save and close
                wb.Save()
                wb.Close(SaveChanges=True)
                wb = None
                
                # Delete backup if everything succeeded
                if path.exists(backup_file):
                    remove(backup_file)
                    
                return True
                    
            except pywintypes.com_error as e:
                # Handle specific COM errors
                if e.hresult == -2147352567:  # File is locked for editing
                    return False  # Return False to indicate file was locked
                elif e.hresult == -2147417848:  # Excel automation server error
                    if self.excel_app:
                        self.excel_app.Quit()
                        self.excel_app = None
                raise
                    
        except Exception as e:
            # Restore from backup if something went wrong
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
                    
            if path.exists(backup_file):
                try:
                    copy2(backup_file, excel_file)
                except:
                    pass
            raise Exception(f"Error updating Excel with PDF link: {str(e)}")
            
        finally:
            # Clean up
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
                    
            if path.exists(backup_file):
                try:
                    remove(backup_file)
                except:
                    pass
            
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
