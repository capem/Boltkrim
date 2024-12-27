from pandas import read_excel, ExcelFile, DataFrame, Series
from win32com.client import Dispatch, pywintypes
from shutil import copy2
from os import path, remove
from socket import socket, AF_INET, SOCK_STREAM, getdefaulttimeout, setdefaulttimeout, timeout
from typing import Optional, List, Tuple
from time import sleep
from random import uniform
import pandas as pd

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
    def update_pdf_link(self, excel_file: str, sheet_name: str, row_idx: int, pdf_path: str, filter2_col: str, force: bool = False) -> Tuple[bool, bool]:
        """Update Excel with PDF link. 
        
        Args:
            excel_file: Path to the Excel file
            sheet_name: Name of the sheet to update
            row_idx: Row index to update (0-based)
            pdf_path: Path to the PDF file
            filter2_col: Column name for the hyperlink
            force: Whether to force update even if hyperlink exists
            
        Returns:
            Tuple[bool, bool]: (success, had_existing_link)
            - success: True if update was successful, False if file was locked
            - had_existing_link: True if there was an existing hyperlink
        """
        if not path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
            
        if not path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
            
        print(f"[DEBUG] Updating Excel link in {sheet_name}, row {row_idx + 2}, column {filter2_col}")
        
        wb = None
        backup_created = False
        
        try:
            # Initialize Excel application
            if self.excel_app is None:
                print("[DEBUG] Initializing Excel application")
                self.excel_app = Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.excel_app.DisplayAlerts = False
            
            try:
                # Open the workbook with read/write access
                print("[DEBUG] Opening workbook")
                wb = self.excel_app.Workbooks.Open(
                    excel_file,
                    UpdateLinks=False,
                    ReadOnly=not force  # Only open in write mode if we're going to update
                )
                print("[DEBUG] Getting worksheet")
                ws = wb.Worksheets(sheet_name)
                
                # Find the target column for the PDF link
                print("[DEBUG] Finding target column")
                target_col = None
                for col in range(1, ws.UsedRange.Columns.Count + 1):
                    header_value = ws.Cells(1, col).Value
                    print(f"[DEBUG] Column {col} header: {header_value}")
                    if header_value == filter2_col:
                        target_col = col
                        break
                
                if target_col is None:
                    raise Exception(f"Column {filter2_col} not found in Excel sheet")
                
                print(f"[DEBUG] Target column found at position {target_col}")
                
                # Get the cell and check for existing hyperlink
                print(f"[DEBUG] Accessing cell at row {row_idx + 2}, column {target_col}")
                cell = ws.Cells(row_idx + 2, target_col)
                original_value = cell.Value
                print(f"[DEBUG] Original cell value: {original_value}")
                
                # Check for existing hyperlink
                has_existing_link = False
                try:
                    has_existing_link = cell.Hyperlinks.Count > 0
                    if has_existing_link and not force:
                        # Close workbook without saving and return
                        wb.Close(SaveChanges=False)
                        wb = None
                        return True, True  # File accessible but has existing link
                except Exception as e:
                    print(f"[DEBUG] Error checking hyperlink: {str(e)}")
                
                if force or not has_existing_link:
                    # Create backup only if we're going to modify the file
                    backup_file = excel_file + '.bak'
                    print("[DEBUG] Creating backup file")
                    copy2(excel_file, backup_file)
                    backup_created = True
                    
                    # Create relative path for Excel link
                    rel_path = path.relpath(
                        pdf_path,  # Use the path directly since it's already sanitized by PDFManager
                        path.dirname(excel_file)
                    )
                    print(f"[DEBUG] Created relative path: {rel_path}")
                    
                    # Remove existing hyperlink if any
                    try:
                        if has_existing_link:
                            print("[DEBUG] Removing existing hyperlink")
                            cell.Hyperlinks.Delete()
                    except Exception as e:
                        print(f"[DEBUG] Error removing hyperlink: {str(e)}")
                    
                    # Add new hyperlink while preserving the cell value
                    print("[DEBUG] Adding new hyperlink")
                    ws.Hyperlinks.Add(
                        Anchor=cell,
                        Address=rel_path,
                        TextToDisplay=original_value or path.basename(pdf_path)
                    )
                    
                    # Save and close
                    print("[DEBUG] Saving workbook")
                    wb.Save()
                    print("[DEBUG] Closing workbook")
                    wb.Close(SaveChanges=True)
                    wb = None
                    
                    # Delete backup if everything succeeded
                    if backup_created and path.exists(backup_file):
                        print("[DEBUG] Removing backup file")
                        remove(backup_file)
                        
                    print("[DEBUG] Excel update completed successfully")
                    
                return True, has_existing_link
                    
            except pywintypes.com_error as e:
                print(f"[DEBUG] COM error: {str(e)}")
                # Handle specific COM errors
                if e.hresult == -2147352567:  # File is locked for editing
                    print("[DEBUG] Excel file is locked for editing")
                    return False, False  # Return False to indicate file was locked
                elif e.hresult == -2147417848:  # Excel automation server error
                    if self.excel_app:
                        print("[DEBUG] Excel automation error, quitting Excel")
                        self.excel_app.Quit()
                        self.excel_app = None
                raise
                    
        except Exception as e:
            print(f"[DEBUG] Error in update_pdf_link: {str(e)}")
            # Restore from backup if something went wrong
            if wb:
                try:
                    print("[DEBUG] Closing workbook without saving due to error")
                    wb.Close(SaveChanges=False)
                except:
                    pass
                    
            if backup_created and path.exists(backup_file):
                try:
                    print("[DEBUG] Restoring from backup")
                    copy2(backup_file, excel_file)
                except:
                    pass
            raise Exception(f"Error updating Excel with PDF link: {str(e)}")
            
        finally:
            # Clean up
            if wb:
                try:
                    print("[DEBUG] Cleanup: Closing workbook")
                    wb.Close(SaveChanges=False)
                except:
                    pass
                    
            if backup_created and path.exists(backup_file):
                try:
                    print("[DEBUG] Cleanup: Removing backup file")
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
            
        # Convert all columns to string and strip whitespace
        df = self.excel_data.copy()
        
        # Handle datetime columns specially
        for col, val in [(filter1_col, value1), (filter2_col, value2), (filter3_col, value3)]:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                print(f"[DEBUG] Column {col} is datetime type")
                # Keep datetime type for the column
                continue
            else:
                # Convert to string and strip for non-datetime columns
                df[col] = df[col].astype(str).str.strip()
        
        # Convert and strip input values, except for datetime columns
        value1 = str(value1).strip()
        value2 = str(value2).strip()
        value3 = str(value3).strip()
        
        print(f"[DEBUG] Looking for combination: {value1} | {value2} | {value3}")
        print(f"[DEBUG] Value3 length: {len(value3)}")
        print(f"[DEBUG] Value3 repr: {repr(value3)}")
        print(f"[DEBUG] In columns: {filter1_col} | {filter2_col} | {filter3_col}")
        print(f"[DEBUG] Column types: {df[filter1_col].dtype} | {df[filter2_col].dtype} | {df[filter3_col].dtype}")
        
        # Create the mask for each condition and combine them
        def create_mask(col: str, value: str) -> pd.Series:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                try:
                    # Try parsing the value as datetime for datetime columns
                    parsed_date = pd.to_datetime(value)
                    return df[col].dt.date == parsed_date.date()
                except Exception as e:
                    print(f"[DEBUG] Failed to parse date for {col}: {str(e)}")
                    return pd.Series(False, index=df.index)
            else:
                return df[col] == value
        
        mask1 = create_mask(filter1_col, value1)
        mask2 = create_mask(filter2_col, value2)
        mask3 = create_mask(filter3_col, value3)
        
        # Debug output for matching conditions
        print(f"[DEBUG] Rows matching first condition: {mask1.sum()}")
        print(f"[DEBUG] Rows matching second condition: {mask2.sum()}")
        print(f"[DEBUG] Rows matching third condition: {mask3.sum()}")
        
        mask = mask1 & mask2 & mask3
        
        if not mask.any():
            # Debug output to help identify the issue
            matching_rows = df[mask1 & mask2]  # Show rows that match first two conditions
            if not matching_rows.empty:
                print(f"[DEBUG] Found rows matching first two conditions:")
                print(f"[DEBUG] {matching_rows[[filter1_col, filter2_col, filter3_col]].to_string()}")
                print(f"[DEBUG] Available values in filtered rows: {matching_rows[filter3_col].unique().tolist()}")
                # Add detailed comparison for the third column
                for idx, row in matching_rows.iterrows():
                    actual_value = row[filter3_col]
                    print(f"[DEBUG] Comparing FA values:")
                    print(f"[DEBUG] Expected (len={len(value3)}): {repr(value3)}")
                    print(f"[DEBUG] Actual (len={len(actual_value)}): {repr(actual_value)}")
                    print(f"[DEBUG] Values equal: {actual_value == value3}")
                    if actual_value != value3:
                        # Compare character by character
                        min_len = min(len(actual_value), len(value3))
                        for i in range(min_len):
                            if actual_value[i] != value3[i]:
                                print(f"[DEBUG] First difference at position {i}:")
                                print(f"[DEBUG] Expected char: {repr(value3[i])} (ord={ord(value3[i])})")
                                print(f"[DEBUG] Actual char: {repr(actual_value[i])} (ord={ord(actual_value[i])})")
                                break
            else:
                print("[DEBUG] No rows match even the first two conditions")
                print(f"[DEBUG] Values in first column ({filter1_col}): {df[filter1_col].unique().tolist()[:5]}...")
                print(f"[DEBUG] Values in second column ({filter2_col}): {df[filter2_col].unique().tolist()[:5]}...")
            return None, None
            
        return df[mask].iloc[0], mask.idxmax()

    def __del__(self) -> None:
        """Cleanup Excel application on object destruction"""
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass
