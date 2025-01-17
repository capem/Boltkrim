from os import path, remove
from pandas import read_excel, ExcelFile, DataFrame, Series, to_datetime
from pandas.api.types import is_datetime64_any_dtype
from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font
from shutil import copy2
from socket import (
    socket,
    AF_INET,
    SOCK_STREAM,
    getdefaulttimeout,
    setdefaulttimeout,
    timeout,
)
from typing import Optional, List, Tuple, Union
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
    print(f"[DEBUG] Checking path availability for: {filepath}")

    if not filepath.startswith("\\\\"):  # Not a network path
        exists = path.exists(filepath)
        print(f"[DEBUG] Local path check result: {exists}")
        return exists

    try:
        # Extract server name from UNC path
        server = filepath.split("\\")[2]
        print(f"[DEBUG] Extracted server name: {server}")

        # Try to connect to the server
        print(f"[DEBUG] Attempting connection to {server}:445 with {timeout}s timeout")
        sock = socket(AF_INET, SOCK_STREAM)
        sock.settimeout(timeout)
        sock.connect((server, 445))  # 445 is the SMB port
        sock.close()
        print("[DEBUG] Successfully connected to server")
        return True
    except Exception as e:
        print(f"[DEBUG] Connection failed: {str(e)}")
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
            except (IOError, PermissionError) as e:
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
    """A class to manage Excel file operations with caching and network path handling.

    This class provides functionality to:
    - Load and cache Excel data
    - Update PDF hyperlinks in Excel files
    - Find matching rows based on filter criteria
    - Handle network paths with timeouts and retries
    - Manage Excel sheets and columns

    Attributes:
        excel_data (Optional[DataFrame]): Cached pandas DataFrame containing Excel data
        _cached_file (Optional[str]): Path to the currently cached Excel file
        _cached_sheet (Optional[str]): Name of the currently cached sheet
        _last_modified (Optional[float]): Last modification timestamp of cached file
        _network_timeout (int): Timeout in seconds for network operations
        _hyperlink_cache (Dict[int, bool]): Cache of row indices to hyperlink status
    """

    def __init__(self):
        self.excel_data: Optional[DataFrame] = None
        self._cached_file = None
        self._cached_sheet = None
        self._last_modified = None
        self._network_timeout = 5  # 5 seconds timeout for network operations
        self._hyperlink_cache = {}  # Cache for hyperlink status

    @retry_with_backoff
    def load_excel_data(self, excel_file: str, sheet_name: str) -> bool:
        """Load data from Excel file with caching and retry logic.

        This method loads Excel data while implementing:
        - Network path availability checking
        - Caching to avoid unnecessary reloads
        - Retry logic with exponential backoff
        - Timeout handling for network operations

        Args:
            excel_file: Path to the Excel file to load
            sheet_name: Name of the sheet to load

        Returns:
            bool: True if data was reloaded, False if using cached data

        Raises:
            Exception: If network path is unavailable or Excel file cannot be loaded
        """
        try:
            # Check if we need to reload by comparing file modification time
            try:
                current_modified = path.getmtime(excel_file)
            except (OSError, PermissionError) as e:
                print(f"[DEBUG] Failed to get file modification time: {str(e)}")
                current_modified = None

            # Use cached data if available and not modified
            if (
                self.excel_data is not None
                and self._cached_file == excel_file
                and self._cached_sheet == sheet_name
                and self._last_modified == current_modified
                and current_modified is not None
            ):
                print("[DEBUG] Using cached Excel data")
                return False  # Using cached data

            # Only check path availability if we need to reload
            if not is_path_available(excel_file):
                raise Exception("Network path is not available")

            # Load new data with timeout
            original_timeout = getdefaulttimeout()
            setdefaulttimeout(self._network_timeout)
            try:
                print("[DEBUG] Loading fresh Excel data")
                self.excel_data = read_excel(excel_file, sheet_name=sheet_name)
                # Clear existing cache
                self._hyperlink_cache = {}
                # Clear the last cached key when loading new data
                if hasattr(self, '_last_cached_key'):
                    delattr(self, '_last_cached_key')
            finally:
                setdefaulttimeout(original_timeout)

            # Update cache info
            self._cached_file = excel_file
            self._cached_sheet = sheet_name
            self._last_modified = current_modified

            return True  # Data was reloaded
        except Exception as e:
            if isinstance(e, timeout):
                raise Exception("Network timeout while accessing Excel file")
            raise Exception(f"Error loading Excel data: {str(e)}")

    def cache_hyperlinks_for_column(
        self, excel_file: str, sheet_name: str, column_name: str
    ) -> None:
        """Cache hyperlink information for a specific column.

        Args:
            excel_file: Path to the Excel file
            sheet_name: Name of the sheet
            column_name: Name of the column to check for hyperlinks
        """
        # Skip if Excel data isn't loaded
        if self.excel_data is None:
            return

        # Skip if hyperlinks are already cached for this file/sheet/column combination
        cache_key = f"{excel_file}|{sheet_name}|{column_name}"
        if hasattr(self, '_last_cached_key') and self._last_cached_key == cache_key:
            print(f"[DEBUG] Hyperlinks already cached for {column_name}, skipping")
            return

        print(f"[DEBUG] Loading hyperlink information for column {column_name}")
        try:
            wb = load_workbook(excel_file, data_only=True)
            ws = wb[sheet_name]
            df_cols = read_excel(excel_file, sheet_name=sheet_name, nrows=0)

            # Find the target column
            if column_name not in df_cols.columns:
                wb.close()
                return

            target_col = df_cols.columns.get_loc(column_name) + 1

            # Clear existing cache
            self._hyperlink_cache = {}

            # Cache hyperlink information only for the target column
            for idx in range(len(self.excel_data)):
                cell = ws.cell(
                    row=idx + 2, column=target_col
                )  # +2 for header and 1-based indexing
                self._hyperlink_cache[idx] = cell.hyperlink is not None

            # Store the cache key
            self._last_cached_key = cache_key

            wb.close()
            print(f"[DEBUG] Hyperlink information cached for column {column_name}")
        except Exception as e:
            print(f"[DEBUG] Error caching hyperlink information: {str(e)}")

    @retry_with_backoff
    def update_pdf_link(
        self,
        excel_file: str,
        sheet_name: str,
        row_idx: int,
        pdf_path: str,
        filter2_col: str,
    ) -> Optional[str]:
        """Update Excel with PDF link and capture the original hyperlink.

        Args:
            excel_file: Path to the Excel file
            sheet_name: Name of the sheet to update
            row_idx: Row index to update (0-based)
            pdf_path: Path to the PDF file
            filter2_col: Column name for the hyperlink

        Returns:
            Optional[str]: The original hyperlink if it existed, else None
        """
        if not is_path_available(excel_file):
            raise FileNotFoundError(
                f"Excel file not found or not accessible: {excel_file}"
            )

        if not is_path_available(pdf_path):
            raise FileNotFoundError(f"PDF file not found or not accessible: {pdf_path}")

        print(
            f"[DEBUG] Updating Excel link in {sheet_name}, row {row_idx + 2}, column {filter2_col}"
        )

        wb = None
        backup_created = False
        original_hyperlink = None

        try:
            # Load workbook with data_only=False to preserve all formulas in the workbook
            wb = load_workbook(excel_file, data_only=False)
            ws = wb[sheet_name]

            # Identify the column index for filter2_col
            header = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
            if filter2_col not in header:
                raise ValueError(
                    f"Column '{filter2_col}' not found in sheet '{sheet_name}'"
                )
            col_idx = header[filter2_col]

            # Identify the cell to update
            cell = ws.cell(row=row_idx + 2, column=col_idx)  # +2: 1 for header, 1-based

            # Capture the original hyperlink
            if cell.hyperlink:
                original_hyperlink = cell.hyperlink.target

            # Create a backup copy
            if not backup_created:
                backup_file = f"{excel_file}.bak"
                copy2(excel_file, backup_file)
                backup_created = True
                print(f"[DEBUG] Backup created at {backup_file}")

            # Set the new hyperlink
            relative_pdf_path = path.relpath(pdf_path, path.dirname(excel_file))

            # Add new hyperlink while preserving the cell value
            print("[DEBUG] Adding new hyperlink")

            # Create a proper hyperlink with all required properties
            hyperlink = Hyperlink(
                ref=cell.coordinate,
                target=relative_pdf_path,
                display=cell.value,  # Keep existing cell value as display text
            )
            cell.hyperlink = hyperlink

            # Apply Excel's built-in Hyperlink style which includes visited state handling
            cell.style = 'Hyperlink'

            # Save the workbook
            print("[DEBUG] Saving workbook")
            wb.save(excel_file)
            print(
                f"[DEBUG] Excel file updated with new hyperlink at row {row_idx + 2}, column {filter2_col}"
            )

            return original_hyperlink

        except Exception as e:
            print(f"[DEBUG] Error in update_pdf_link: {str(e)}")
            # Restore from backup if something went wrong
            if wb:
                try:
                    print("[DEBUG] Closing workbook without saving due to error")
                    wb.close()
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
                wb.close()

    @retry_with_backoff
    def revert_pdf_link(
        self,
        excel_file: str,
        sheet_name: str,
        row_idx: int,
        filter2_col: str,
        original_hyperlink: str,
        original_value: str,
    ) -> None:
        """Revert the Excel cell's hyperlink and value to its original state."""
        try:
            print(f"[DEBUG] Attempting to revert Excel cell at row {row_idx}, column {filter2_col}")
            
            # Load workbook with data_only=False to preserve all formulas in the workbook
            wb = load_workbook(excel_file, data_only=False)
            ws = wb[sheet_name]
            
            # Map headers to column indices (1-based)
            header = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
            if filter2_col not in header:
                raise ValueError(f"Column '{filter2_col}' not found in sheet '{sheet_name}'")
            col_idx = header[filter2_col]
            
            # Get the cell to update
            cell = ws.cell(row=row_idx + 2, column=col_idx)  # +2 for header and 1-based indexing
            
            # Clear existing hyperlink
            if cell.hyperlink:
                cell.hyperlink = None
            
            # Reset the cell value
            cell.value = original_value
            
            if original_hyperlink:
                # If there was an original hyperlink, restore it and keep hyperlink style
                cell.hyperlink = Hyperlink(
                    ref=cell.coordinate,
                    target=original_hyperlink
                )
                cell.font = Font(underline="single", color="0000FF")
                print(f"[DEBUG] Restored hyperlink and style: {original_hyperlink}")
            else:
                # If there was no original hyperlink, clear the hyperlink style
                cell.font = Font()
                print("[DEBUG] Cleared hyperlink style")
            
            # Save the workbook
            wb.save(excel_file)
            print(f"[DEBUG] Successfully reverted Excel cell")
            
        except Exception as e:
            print(f"[DEBUG] Error reverting Excel cell: {str(e)}")
            raise Exception(f"Failed to revert Excel cell: {str(e)}")
        finally:
            if wb:
                wb.close()

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

    def find_matching_row(
        self,
        filter_columns: List[str],
        filter_values: List[str],
    ) -> Tuple[Optional[Series], Optional[int]]:
        """Find a row matching the given filter values.

        Args:
            filter_columns: List of column names to filter on
            filter_values: List of values to match in corresponding columns

        Returns:
            Tuple[Optional[Series], Optional[int]]: Matching row data and index if found
        """
        try:
            if self.excel_data is None:
                raise Exception("Excel data not loaded")

            if len(filter_columns) != len(filter_values):
                raise Exception(f"Number of filter columns ({len(filter_columns)}) and values ({len(filter_values)}) must match")

            print(f"[DEBUG] Finding row with filters: {dict(zip(filter_columns, filter_values))}")

            df = self.excel_data

            def create_mask(col: str, value: str) -> Series:
                """Create a boolean mask for filtering DataFrame."""
                if col not in df.columns:
                    raise Exception(f"Column '{col}' not found in Excel file. Available columns: {', '.join(df.columns)}")

                print(f"[DEBUG] Creating mask for column '{col}' with value '{value}'")

                # Convert column to string for comparison
                df[col] = df[col].astype(str)
                
                # Handle date columns
                if "DATE" in col.upper():
                    try:
                        # Try to parse the date value
                        date_value = pd.to_datetime(value)
                        print(f"[DEBUG] Parsed date value: {date_value}")
                        # Convert column to datetime if not already
                        if not is_datetime64_any_dtype(df[col]):
                            df[col] = pd.to_datetime(df[col], errors='coerce')
                        return df[col] == date_value
                    except:
                        # If date parsing fails, fall back to string comparison
                        print(f"[DEBUG] Date parsing failed, using string comparison for '{col}'")
                        return df[col].str.strip() == str(value).strip()
                else:
                    return df[col].str.strip() == str(value).strip()

            # Create mask for each filter
            masks = [create_mask(col, val) for col, val in zip(filter_columns, filter_values)]
            
            # Combine all masks with AND operation
            final_mask = pd.Series([True] * len(df))
            for mask in masks:
                final_mask &= mask

            matching_rows = df[final_mask]
            print(f"[DEBUG] Found {len(matching_rows)} matching rows")

            if len(matching_rows) == 0:
                print("[DEBUG] No matching rows found")
                return None, None
            elif len(matching_rows) == 1:
                row_idx = matching_rows.index[0]
                print(f"[DEBUG] Found exactly one match at row index {row_idx}")
                return matching_rows.iloc[0], row_idx
            else:
                # If multiple matches, return the first one
                row_idx = matching_rows.index[0]
                print(f"[DEBUG] Found multiple matches, using first match at row index {row_idx}")
                return matching_rows.iloc[0], row_idx

        except Exception as e:
            print(f"[DEBUG] Error in find_matching_row: {str(e)}")
            raise Exception(f"Error finding matching row: {str(e)}")

    def has_hyperlink(self, row_idx: int) -> bool:
        """Check if a row has any hyperlinks.

        Args:
            row_idx: The 0-based row index to check

        Returns:
            bool: True if the row has any hyperlinks, False otherwise
        """
        return self._hyperlink_cache.get(row_idx, False)

    @retry_with_backoff
    def add_new_row(
        self,
        excel_file: str,
        sheet_name: str,
        filter_columns: List[str],
        filter_values: List[str],
    ) -> Tuple[Series, int]:
        """Add a new row to the Excel file with the given filter values.

        Args:
            excel_file: Path to the Excel file
            sheet_name: Name of the sheet to update
            filter_columns: List of column names to set
            filter_values: List of values to set in corresponding columns

        Returns:
            Tuple[Series, int]: The new row data and its index
        """
        try:
            if len(filter_columns) != len(filter_values):
                raise Exception(f"Number of filter columns ({len(filter_columns)}) and values ({len(filter_values)}) must match")

            print(f"[DEBUG] Adding new row with values: {dict(zip(filter_columns, filter_values))}")

            # Create a temporary backup
            backup_file = excel_file + ".bak"
            copy2(excel_file, backup_file)
            print(f"[DEBUG] Created backup at {backup_file}")

            try:
                # Load workbook
                wb = load_workbook(excel_file)
                ws = wb[sheet_name]

                # Get header row to map column names to indices
                header_row = ws[1]
                col_indices = {cell.value: idx + 1 for idx, cell in enumerate(header_row)}

                # Verify all filter columns exist
                for col in filter_columns:
                    if col not in col_indices:
                        raise Exception(f"Column '{col}' not found in Excel file. Available columns: {', '.join(col_indices.keys())}")

                # Find the last row with data
                last_row = ws.max_row
                new_row_idx = last_row + 1

                # First pass: Copy all formats from the template row
                template_row = last_row
                for col_idx in range(1, len(header_row) + 1):
                    template_cell = ws.cell(row=template_row, column=col_idx)
                    new_cell = ws.cell(row=new_row_idx, column=col_idx)
                    
                    # Copy number format and style
                    new_cell.number_format = template_cell.number_format
                    new_cell._style = template_cell._style

                # Second pass: Set values with proper type conversion
                for col, val in zip(filter_columns, filter_values):
                    col_idx = col_indices[col]
                    template_cell = ws.cell(row=template_row, column=col_idx)
                    new_cell = ws.cell(row=new_row_idx, column=col_idx)
                    
                    # Convert value based on the column type
                    if "DATE" in col.upper() and val:
                        try:
                            # Try to parse date in French format first (dd/mm/yyyy)
                            french_date_formats = ['%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d_%m_%Y']
                            date_val = None
                            
                            for fmt in french_date_formats:
                                try:
                                    date_val = pd.to_datetime(val, format=fmt)
                                    break
                                except:
                                    continue
                                    
                            if date_val is None:
                                # Fallback to pandas default parser
                                date_val = pd.to_datetime(val)
                                
                            new_cell.value = date_val.to_pydatetime()
                            # Set French date format
                            new_cell.number_format = 'DD/MM/YYYY'
                        except:
                            new_cell.value = val
                            print(f"[DEBUG] Could not parse date '{val}' for column '{col}'")
                            
                    elif "MNT" in col.upper() or any(num_indicator in col.upper() for num_indicator in ["MONTANT", "NOMBRE", "NUM", "PRIX"]):
                        try:
                            # Handle French number format (comma as decimal separator)
                            if isinstance(val, str):
                                # Replace comma with dot for conversion, handle thousands separator
                                num_str = val.replace(' ', '').replace(',', '.')
                                num_val = float(num_str)
                                new_cell.value = num_val
                            else:
                                new_cell.value = float(val)
                        except:
                            new_cell.value = val
                            print(f"[DEBUG] Could not parse number '{val}' for column '{col}'")
                    else:
                        new_cell.value = val
                    
                    print(f"[DEBUG] Set value '{val}' for column '{col}'")

                # Save the workbook
                wb.save(excel_file)

                # Update the cached DataFrame to stay in sync
                if self.excel_data is not None:
                    new_row_data = pd.Series(index=self.excel_data.columns)
                    for col, val in zip(filter_columns, filter_values):
                        new_row_data[col] = val
                    self.excel_data = pd.concat([self.excel_data, pd.DataFrame([new_row_data])], ignore_index=True)

                # Get the new row index for return value (0-based)
                row_idx = new_row_idx - 2  # Convert to 0-based index
                print(f"[DEBUG] Successfully added new row at index {row_idx}")

                # Remove backup if successful
                remove(backup_file)
                print("[DEBUG] Removed backup file after successful write")

                return new_row_data, row_idx

            except Exception as e:
                print(f"[DEBUG] Error while writing new row: {str(e)}")
                # Restore from backup
                copy2(backup_file, excel_file)
                remove(backup_file)
                print("[DEBUG] Restored from backup due to error")
                raise Exception(f"Failed to add new row: {str(e)}")

        except Exception as e:
            print(f"[DEBUG] Error in add_new_row: {str(e)}")
            raise Exception(f"Error adding new row: {str(e)}")
