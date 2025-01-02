from pandas import read_excel, ExcelFile, DataFrame, Series, to_datetime
from pandas.api.types import is_datetime64_any_dtype
from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font
from shutil import copy2
from os import path
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
            bool: True if data was loaded successfully

        Raises:
            Exception: If network path is unavailable or Excel file cannot be loaded
        """
        try:
            # Check network path availability first
            if not is_path_available(excel_file):
                raise Exception("Network path is not available")

            # Check if we need to reload
            try:
                current_modified = path.getmtime(excel_file)
            except (OSError, PermissionError) as e:
                print(f"[DEBUG] Failed to get file modification time: {str(e)}")
                current_modified = None  # Handle network/permission errors for file stat
                # Force reload since we can't verify modification time
                self.excel_data = None
                self._cached_file = None
                self._cached_sheet = None
                self._last_modified = None
                self._hyperlink_cache = {}

            if (
                self.excel_data is not None
                and self._cached_file == excel_file
                and self._cached_sheet == sheet_name
                and self._last_modified == current_modified
            ):
                return True  # Use cached data

            # Load new data with timeout
            original_timeout = getdefaulttimeout()
            setdefaulttimeout(self._network_timeout)
            try:
                self.excel_data = read_excel(excel_file, sheet_name=sheet_name)

                # Clear existing cache
                self._hyperlink_cache = {}

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

    def cache_hyperlinks_for_column(
        self, excel_file: str, sheet_name: str, column_name: str
    ) -> None:
        """Cache hyperlink information for a specific column.

        Args:
            excel_file: Path to the Excel file
            sheet_name: Name of the sheet
            column_name: Name of the column to check for hyperlinks
        """
        if not self.excel_data is not None:
            return

        print(f"[DEBUG] Loading hyperlink information for column {column_name}")
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

        wb.close()
        print(f"[DEBUG] Hyperlink information cached for column {column_name}")

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

            # Set hyperlink and style the cell to look like a hyperlink
            cell.hyperlink = Hyperlink(
                ref=cell.coordinate,
                target=relative_pdf_path,
            )
            cell.font = Font(underline="single", color="0000FF")

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
        filter1_col: str,
        filter2_col: str,
        filter3_col: str,
        value1: str,
        value2: str,
        value3: str,
    ) -> Tuple[Optional[Series], Optional[int]]:
        """Find row matching the filter values."""
        if self.excel_data is None:
            return None, None

        # Convert all columns to string and strip whitespace
        df = self.excel_data.copy()

        # Handle datetime columns specially
        for col, val in [
            (filter1_col, value1),
            (filter2_col, value2),
            (filter3_col, value3),
        ]:
            if is_datetime64_any_dtype(df[col]):
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
        print(
            f"[DEBUG] Column types: {df[filter1_col].dtype} | {df[filter2_col].dtype} | {df[filter3_col].dtype}"
        )

        # Create the mask for each condition and combine them
        def create_mask(col: str, value: str) -> Series:
            if is_datetime64_any_dtype(df[col]):
                try:
                    # Try parsing the value as datetime for datetime columns
                    parsed_date = to_datetime(value)
                    return df[col].dt.date == parsed_date.date()
                except Exception as e:
                    print(f"[DEBUG] Failed to parse date for {col}: {str(e)}")
                    return Series(False, index=df.index)
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
            matching_rows = df[
                mask1 & mask2
            ]  # Show rows that match first two conditions
            if not matching_rows.empty:
                print(f"[DEBUG] Found rows matching first two conditions:")
                print(
                    f"[DEBUG] {matching_rows[[filter1_col, filter2_col, filter3_col]].to_string()}"
                )
                print(
                    f"[DEBUG] Available values in filtered rows: {matching_rows[filter3_col].unique().tolist()}"
                )
                # Add detailed comparison for the third column
                for idx, row in matching_rows.iterrows():
                    actual_value = row[filter3_col]
                    print(f"[DEBUG] Comparing FA values:")
                    print(f"[DEBUG] Expected (len={len(value3)}): {repr(value3)}")
                    print(
                        f"[DEBUG] Actual (len={len(actual_value)}): {repr(actual_value)}"
                    )
                    print(f"[DEBUG] Values equal: {actual_value == value3}")
                    if actual_value != value3:
                        # Compare character by character
                        min_len = min(len(actual_value), len(value3))
                        for i in range(min_len):
                            if actual_value[i] != value3[i]:
                                print(f"[DEBUG] First difference at position {i}:")
                                print(
                                    f"[DEBUG] Expected char: {repr(value3[i])} (ord={ord(value3[i])})"
                                )
                                print(
                                    f"[DEBUG] Actual char: {repr(actual_value[i])} (ord={ord(actual_value[i])})"
                                )
                                break
            else:
                print("[DEBUG] No rows match even the first two conditions")
                print(
                    f"[DEBUG] Values in first column ({filter1_col}): {df[filter1_col].unique().tolist()[:5]}..."
                )
                print(
                    f"[DEBUG] Values in second column ({filter2_col}): {df[filter2_col].unique().tolist()[:5]}..."
                )
            return None, None

        return df[mask].iloc[0], mask.idxmax()

    def has_hyperlink(self, row_idx: int) -> bool:
        """Check if a row has any hyperlinks.

        Args:
            row_idx: The 0-based row index to check

        Returns:
            bool: True if the row has any hyperlinks, False otherwise
        """
        return self._hyperlink_cache.get(row_idx, False)
