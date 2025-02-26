from __future__ import annotations
from tkinter import (
    Event as TkEvent,
    Widget as TkWidget,
    filedialog,
)
from tkinter.ttk import (
    Frame,
    Label,
    Button,
    Style,
    LabelFrame,
)

from os import path, makedirs, remove
from shutil import copy2
from typing import Optional, Dict, List, Callable
from threading import Thread, Lock, Event
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog
from datetime import datetime
import pandas as pd
from ..utils import ConfigManager, ExcelManager, PDFManager, PDFTask
from .queue_display import QueueDisplay
from .pdf_viewer import PDFViewer
import traceback
import time


# Queue Management
class ProcessingQueue:
    def __init__(
        self,
        config_manager: ConfigManager,
        excel_manager: ExcelManager,
        pdf_manager: PDFManager,
    ):
        self.tasks: Dict[str, PDFTask] = {}
        self.lock = Lock()
        self.processing_thread: Optional[Thread] = None
        self.has_changes = False
        self.notification_lock = Lock()  # Separate lock for notifications
        try:
            self.stop_event = Event()
        except (AttributeError, RuntimeError):
            self.stop_event = Event()

        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        self._callbacks: List[Callable] = []

    def _parse_filter2_value(self, formatted_value: str) -> tuple[str, int]:
        """Parse filter2 value to get original value and row number.

        This implementation is used by the UI components to extract Excel row numbers
        from formatted filter values.

        Args:
            formatted_value: String in format "✓ value ⟨Excel Row: N⟩"

        Returns:
            tuple[str, int]: (original value without formatting, 0-based row index)
        """
        import re

        if not formatted_value:
            print("[DEBUG] UI received empty filter2 value")
            return "", -1

        # Remove checkmark if present
        formatted_value = formatted_value.replace("✓ ", "", 1)

        match = re.match(r"(.*?)\s*⟨Excel Row:\s*(\d+)⟩", formatted_value)
        if match:
            value = match.group(1).strip()
            row_num = int(match.group(2))
            print(
                f"[DEBUG] UI parsed filter2 value: '{formatted_value}' -> value='{value}', row={row_num - 2}"
            )
            return value, row_num - 2  # Convert back to 0-based index
        print(f"[DEBUG] UI failed to parse filter2 value: '{formatted_value}'")
        return formatted_value, -1

    def mark_changed(self) -> None:
        """Mark that the queue has changes that need to be displayed."""
        with self.lock:
            self.has_changes = True
        # Call notifications outside the lock to prevent deadlocks
        self._notify_status_change()

    def _notify_status_change(self) -> None:
        """Notify callbacks of status changes without holding the main lock."""
        with self.notification_lock:  # Use separate lock for notifications
            for callback in self._callbacks:
                try:
                    callback()
                except Exception as e:
                    print(f"[DEBUG] Callback error: {str(e)}")

    def update_task_status(self, task_id: str, new_status: str) -> None:
        """Update a task's status in a thread-safe way."""
        task_to_update = None
        with self.lock:
            # Find task by ID
            for task in self.tasks.values():
                if task.task_id == task_id:
                    task.status = new_status
                    # Set end time when task is completed or failed
                    if new_status in ["completed", "failed"]:
                        task.end_time = datetime.now()
                    task_to_update = task
                    self.has_changes = True
                    break

        # Notify outside the lock if we found and updated the task
        if task_to_update:
            self._notify_status_change()

    def get_task_by_id(self, task_id: str) -> Optional[PDFTask]:
        """Get a task by its ID in a thread-safe way."""
        with self.lock:
            for task in self.tasks.values():
                if task.task_id == task_id:
                    return task
        return None

    def add_task(self, task: PDFTask) -> None:
        with self.lock:
            self.tasks[task.pdf_path] = task
        self.mark_changed()
        self._ensure_processing()

    def clear_completed(self) -> None:
        """Clear completed tasks from the queue."""
        with self.lock:
            self.tasks = {
                k: v for k, v in self.tasks.items() if v.status != "completed"
            }
        self.mark_changed()

    def retry_failed(self) -> None:
        """Retry failed tasks in the queue."""
        with self.lock:
            for task in self.tasks.values():
                if task.status == "failed":
                    task.status = "pending"
                    task.error_msg = ""
        self.mark_changed()
        self._ensure_processing()

    def _ensure_processing(self) -> None:
        if self.processing_thread is None or not self.processing_thread.is_alive():
            self.stop_event.clear()
            self.processing_thread = Thread(target=self._process_queue, daemon=True)
            self.processing_thread.start()

    def get_task_status(self) -> Dict[str, List[PDFTask]]:
        with self.lock:
            result = {
                "pending": [],
                "processing": [],
                "failed": [],
                "completed": [],
                "reverted": [],
                "skipped": [],  # Add skipped status to tracking
            }
            for task in self.tasks.values():
                if task.status in result:
                    result[task.status].append(task)
            return result

    def check_and_clear_changes(self) -> bool:
        """Check if there are changes and clear the flag. Returns whether there were changes."""
        with self.lock:
            had_changes = self.has_changes
            self.has_changes = False
            return had_changes

    def stop(self) -> None:
        self.stop_event.set()
        if self.processing_thread:
            self.processing_thread.join(timeout=1)

    def _process_queue(self) -> None:
        while not self.stop_event.is_set():
            task_to_process = None

            with self.lock:
                for task in self.tasks.values():
                    if task.status == "pending":
                        task.status = "processing"
                        task_to_process = task
                        self.has_changes = True  # Set flag when status changes
                        self._notify_status_change()
                        break

            if task_to_process is None:
                self.stop_event.wait(0.1)
                continue

            try:
                config = self.config_manager.get_config()
                excel_manager = self.excel_manager

                # Load Excel data
                excel_manager.load_excel_data(
                    config["excel_file"], config["excel_sheet"]
                )

                # Get filter columns dynamically based on the number of filter values
                filter_columns = []
                for i in range(1, len(task_to_process.filter_values) + 1):
                    column_key = f"filter{i}_column"
                    if column_key not in config:
                        raise Exception(
                            f"Missing filter column configuration for filter {i}"
                        )
                    filter_columns.append(config[column_key])

                # Get the row index from the second filter value if available
                if len(task_to_process.filter_values) > 1:
                    filter_value, extracted_row_idx = self._parse_filter2_value(
                        task_to_process.filter_values[1]
                    )
                    if extracted_row_idx >= 0:
                        print(
                            f"[DEBUG] Using row index {extracted_row_idx} from filter2 value"
                        )
                        row_idx = extracted_row_idx
                        # Replace the formatted filter2 value with the actual value
                        task_to_process.filter_values[1] = filter_value
                        # Get the row data directly using the index
                        # Verify the row index is within valid range
                        if 0 <= row_idx < len(excel_manager.excel_data):
                            row_data = excel_manager.excel_data.iloc[row_idx]

                            # Verify the data matches our filter values
                            mismatched_filters = []
                            for i, (col, val) in enumerate(
                                zip(filter_columns, task_to_process.filter_values)
                            ):
                                if i != 1:  # Skip filter2 since we already processed it
                                    # Handle date formatting for comparison
                                    row_value = row_data[col]
                                    if "DATE" in col.upper() and pd.notnull(row_value):
                                        if isinstance(row_value, datetime):
                                            row_value = row_value.strftime("%d/%m/%Y")
                                        else:
                                            # Try parsing as date if it's not already a datetime
                                            try:
                                                parsed_date = datetime.strptime(str(row_value).strip(), "%Y-%m-%d %H:%M:%S")
                                                row_value = parsed_date.strftime("%d/%m/%Y")
                                            except ValueError:
                                                row_value = str(row_value).strip()
                                    else:
                                        row_value = str(row_value).strip()

                                    if row_value != str(val).strip():
                                        mismatched_filters.append(
                                            f"{col}: expected '{val}', got '{row_value}'"
                                        )

                            if mismatched_filters:
                                print(
                                    f"[DEBUG] Row {row_idx} data doesn't match filter values"
                                )
                                print(f"[DEBUG] Mismatches: {mismatched_filters}")
                                task_to_process.status = "failed"
                                task_to_process.error_msg = f"Selected row data doesn't match filter values: {', '.join(mismatched_filters)}"
                                self.mark_changed()
                                continue

                        else:
                            print(
                                f"[DEBUG] Row index {row_idx} is out of range (max: {len(excel_manager.excel_data) - 1})"
                            )
                            task_to_process.status = "failed"
                            task_to_process.error_msg = f"Invalid Excel row number {row_idx + 2} (exceeds file length)"
                            self.mark_changed()
                            continue
                    else:
                        print("[DEBUG] Invalid row index extracted from filter2 value")
                        task_to_process.status = "failed"
                        task_to_process.error_msg = "Could not extract valid Excel row number from filter2 value"
                        self.mark_changed()
                        continue
                else:
                    print("[DEBUG] No filter2 value available")
                    task_to_process.status = "failed"
                    task_to_process.error_msg = "Missing filter2 value with row number"
                    self.mark_changed()
                    continue

                # Update task with the row index
                task_to_process.row_idx = row_idx
                print(
                    f"[DEBUG] Using row index: {row_idx} with filter columns: {filter_columns}"
                )

                # Define date formats as a constant
                DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"]

                # Process all columns in one pass
                template_data = {}

                # Add filter values to template data
                for i, (column, value) in enumerate(
                    zip(filter_columns, task_to_process.filter_values), 1
                ):
                    template_data[f"filter{i}"] = value
                    template_data[column] = value

                # Process all columns in one pass
                for column in row_data.index:
                    value = row_data[column]

                    # Handle any column that might contain dates
                    if "DATE" in column.upper():
                        if pd.isnull(value):
                            template_data[column] = None
                        elif isinstance(value, datetime):
                            template_data[column] = value
                        else:
                            # Try to parse the date string
                            parsed_date = None
                            for date_format in DATE_FORMATS:
                                try:
                                    parsed_date = datetime.strptime(
                                        str(value).strip(), date_format
                                    )
                                    break
                                except ValueError:
                                    continue

                            if parsed_date is None:
                                raise ValueError(
                                    f"Could not parse date '{value}' in column '{column}'"
                                )

                            template_data[column] = parsed_date
                    else:
                        template_data[column] = value

                print("[DEBUG] Template data for dates:")
                for col, val in template_data.items():
                    if "DATE" in col.upper():
                        print(f"[DEBUG] {col}: {val} (type: {type(val)})")

                # Add processed_folder to template data and process PDF
                template_data["processed_folder"] = config["processed_folder"]
                processed_path = self.pdf_manager.generate_output_path(
                    config["output_template"], template_data
                )

                # Capture the original hyperlink before updating
                original_hyperlink = self.excel_manager.update_pdf_link(
                    config["excel_file"],
                    config["excel_sheet"],
                    task_to_process.row_idx,
                    processed_path,
                    config["filter2_column"],
                )

                # Assign the captured original hyperlink to the task
                task_to_process.original_excel_hyperlink = original_hyperlink

                # Assign the original PDF location
                task_to_process.original_pdf_location = task_to_process.pdf_path

                # Continue with processing...
                self.pdf_manager.process_pdf(
                    task_to_process,
                    template_data,
                    config["processed_folder"],
                    config["output_template"],
                )

                # Update task status to completed
                with self.lock:
                    task_to_process.status = "completed"
                    self.has_changes = True  # Set flag when status changes
                    self._notify_status_change()

            except Exception as e:
                with self.lock:
                    task_to_process.status = "failed"
                    task_to_process.error_msg = str(e)
                    self.has_changes = True  # Set flag when status changes
                    self._notify_status_change()
                print(f"[DEBUG] Task failed: {str(e)}")
            finally:
                # If task is still in processing state, mark it as failed
                with self.lock:
                    if task_to_process and task_to_process.status == "processing":
                        task_to_process.status = "failed"
                        task_to_process.error_msg = (
                            "Task timed out or failed unexpectedly"
                        )
                        self.has_changes = True  # Set flag when status changes
                        self._notify_status_change()
                        print(
                            "[DEBUG] Task marked as failed due to timeout or unexpected state"
                        )

    def add_skipped_task(self, task: PDFTask) -> None:
        """Add a skipped task to the queue without triggering processing."""
        with self.lock:
            task.end_time = datetime.now()  # Set end time immediately for skipped tasks
            self.tasks[task.pdf_path] = task
        self.mark_changed()


class ProcessingTab(Frame):
    """A modernized tab for processing PDF files with Excel data integration."""

    _instance = None  # Class-level instance tracking

    @classmethod
    def get_instance(cls) -> Optional['ProcessingTab']:
        """Get the current ProcessingTab instance."""
        return cls._instance

    def __init__(
        self,
        master: TkWidget,
        config_manager: ConfigManager,
        excel_manager: ExcelManager,
        pdf_manager: PDFManager,
        error_handler: Callable[[Exception, str], None],
        status_handler: Callable[[str], None],
    ) -> None:
        self._pending_config_change_id = None  # Track pending config change operations
        self._is_reloading = False  # Track Excel data reload state
        ProcessingTab._instance = self  # Store instance
        super().__init__(master)
        self.master = master
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        self._handle_error = error_handler
        self._update_status = status_handler
        self.filter_frames = []  # Store filter frames for dynamic handling

        self.pdf_queue = ProcessingQueue(config_manager, excel_manager, pdf_manager)
        self.current_pdf: Optional[str] = None
        self.current_pdf_start_time: Optional[datetime] = None

        # Configure styles
        self._setup_styles()

        # Setup main layout
        self._setup_ui()
        self.update_queue_display()
        self.after(100, self._periodic_update)

        # Register for config changes
        self.config_manager.add_change_callback(self.on_config_change)

    def load_initial_data(self) -> None:
        """Load initial data asynchronously after window is shown."""
        try:
            config = self.config_manager.get_config()
            if config["source_folder"]:
                self.load_next_pdf()

            self._update_status("Ready")
        except Exception as e:
            self._handle_error(e, "initial data load")

    def on_config_change(self) -> None:
        """Handle configuration changes and preset loading."""
        # Cancel any pending operation
        if self._pending_config_change_id:
            print("[DEBUG] Canceling pending config change operation")
            self.after_cancel(self._pending_config_change_id)

        # Get the current config
        config = self.config_manager.get_config()
        
        # Only proceed if we have valid config values
        if not all([config[key] for key in ["excel_file", "excel_sheet"]]):
            print("[DEBUG] Skipping config change - incomplete configuration")
            return

        print("[DEBUG] Configuration change detected")
        self._update_status("Loading...")

        # Schedule the actual config change with a delay to debounce rapid changes
        def delayed_config_change():
            try:
                print("[DEBUG] Executing delayed config change")
                # Log config details
                print("[DEBUG] Applied preset config:", config)

                # Clear existing filters
                for frame in self.filter_frames:
                    frame["frame"].destroy()
                self.filter_frames.clear()

                # Load new filters from config
                filter_columns = []
                i = 1
                while True:
                    filter_key = f"filter{i}_column"
                    if filter_key not in config:
                        break
                    filter_columns.append(config[filter_key])
                    i += 1

                # Create filter frames
                for i, column in enumerate(filter_columns, 1):
                    self._add_filter(column, f"filter{i}")

                # Complete config change
                self._finish_config_change()
            except Exception as e:
                print(f"[DEBUG] Error in delayed config change: {str(e)}")
                print(traceback.format_exc())

        # Schedule the delayed change
        self._pending_config_change_id = self.after(250, delayed_config_change)

    def _finish_config_change(self) -> None:
        """Complete the configuration change by reloading data and updating status."""
        self._pending_config_change_id = None
        self.reload_excel_data_and_update_ui(trigger_source="_finish_config_change")
        self._check_source_folder_change()
        self._update_status("Ready")
        
    def _check_source_folder_change(self) -> None:
        """Check if source folder changed and refresh PDF viewer if needed."""
        config = self.config_manager.get_config()
        current_source = getattr(self, '_current_source_folder', None)
        new_source = config["source_folder"]
        
        if current_source != new_source:
            self._current_source_folder = new_source
            print(f"[DEBUG] Source folder changed from '{current_source}' to '{new_source}', refreshing PDF viewer")
            
            # Clear current PDF reference
            self.current_pdf = None
            
            # Clear the PDF viewer
            if hasattr(self.pdf_viewer, "canvas"):
                self.pdf_viewer.canvas.delete("all")
                
            # Load the next PDF from the new source folder
            self.load_next_pdf()

    def _setup_styles(self) -> None:
        """Configure custom styles for the interface."""
        style = Style()

        # Configure main theme settings
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))

        # Configure button styles
        style.configure(
            "Primary.TButton",
            padding=10,
            font=("Segoe UI", 10, "bold"),
            background="#007bff",
        )
        style.configure("Secondary.TButton", padding=10, font=("Segoe UI", 10))
        style.configure(
            "Success.TButton",
            padding=10,
            font=("Segoe UI", 10, "bold"),
            background="#28a745",
        )

        # Configure frame styles
        style.configure(
            "Card.TFrame", background="#ffffff", relief="solid", borderwidth=1
        )

        # Configure Treeview
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    def _setup_ui(self) -> None:
        """Setup the main user interface with a modern, clean layout."""
        self.configure(padding=20)

        # Configure grid weights for the main frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create main container with three columns
        main_container = Frame(self)
        main_container.grid(row=0, column=0, sticky="nsew")

        # Configure column weights for the container
        main_container.grid_columnconfigure(0, weight=0)  # Left panel (collapsible)
        main_container.grid_columnconfigure(1, weight=0)  # Handle column
        main_container.grid_columnconfigure(2, weight=8)  # Center panel
        main_container.grid_columnconfigure(3, weight=1)  # Right panel
        main_container.grid_rowconfigure(0, weight=1)

        # Left Panel (collapsible)
        self.left_panel = self._create_left_panel(main_container)
        self.left_panel.grid(row=0, column=0, sticky="nsew")

        # Style for handle frame and button
        style = Style()
        style.configure("Handle.TFrame", background="#e0e0e0")
        style.configure("Handle.TButton", padding=0, relief="flat", borderwidth=0)
        style.map("Handle.TButton", background=[("active", "#d0d0d0")])

        # Handle for collapsing/expanding
        self.handle_frame = Frame(main_container, width=8, style="Handle.TFrame")
        self.handle_frame.grid(row=0, column=1, sticky="nsew")
        self.handle_frame.grid_propagate(False)

        # Create a separate button for collapsing/expanding
        self.handle_button = Button(
            self.handle_frame,
            text="⋮",  # Vertical dots as handle icon
            style="Handle.TButton",
            command=self._toggle_left_panel,
        )
        self.handle_button.place(relx=0.5, rely=0.5, anchor="center", relheight=0.1)

        # Make handle frame draggable for resizing
        self.handle_frame.bind(
            "<Enter>", lambda e: self.handle_frame.configure(cursor="sb_h_double_arrow")
        )
        self.handle_frame.bind(
            "<Leave>", lambda e: self.handle_frame.configure(cursor="")
        )
        self.handle_frame.bind("<Button-1>", self._start_resize)
        self.handle_frame.bind("<B1-Motion>", self._do_resize)
        self.handle_frame.bind("<ButtonRelease-1>", self._end_resize)

        # Center Panel
        self.center_panel = self._create_center_panel(main_container)
        self.center_panel.grid(row=0, column=2, sticky="nsew", padx=(5, 5))

        # Right Panel
        self.right_panel = self._create_right_panel(main_container)
        self.right_panel.grid(row=0, column=3, sticky="nsew", padx=(5, 0))

        # Store initial width of left panel
        self.left_panel_width = 250  # Default width
        self.left_panel_visible = True
        self.left_panel.configure(width=self.left_panel_width)
        self.left_panel.grid_propagate(False)

        # Bind to Configure event to handle window resizing
        self.bind("<Configure>", self._on_window_resize)

    def _start_resize(self, event: TkEvent) -> None:
        """Start resizing the left panel."""
        if not self.left_panel_visible:
            return
        self.start_x = event.x_root
        self.start_width = self.left_panel.winfo_width()
        self.resizing = True
        style = Style()
        style.configure("Handle.TFrame", background="#d0d0d0")
        # Prevent button from interfering with resize
        self.handle_button.place_forget()

    def _do_resize(self, event: TkEvent) -> None:
        """Resize the left panel based on mouse drag."""
        if (
            not hasattr(self, "resizing")
            or not self.resizing
            or not self.left_panel_visible
        ):
            return
        delta_x = event.x_root - self.start_x
        new_width = max(
            200, min(800, self.start_width + delta_x)
        )  # Limit width between 200 and 400

        # Only update if width actually changed
        if new_width != self.left_panel_width:
            self.left_panel_width = new_width
            self.left_panel.configure(width=new_width)
            # Update layout
            self.update_idletasks()

    def _end_resize(self, event: TkEvent) -> None:
        """End the resize operation."""
        if not hasattr(self, "resizing") or not self.resizing:
            return
        self.resizing = False
        style = Style()
        style.configure("Handle.TFrame", background="#e0e0e0")
        # Restore the button after resize
        self.handle_button.place(relx=0.5, rely=0.5, anchor="center", relheight=0.1)

    def _toggle_left_panel(self) -> None:
        """Toggle the visibility of the left panel."""
        if self.left_panel_visible:
            self.left_panel.grid_remove()
            # Right arrow when collapsed
            self.handle_button.configure(text="›")
            self.left_panel_visible = False
        else:
            self.left_panel.grid()
            # Vertical dots when expanded
            self.handle_button.configure(text="⋮")
            self.left_panel_visible = True
            self.left_panel.configure(width=self.left_panel_width)

        # Update layout
        self.update_idletasks()

    def _on_window_resize(self, event: TkEvent) -> None:
        """Handle window resize events."""
        if event.widget == self:
            # Ensure minimum width for panels
            min_width = 800 if self.left_panel_visible else 600
            current_width = self.winfo_width()

            if current_width < min_width:
                # Instead of changing window geometry, adjust panel sizes
                if self.left_panel_visible:
                    self.left_panel_width = max(
                        200, self.left_panel_width - (min_width - current_width)
                    )
                    self.left_panel.configure(width=self.left_panel_width)

            # Update layout
            self.update_idletasks()

    def _create_left_panel(self, parent: TkWidget) -> Frame:
        """Create the left panel containing file information and queue."""
        panel = Frame(parent)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)  # Make queue section expandable

        # File Information Section
        info_frame = LabelFrame(panel, text="File Information", padding=10)
        info_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # File info container
        file_info_container = Frame(info_frame)
        file_info_container.pack(fill="x", pady=5)

        self.file_info = Label(
            file_info_container,
            text="No file loaded (click to select file)",
            style="Header.TLabel",
            wraplength=200,
            cursor="hand2",  # Change cursor to hand when hovering
        )
        self.file_info.pack(fill="x", pady=5)
        self.file_info.bind("<Button-1>", self._on_file_info_click)

        # Queue statistics
        stats_frame = Frame(info_frame)
        stats_frame.pack(fill="x", pady=5)

        self.queue_stats = Label(stats_frame, text="Queue: 0 total")
        self.queue_stats.pack(fill="x")

        # Processing Queue Section
        queue_frame = LabelFrame(panel, text="Processing Queue", padding=10)
        queue_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.queue_display = QueueDisplay(queue_frame)
        self.queue_display.pack(fill="both", expand=True)

        return panel

    def _create_center_panel(self, parent: TkWidget) -> Frame:
        """Create the center panel containing the PDF viewer."""
        panel = Frame(parent)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(0, weight=1)

        # PDF Viewer Section
        viewer_frame = LabelFrame(panel, text="PDF Viewer", padding=10)
        viewer_frame.grid(row=0, column=0, sticky="nsew")
        viewer_frame.grid_columnconfigure(0, weight=1)
        viewer_frame.grid_rowconfigure(0, weight=1)

        self.pdf_viewer = PDFViewer(viewer_frame, self.pdf_manager)
        self.pdf_viewer.grid(row=0, column=0, sticky="nsew")

        # Viewer Controls
        controls_frame = Frame(viewer_frame)
        controls_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        # Zoom Controls
        zoom_frame = Frame(controls_frame)
        zoom_frame.pack(side="left")

        Button(zoom_frame, text="−", width=3, command=self.pdf_viewer.zoom_out).pack(
            side="left", padx=2
        )
        self.zoom_label = Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side="left", padx=5)
        Button(zoom_frame, text="+", width=3, command=self.pdf_viewer.zoom_in).pack(
            side="left", padx=2
        )

        # Rotation Controls
        rotation_frame = Frame(controls_frame)
        rotation_frame.pack(side="right")

        Button(
            rotation_frame, text="↶", width=3, command=self.rotate_counterclockwise
        ).pack(side="left", padx=2)
        self.rotation_label = Label(rotation_frame, text="0°", width=6)
        self.rotation_label.pack(side="left", padx=5)
        Button(rotation_frame, text="↷", width=3, command=self.rotate_clockwise).pack(
            side="left", padx=2
        )

        return panel

    def _create_right_panel(self, parent: TkWidget) -> Frame:
        """Create the right panel containing filters and actions."""
        panel = Frame(parent)

        # Filters Frame (replacing notebook)
        filters_frame = LabelFrame(panel, text="Filters", padding=10)
        filters_frame.pack(fill="both", expand=True, pady=(0, 10))

        self._setup_filters(filters_frame)

        # Action Buttons
        actions_frame = Frame(panel)
        actions_frame.pack(fill="x", pady=(10, 0))

        self.confirm_button = Button(
            actions_frame,
            text="Process File",
            command=self.process_current_file,
            style="Success.TButton",
        )
        self.confirm_button.pack(fill="x", pady=(0, 5))

        # Add key bindings for button activation when focused
        def trigger_if_focused(event):
            if event.widget.focus_get() == event.widget:
                self.process_current_file()

        self.confirm_button.bind("<Return>", trigger_if_focused)

        self.skip_button = Button(
            actions_frame,
            text="Next File (Ctrl+N)",
            command=lambda: self.load_next_pdf(move_to_skipped=True),
            style="Primary.TButton",
        )
        self.skip_button.pack(fill="x")

        return panel

    def _setup_filters(self, parent: TkWidget) -> None:
        """Setup dynamic filter controls based on configuration."""
        # Create container for filters
        self.filters_container = Frame(parent)
        self.filters_container.pack(fill="x", expand=True)

        # Load filters from config
        config = self.config_manager.get_config()
        filter_columns = []
        i = 1
        while True:
            filter_key = f"filter{i}_column"
            if filter_key not in config:
                break
            filter_columns.append(config[filter_key])
            i += 1

        # Create filter frames
        for i, column in enumerate(filter_columns, 1):
            self._add_filter(column, f"filter{i}")

        # Initialize Excel data for filters
        # self.reload_excel_data_and_update_ui(trigger_source="_setup_filters")

    def _add_filter(self, column_name: str, identifier: str) -> None:
        """Add a new filter frame."""
        filter_frame = Frame(self.filters_container)
        filter_frame.pack(fill="x", pady=(0, 15))

        # Label with column name
        label = Label(filter_frame, text=column_name, style="Header.TLabel")
        label.pack(pady=(0, 5))

        # Calculate the current filter index before adding the new frame
        current_index = len(self.filter_frames)

        # Create fuzzy search frame
        fuzzy_frame = FuzzySearchFrame(
            filter_frame,
            width=30,
            identifier=f"processing_{identifier}",
            on_tab=lambda e: self._handle_filter_tab(
                e, current_index
            ),  # Use the pre-calculated index
        )
        fuzzy_frame.pack(fill="x")

        # Store filter info
        self.filter_frames.append(
            {
                "frame": filter_frame,
                "label": label,
                "fuzzy_frame": fuzzy_frame,
                "identifier": identifier,
                "values": [],  # Store available values for this filter
            }
        )

        # Initialize with available values if this is the first filter
        if current_index == 0:  # Use current_index instead of len(self.filter_frames)
            try:
                config = self.config_manager.get_config()
                if config["excel_file"] and config["excel_sheet"]:
                    if self.excel_manager.excel_data is None:
                        self.excel_manager.load_excel_data(
                            config["excel_file"], config["excel_sheet"]
                        )

                    if column_name:
                        df = self.excel_manager.excel_data
                        values = sorted(df[column_name].astype(str).unique().tolist())
                        values = [str(x).strip() for x in values]
                        fuzzy_frame.set_values(values)
            except Exception as e:
                print(f"[DEBUG] Error initializing first filter values: {str(e)}")

        # Bind events
        fuzzy_frame.bind(
            "<<ValueSelected>>", lambda e: self._on_filter_select(current_index)
        )  # Use current_index
        fuzzy_frame.entry.bind("<KeyRelease>", lambda e: self.update_confirm_button())

    def _handle_filter_tab(self, event: Event, filter_index: int) -> str:
        """Handle tab key in filter to move focus to next filter or confirm button."""
        # Move focus to next filter or confirm button
        if filter_index < len(self.filter_frames) - 1:
            self.filter_frames[filter_index + 1]["fuzzy_frame"].entry.focus_set()
        else:
            self.confirm_button.focus_set()

        return "break"

    def _on_filter_select(self, filter_index: int) -> None:
        """Handle filter selection."""
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is None:
                return

            # Get selected values up to current filter using FuzzySearchFrame's get method
            selected_values = []
            selected_row_idx = -1  # Store the row index from filter2 if available
            for i in range(filter_index + 1):
                fuzzy_frame = self.filter_frames[i]["fuzzy_frame"]
                value = fuzzy_frame.get().strip()
                if not value:  # If any previous filter is empty, stop processing
                    # Clear all subsequent filters using FuzzySearchFrame's clear method
                    for j in range(i + 1, len(self.filter_frames)):
                        self.filter_frames[j]["fuzzy_frame"].clear()
                    return
                # For filter2, we need to parse it only if we haven't already gotten a row index
                if i == 1:  # Second filter
                    if selected_row_idx < 0:  # Only parse if we don't have a valid row index
                        _, parsed_row_idx = self.pdf_queue._parse_filter2_value(value)
                        if parsed_row_idx >= 0:
                            selected_row_idx = parsed_row_idx
                selected_values.append(value)  # Keep the formatted value

            # Start with the full DataFrame
            df = self.excel_manager.excel_data.copy()

            # If we're past filter2 and have a valid row index, filter based on that row
            if filter_index >= 1 and selected_row_idx >= 0:
                df = df.iloc[[selected_row_idx]]
            else:
                # Apply filters sequentially based on selected values up to filter2
                for i, value in enumerate(
                    selected_values[: min(2, len(selected_values))]
                ):
                    column = config[f"filter{i + 1}_column"]
                    df = df[df[column].astype(str).str.strip() == value]

            # Update next filter's values if there is one
            if filter_index < len(self.filter_frames) - 1:
                next_filter = self.filter_frames[filter_index + 1]
                next_column = config[f"filter{filter_index + 2}_column"]

                # Special handling for filter2 (index 1) to include row information
                if filter_index == 0:  # This means we're updating filter2
                    filter_values = []
                    for idx, row in df.iterrows():
                        value = str(row[next_column]).strip()
                        has_hyperlink = self.excel_manager.has_hyperlink(idx)
                        formatted_value = self._format_filter2_value(
                            value, idx, has_hyperlink
                        )
                        filter_values.append(formatted_value)
                else:
                    # For filters after filter2, if we have a row index, only show that row's value
                    if selected_row_idx >= 0:
                        # Handle date formatting for the single row
                        value = df.iloc[0][next_column]
                        if "DATE" in next_column.upper():
                            if pd.notnull(value) and isinstance(value, datetime):
                                value = value.strftime("%d/%m/%Y")
                            elif pd.notnull(value):
                                # Try parsing as date if it's not already a datetime
                                for date_format in [
                                    "%d/%m/%Y",
                                    "%d-%m-%Y",
                                    "%Y-%m-%d",
                                    "%Y/%m/%d",
                                ]:
                                    try:
                                        parsed_date = datetime.strptime(
                                            str(value).strip(), date_format
                                        )
                                        value = parsed_date.strftime("%d/%m/%Y")
                                        break
                                    except ValueError:
                                        continue
                        filter_values = (
                            [str(value).strip()] if pd.notnull(value) else []
                        )
                    else:
                        # Handle date formatting for multiple values
                        values = df[next_column].unique()
                        filter_values = []
                        for value in values:
                            if "DATE" in next_column.upper():
                                if pd.notnull(value) and isinstance(value, datetime):
                                    formatted_value = value.strftime("%d/%m/%Y")
                                elif pd.notnull(value):
                                    # Try parsing as date if it's not already a datetime
                                    try:
                                        parsed_date = datetime.strptime(
                                            str(value).strip(), "%d/%m/%Y"
                                        )
                                        formatted_value = parsed_date.strftime(
                                            "%d/%m/%Y"
                                        )
                                    except ValueError:
                                        formatted_value = str(value).strip()
                                else:
                                    continue
                            else:
                                formatted_value = str(value).strip()
                            filter_values.append(formatted_value)
                        filter_values = sorted(filter_values)

                # Use FuzzySearchFrame's methods to update values
                next_filter["fuzzy_frame"].clear()
                next_filter["fuzzy_frame"].set_values(filter_values)

                # Clear all subsequent filters using FuzzySearchFrame's methods
                for i in range(filter_index + 2, len(self.filter_frames)):
                    self.filter_frames[i]["fuzzy_frame"].clear()

                # Special handling for filter2 updates
                if filter_index == 1:
                    print("[DEBUG] Updating filter2 values:")
                    print(f"[DEBUG] - Sheet: {config['excel_sheet']}")
                    print(f"[DEBUG] - Column: {config['filter2_column']}")
                    print(f"[DEBUG] - Cache size before: {len(self.excel_manager._hyperlink_cache)}")
                    print(f"[DEBUG] - Cache key before: {getattr(self.excel_manager, '_last_cached_key', 'None')}")
                    
                    # Ensure hyperlinks are cached
                    self.excel_manager.cache_hyperlinks_for_column(
                        config["excel_file"],
                        config["excel_sheet"],
                        config["filter2_column"]
                    )
                    
                    print(f"[DEBUG] - Cache size after: {len(self.excel_manager._hyperlink_cache)}")
                    print(f"[DEBUG] - Cache key after: {getattr(self.excel_manager, '_last_cached_key', 'None')}")

                    # Format values with hyperlink status
                    current_filter = self.filter_frames[1]
                    current_values = []
                    print("[DEBUG] Formatting filter2 values:")
                    for idx, row in df.iterrows():
                        value = str(row[config["filter2_column"]]).strip()
                        has_hyperlink = self.excel_manager.has_hyperlink(idx)
                        formatted_value = self._format_filter2_value(value, idx, has_hyperlink)
                        print(f"[DEBUG] - Row {idx}: value='{value}', has_hyperlink={has_hyperlink}")
                        current_values.append(formatted_value)
                        
                    if current_values:
                        print(f"[DEBUG] Setting {len(current_values)} filter2 values")
                        current_filter["fuzzy_frame"].set_values(current_values)

            # Update confirm button state after filter selection
            self.update_confirm_button()

        except Exception as e:
            print(f"[DEBUG] Error in _on_filter_select: {str(e)}")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def update_confirm_button(self) -> None:
        """Update the confirm button state based on filter selections."""
        all_filters_selected = all(
            frame["fuzzy_frame"].get() for frame in self.filter_frames
        )

        if all_filters_selected:
            self.confirm_button.state(["!disabled"])
            self._update_status("Ready to process")
        else:
            self.confirm_button.state(["disabled"])
            self._update_status("Select all filters")

    def rotate_clockwise(self) -> None:
        """Rotate the PDF view clockwise."""
        self.pdf_manager.rotate_page(clockwise=True)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(
            self.current_pdf, self.pdf_viewer.zoom_level, show_loading=False
        )

    def rotate_counterclockwise(self) -> None:
        """Rotate the PDF view counterclockwise."""
        self.pdf_manager.rotate_page(clockwise=False)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(
            self.current_pdf, self.pdf_viewer.zoom_level, show_loading=False
        )

    def _move_to_skipped_folder(self, pdf_path: str) -> None:
        """Move a skipped PDF file to the skipped documents folder."""
        try:
            skipped_folder = r"\\192.168.0.77\tarec\Archive\SCANNER\SKIPPED DOCUMENT"
            if not path.exists(skipped_folder):
                makedirs(skipped_folder, exist_ok=True)

            # Get the filename and create the destination path
            filename = path.basename(pdf_path)
            dest_path = path.join(skipped_folder, filename)

            # If file already exists in destination, get a versioned name
            if path.exists(dest_path):
                base_name = path.splitext(filename)[0]
                ext = path.splitext(filename)[1]
                counter = 1
                while path.exists(dest_path):
                    new_filename = f"{base_name}_v{counter}{ext}"
                    dest_path = path.join(skipped_folder, new_filename)
                    counter += 1

            # Clear all PDF handles
            # 1. Clear the PDF viewer canvas
            if hasattr(self.pdf_viewer, "canvas"):
                self.pdf_viewer.canvas.delete("all")

            # 2. Clear the PDF viewer's cached image
            if hasattr(self.pdf_viewer, "current_image"):
                self.pdf_viewer.current_image = None

            # 3. Close any open PDF files in the PDF manager
            self.pdf_manager.clear_cache()  # Clear the cached PDF document
            self.pdf_manager.close_current_pdf()  # Close any other open PDFs

            # Try to move the file with retries
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    copy2(pdf_path, dest_path)
                    remove(pdf_path)
                    self._update_status("File skipped and moved to archive")
                    break
                except PermissionError:
                    retry_count += 1
                    if retry_count == max_retries:
                        raise
                    time.sleep(0.5)  # Increased wait time between retries

        except Exception as e:
            ErrorDialog(self, "Error", f"Failed to move skipped file: {str(e)}")

    def load_next_pdf(self, move_to_skipped: bool = False) -> None:
        """Load the next PDF file from the source folder.

        Args:
            move_to_skipped: If True, moves current file to skipped folder before loading next.
        """
        try:
            config = self.config_manager.get_config()
            current_file = self.current_pdf

            # Clear current PDF reference before moving to prevent double-skipping
            self.current_pdf = None

            # If there's a current PDF and we should move it to skipped
            if move_to_skipped and current_file and path.exists(current_file):
                # Create a skipped task with the existing start time
                task = PDFTask(
                    task_id=PDFTask.generate_id(),
                    pdf_path=current_file,
                    filter_values=[""]
                    * len(self.filter_frames),  # Empty values for all filters
                    status="skipped",  # Set initial status as skipped
                    start_time=self.current_pdf_start_time,
                )

                # Add to queue as skipped (won't trigger processing)
                self.pdf_queue.add_skipped_task(task)

                # Move the file to skipped folder
                self._move_to_skipped_folder(current_file)

            # Reload Excel data to ensure we have fresh data
            if config["excel_file"] and config["excel_sheet"]:
                self.reload_excel_data_and_update_ui(trigger_source="load_next_pdf")

            if not config["source_folder"]:
                self._update_status("Source folder not configured")
                return

            # Clear current display if no PDF is loaded
            if not path.exists(config["source_folder"]):
                self.file_info["text"] = "Source folder not found"
                self._update_status("Source folder does not exist")
                ErrorDialog(
                    self,
                    "Error",
                    f"Source folder not found: {config['source_folder']}",
                )
                return

            # Get active tasks to avoid reloading files being processed
            active_tasks = {}
            with self.pdf_queue.lock:
                active_tasks = {
                    k: v
                    for k, v in self.pdf_queue.tasks.items()
                    if v.status in ["pending", "processing"]
                }

            next_pdf = self.pdf_manager.get_next_pdf(
                config["source_folder"], active_tasks
            )
            if next_pdf:
                self.current_pdf = next_pdf
                self.current_pdf_start_time = (
                    datetime.now()
                )  # Set start time when PDF is loaded
                # Store current source folder for change detection
                self._current_source_folder = config["source_folder"]
                self.file_info["text"] = path.basename(next_pdf)
                self.pdf_viewer.display_pdf(next_pdf, 1)
                self.rotation_label.config(text="0°")
                self.zoom_label.config(text="100%")

                # Clear all filters
                for frame in self.filter_frames:
                    frame["fuzzy_frame"].clear()

                # Reset first filter values if available
                if len(self.filter_frames) > 0:
                    config = self.config_manager.get_config()
                    if (
                        config["excel_file"]
                        and config["excel_sheet"]
                        and self.excel_manager.excel_data is not None
                    ):
                        first_column = config.get("filter1_column")
                        if first_column:
                            df = self.excel_manager.excel_data
                            values = sorted(
                                df[first_column].astype(str).unique().tolist()
                            )
                            values = [str(x).strip() for x in values]
                            self.filter_frames[0]["fuzzy_frame"].set_values(values)

                # Focus the first filter
                if self.filter_frames:
                    self.filter_frames[0]["fuzzy_frame"].entry.focus_set()

                self._update_status("New file loaded")
            else:
                self.file_info["text"] = "No PDF files found"
                self._update_status("No files to process")
                # Clear the PDF viewer
                if hasattr(self.pdf_viewer, "canvas"):
                    self.pdf_viewer.canvas.delete("all")
                # Disable the confirm button since there's no file to process
                self.confirm_button.state(["disabled"])

        except Exception as e:
            print("[DEBUG] Error in load_next_pdf:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")

    def _on_file_info_click(self, event: TkEvent) -> None:
        """Handle click on file info label to open file picker."""
        try:
            config = self.config_manager.get_config()

            # Reload Excel data to ensure we have fresh data
            if config["excel_file"] and config["excel_sheet"]:
                self.reload_excel_data_and_update_ui(trigger_source="_on_file_info_click")

            source_folder = config["source_folder"]

            if not source_folder:
                self._update_status("Source folder not configured")
                return

            if not path.exists(source_folder):
                self.file_info["text"] = "Source folder not found"
                self._update_status("Source folder does not exist")
                ErrorDialog(
                    self,
                    "Error",
                    f"Source folder not found: {source_folder}",
                )
                return

            file_path = filedialog.askopenfilename(
                initialdir=source_folder,
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf")],
            )

            if file_path:
                # Clear current PDF reference and set new one
                self.current_pdf = None  # Clear first to prevent any state issues
                self.current_pdf = file_path
                self.current_pdf_start_time = (
                    datetime.now()
                )  # Set start time when PDF is loaded

                # Update UI elements
                self.file_info["text"] = path.basename(file_path)
                self.pdf_viewer.display_pdf(file_path, 1)
                self.rotation_label.config(text="0°")
                self.zoom_label.config(text="100%")

                # Clear all filters
                for frame in self.filter_frames:
                    frame["fuzzy_frame"].clear()

                # Reset first filter values if available
                if len(self.filter_frames) > 0:
                    config = self.config_manager.get_config()
                    if (
                        config["excel_file"]
                        and config["excel_sheet"]
                        and self.excel_manager.excel_data is not None
                    ):
                        first_column = config.get("filter1_column")
                        if first_column:
                            df = self.excel_manager.excel_data
                            values = sorted(
                                df[first_column].astype(str).unique().tolist()
                            )
                            values = [str(x).strip() for x in values]
                            self.filter_frames[0]["fuzzy_frame"].set_values(values)

                # Focus the first filter
                if self.filter_frames:
                    self.filter_frames[0]["fuzzy_frame"].entry.focus_set()

                # Enable confirm button if needed
                self.update_confirm_button()

                self._update_status("New file loaded")

        except Exception as e:
            print("[DEBUG] Error in _on_file_info_click:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading PDF: {str(e)}")

            # Clear PDF viewer on error
            if hasattr(self.pdf_viewer, "canvas"):
                self.pdf_viewer.canvas.delete("all")
            # Disable the confirm button since there's no file to process
            self.confirm_button.state(["disabled"])

    def _setup_ui(self) -> None:
        """Setup the main user interface with a modern, clean layout."""
        self.configure(padding=20)

        # Configure grid weights for the main frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create main container with three columns
        main_container = Frame(self)
        main_container.grid(row=0, column=0, sticky="nsew")

        # Configure column weights for the container
        main_container.grid_columnconfigure(0, weight=0)  # Left panel (collapsible)
        main_container.grid_columnconfigure(1, weight=0)  # Handle column
        main_container.grid_columnconfigure(2, weight=8)  # Center panel
        main_container.grid_columnconfigure(3, weight=1)  # Right panel
        main_container.grid_rowconfigure(0, weight=1)

        # Left Panel (collapsible)
        self.left_panel = self._create_left_panel(main_container)
        self.left_panel.grid(row=0, column=0, sticky="nsew")

        # Style for handle frame and button
        style = Style()
        style.configure("Handle.TFrame", background="#e0e0e0")
        style.configure("Handle.TButton", padding=0, relief="flat", borderwidth=0)
        style.map("Handle.TButton", background=[("active", "#d0d0d0")])

        # Handle for collapsing/expanding
        self.handle_frame = Frame(main_container, width=8, style="Handle.TFrame")
        self.handle_frame.grid(row=0, column=1, sticky="nsew")
        self.handle_frame.grid_propagate(False)

        # Create a separate button for collapsing/expanding
        self.handle_button = Button(
            self.handle_frame,
            text="⋮",  # Vertical dots as handle icon
            style="Handle.TButton",
            command=self._toggle_left_panel,
        )
        self.handle_button.place(relx=0.5, rely=0.5, anchor="center", relheight=0.1)

        # Make handle frame draggable for resizing
        self.handle_frame.bind(
            "<Enter>", lambda e: self.handle_frame.configure(cursor="sb_h_double_arrow")
        )
        self.handle_frame.bind(
            "<Leave>", lambda e: self.handle_frame.configure(cursor="")
        )
        self.handle_frame.bind("<Button-1>", self._start_resize)
        self.handle_frame.bind("<B1-Motion>", self._do_resize)
        self.handle_frame.bind("<ButtonRelease-1>", self._end_resize)

        # Center Panel
        self.center_panel = self._create_center_panel(main_container)
        self.center_panel.grid(row=0, column=2, sticky="nsew", padx=(5, 5))

        # Right Panel
        self.right_panel = self._create_right_panel(main_container)
        self.right_panel.grid(row=0, column=3, sticky="nsew", padx=(5, 0))

        # Store initial width of left panel
        self.left_panel_width = 250  # Default width
        self.left_panel_visible = True
        self.left_panel.configure(width=self.left_panel_width)
        self.left_panel.grid_propagate(False)

        # Bind to Configure event to handle window resizing
        self.bind("<Configure>", self._on_window_resize)

    def process_current_file(self) -> None:
        """Process the current file."""
        print("[DEBUG] Starting process_current_file")
        try:
            if not self.current_pdf:
                self._update_status("No file selected")
                return

            # Get all filter values and ensure they're all filled
            filter_values = []
            for frame in self.filter_frames:
                value = frame["fuzzy_frame"].get()
                if not value:
                    self._update_status("Select all filters")
                    return
                filter_values.append(value)

            # Get the config to access filter columns
            config = self.config_manager.get_config()

            # Get all filter column names from config
            filter_columns = []
            for i in range(1, len(filter_values) + 1):
                column_key = f"filter{i}_column"
                if column_key in config:
                    filter_columns.append(config[column_key])
                else:
                    raise Exception(
                        f"Missing filter column configuration for filter {i}"
                    )

            # Get the row index from filter2 value if it exists
            row_idx = -1
            if len(filter_values) > 1:
                # Parse the filter2 value to extract row index if present
                filter2_value, extracted_row_idx = self.pdf_queue._parse_filter2_value(filter_values[1])
                print(f"[DEBUG] Parsed filter2 value '{filter_values[1]}' -> value='{filter2_value}', row={extracted_row_idx}")

                if extracted_row_idx >= 0:
                    row_idx = extracted_row_idx
                    # Keep the formatted value in filter_values to preserve row information
                else:
                    # No valid row found - try to add a new row
                    print(f"[DEBUG] No existing row found for filter2 value '{filter2_value}' - attempting to add new row")
                    try:
                        # Create a new row with the filter values
                        filter_values[1] = filter2_value  # Use the raw value without formatting
                        new_row_data, new_row_idx = self.excel_manager.add_new_row(
                            config["excel_file"],
                            config["excel_sheet"],
                            filter_columns,
                            filter_values
                        )
                        
                        # Update row_idx and filter2 value with the new row information
                        row_idx = new_row_idx
                        filter_values[1] = self._format_filter2_value(filter2_value, row_idx, False)
                        print(f"[DEBUG] Added new row {row_idx} for filter2 value '{filter2_value}'")
                        
                    except Exception as e:
                        print(f"[DEBUG] Failed to add new row: {str(e)}")
                        self._update_status(f"Failed to add new row: {str(e)}")
                        return

            # Create PDFTask with a unique task ID, preserving the formatted filter2 value
            task = PDFTask(
                task_id=PDFTask.generate_id(),
                pdf_path=self.current_pdf,
                filter_values=filter_values,  # Keep the formatted value to avoid re-parsing
                row_idx=row_idx,
            )

            # Add task to queue
            self.pdf_queue.add_task(task)

            # Load next file
            self.load_next_pdf()

        except Exception as e:
            print(f"[DEBUG] Error in process_current_file: {str(e)}")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error processing file: {str(e)}")
            self._update_status("Error processing file")

    def _show_error_details(self, event: TkEvent) -> None:
        """Show error details for failed tasks."""
        selection = self.queue_display.table.selection()
        if not selection:
            return

        item = selection[0]
        task_path = self.queue_display.table.item(item)["values"][0]

        with self.pdf_queue.lock:
            task = self.pdf_queue.tasks.get(task_path)
            if task and task.status == "failed" and task.error_msg:
                ErrorDialog(
                    self,
                    "Processing Error",
                    f"Error processing {path.basename(task_path)}:\n{task.error_msg}",
                )

    def _clear_completed(self) -> None:
        """Clear completed tasks from the queue."""
        self.pdf_queue.clear_completed()
        self._update_status("Completed tasks cleared")

    def _retry_failed(self) -> None:
        """Retry failed tasks in the queue."""
        self.pdf_queue.retry_failed()
        self._update_status("Retrying failed tasks")

    def update_queue_display(self) -> None:
        """Update the queue display with current tasks."""
        try:
            with self.pdf_queue.lock:
                tasks = self.pdf_queue.tasks.copy()
                # Update queue statistics
                total = len(tasks)
                completed = sum(1 for t in tasks.values() if t.status == "completed")
                failed = sum(1 for t in tasks.values() if t.status == "failed")
                skipped = sum(1 for t in tasks.values() if t.status == "skipped")
                pending = sum(
                    1 for t in tasks.values() if t.status in ["pending", "processing"]
                )

            # Update the display
            self.queue_display.update_display(tasks)

            # Update statistics display
            if total > 0:
                self.queue_stats.configure(
                    text=f"Queue: {total} total ({completed} completed, {
                        failed
                    } failed, {skipped} skipped, {pending} pending)"
                )
            else:
                self.queue_stats.configure(text="Queue: 0 total")

        except Exception as e:
            print(f"[DEBUG] Error updating queue display: {str(e)}")
            import traceback

            print(traceback.format_exc())

    def _periodic_update(self) -> None:
        """Periodically check for changes and update the queue display only if needed."""
        try:
            if (
                self.pdf_queue.check_and_clear_changes()
            ):  # Only update if there were changes
                self.update_queue_display()
        except Exception as e:
            print(f"[DEBUG] Error in periodic update: {str(e)}")
        finally:
            self.after(500, self._periodic_update)

    def __del__(self) -> None:
        """Clean up resources when the tab is destroyed."""
        try:
            if hasattr(self, "pdf_queue"):
                self.pdf_queue.stop()
        except (RuntimeError, AttributeError) as e:
            # Log but don't raise errors during cleanup since object is being destroyed
            print(f"[DEBUG] Error during ProcessingTab cleanup: {str(e)}")

    def handle_config_change(self) -> None:
        """Handle configuration changes by reloading the current PDF if one is loaded."""
        # Check if source folder has changed
        self._check_source_folder_change()
        
        # If we still have a current PDF after the source folder check, reload it
        if self.current_pdf:
            self.pdf_viewer.display_pdf(self.current_pdf)
            self.update_queue_display()

    def reload_excel_data_and_update_ui(self, trigger_source: str = "unknown") -> None:
        """Reload Excel data and update UI elements.
        
        Args:
            trigger_source: Identifier for the code path triggering the reload
        """
        # Prevent recursive or concurrent reloads
        if self._is_reloading:
            print(f"[DEBUG] Skipping reload from {trigger_source} - reload already in progress")
            return
            
        print(f"[DEBUG] Entering reload_excel_data_and_update_ui - Triggered by: {trigger_source}")
        self._is_reloading = True
        try:
            config = self.config_manager.get_config()
            if not all(
                [
                    config["excel_file"],
                    config["excel_sheet"],
                ]
            ):
                print("Missing configuration values")
                return

            print(
                "[DEBUG] Cache state before Excel load - size:",
                len(self.excel_manager._hyperlink_cache),
            )
            # Load Excel data only if needed and store old filter2 value if it exists
            old_filter2_value = None
            if len(self.filter_frames) > 1:
                old_filter2_value = self.filter_frames[1]["fuzzy_frame"].get()
                if old_filter2_value:
                    print(f"[DEBUG] Preserving filter2 value: {old_filter2_value}")

            excel_loaded = self.excel_manager.load_excel_data(
                config["excel_file"], config["excel_sheet"]
            )
            print(
                f"[DEBUG] Excel data {'was reloaded' if excel_loaded else 'used cached version'}"
            )
            print(
                "[DEBUG] Cache state after Excel load - size:",
                len(self.excel_manager._hyperlink_cache),
            )

            # Restore filter2 value if it was previously set
            if old_filter2_value:
                print(f"[DEBUG] Restoring preserved filter2 value: {old_filter2_value}")
                if len(self.filter_frames) > 1:
                    self.filter_frames[1]["fuzzy_frame"].set_values([old_filter2_value])

            # Cache hyperlinks for filter2 column in all cases to ensure it's up to date
            if len(self.filter_frames) > 1:
                filter2_column = config.get("filter2_column")
                if filter2_column:
                    print("[DEBUG] Processing hyperlinks for filter2:")
                    print(f"[DEBUG] - Sheet: {config['excel_sheet']}")
                    print(f"[DEBUG] - Column: {filter2_column}")
                    print(f"[DEBUG] - Cache size before: {len(self.excel_manager._hyperlink_cache)}")
                    print(f"[DEBUG] - Cache key before: {getattr(self.excel_manager, '_last_cached_key', 'None')}")
                    
                    self.excel_manager.cache_hyperlinks_for_column(
                        config["excel_file"], config["excel_sheet"], filter2_column
                    )
                    
                    print(f"[DEBUG] - Cache size after: {len(self.excel_manager._hyperlink_cache)}")
                    print(f"[DEBUG] - Cache key after: {getattr(self.excel_manager, '_last_cached_key', 'None')}")
                    
                    if len(self.filter_frames) > 1 and self.filter_frames[1]["fuzzy_frame"].get():
                        current_value = self.filter_frames[1]["fuzzy_frame"].get()
                        print(f"[DEBUG] Current filter2 value: {current_value}")

            # Update filter labels
            for i, frame in enumerate(self.filter_frames, 1):
                column_name = config.get(f"filter{i}_column")
                if column_name:
                    frame["label"]["text"] = column_name

            df = self.excel_manager.excel_data

            # Convert all values to strings to ensure consistent type handling
            def safe_convert_to_str(val):
                if pd.isna(val):  # Handle NaN/None values
                    return ""
                return str(val).strip()

            # Store all values for the first filter regardless of Excel reload status
            if len(self.filter_frames) > 0:
                first_column = config.get("filter1_column")
                if first_column:
                    values = sorted(df[first_column].astype(str).unique().tolist())
                    values = [safe_convert_to_str(x) for x in values]
                    self.filter_frames[0]["fuzzy_frame"].set_values(values)

                # Clear other filters
                for frame in self.filter_frames[1:]:
                    frame["fuzzy_frame"].clear()
                    frame["fuzzy_frame"].set_values([])

        except Exception as e:
            print("[DEBUG] Error in reload_excel_data_and_update_ui:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")
        finally:
            self._is_reloading = False
            print("[DEBUG] Completed Excel data reload - cleared reloading flag")

    def _format_filter2_value(
        self, value: str, row_idx: int, has_hyperlink: bool = False
    ) -> str:
        """Format filter2 value with row number and checkmark if hyperlinked."""
        prefix = "✓ " if has_hyperlink else ""
        # +2 because Excel is 1-based and has header
        return f"{prefix}{value} ⟨Excel Row: {row_idx + 2}⟩"
