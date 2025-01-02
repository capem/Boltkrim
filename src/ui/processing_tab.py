from __future__ import annotations
from tkinter import (
    END as TkEND,
    Event as TkEvent,
    Widget as TkWidget,
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
from typing import Optional, Any, Dict, List, Callable
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
            result = {"pending": [], "processing": [], "failed": [], "completed": []}
            for task in self.tasks.values():
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
                excel_manager = type(self.excel_manager)()

                # Load Excel data and find matching row - ExcelManager handles its own errors
                excel_manager.load_excel_data(
                    config["excel_file"], config["excel_sheet"]
                )
                row_data, _ = excel_manager.find_matching_row(
                    config["filter1_column"],
                    config["filter2_column"],
                    config["filter3_column"],
                    task_to_process.value1,
                    task_to_process.value2,
                    task_to_process.value3,
                )

                # Define date formats as a constant
                DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"]

                # Process all columns in one pass
                template_data = {
                    # Add generic filter names first
                    "filter1": row_data[config["filter1_column"]],
                    "filter2": row_data[config["filter2_column"]],
                    "filter3": row_data[config["filter3_column"]],
                }

                # Process all columns in one pass
                for column in row_data.index:
                    value = row_data[column]
                    template_data[column] = value

                    if "DATE" not in column.upper() or pd.isnull(value):
                        continue

                    if isinstance(value, datetime):
                        continue

                    for date_format in DATE_FORMATS:
                        try:
                            template_data[column] = datetime.strptime(
                                str(value), date_format
                            )
                            break
                        except ValueError:
                            continue
                        except Exception as e:
                            print(
                                f"[DEBUG] Failed to parse date in column {column}: {str(e)}"
                            )
                            break

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


class ProcessingTab(Frame):
    """A modernized tab for processing PDF files with Excel data integration."""

    def __init__(
        self,
        master: TkWidget,
        config_manager: ConfigManager,
        excel_manager: ExcelManager,
        pdf_manager: PDFManager,
        error_handler: Callable[[Exception, str], None],
        status_handler: Callable[[str], None],
    ) -> None:
        super().__init__(master)
        self.master = master
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        self._handle_error = error_handler
        self._update_status = status_handler

        self.pdf_queue = ProcessingQueue(config_manager, excel_manager, pdf_manager)
        self.current_pdf: Optional[str] = None

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
        """Handle configuration changes."""
        self._update_status("Loading...")
        self.filter1_frame.clear()
        self.filter2_frame.clear()
        self.filter3_frame.clear()
        self.after(100, self.load_initial_data)

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
            self.handle_button.configure(text="›")  # Right arrow when collapsed
            self.left_panel_visible = False
        else:
            self.left_panel.grid()
            self.handle_button.configure(text="⋮")  # Vertical dots when expanded
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
            text="No file loaded",
            style="Header.TLabel",
            wraplength=200,
        )
        self.file_info.pack(fill="x", pady=5)

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
            text="Process File (Enter)",
            command=self.process_current_file,
            style="Success.TButton",
        )
        self.confirm_button.pack(fill="x", pady=(0, 5))

        self.skip_button = Button(
            actions_frame,
            text="Next File (Ctrl+N)",
            command=lambda: self.load_next_pdf(move_to_skipped=True),
            style="Primary.TButton",
        )
        self.skip_button.pack(fill="x")

        return panel

    def _setup_filters(self, parent: TkWidget) -> None:
        """Setup the filter controls with improved styling."""
        self.filter1_label = Label(parent, text="", style="Header.TLabel")
        self.filter1_label.pack(pady=(0, 5))
        self.filter1_frame = FuzzySearchFrame(
            parent,
            width=30,
            identifier="processing_filter1",
            on_tab=self._handle_filter1_tab,
        )
        self.filter1_frame.pack(fill="x", pady=(0, 15))

        self.filter2_label = Label(parent, text="", style="Header.TLabel")
        self.filter2_label.pack(pady=(0, 5))
        self.filter2_frame = FuzzySearchFrame(
            parent,
            width=30,
            identifier="processing_filter2",
            on_tab=self._handle_filter2_tab,
        )
        self.filter2_frame.pack(fill="x", pady=(0, 15))

        self.filter3_label = Label(parent, text="", style="Header.TLabel")
        self.filter3_label.pack(pady=(0, 5))
        self.filter3_frame = FuzzySearchFrame(
            parent,
            width=30,
            identifier="processing_filter3",
            on_tab=self._handle_filter3_tab,
        )
        self.filter3_frame.pack(fill="x")

        # Bind events
        self.filter1_frame.bind("<<ValueSelected>>", lambda e: self.on_filter1_select())
        self.filter2_frame.bind("<<ValueSelected>>", lambda e: self.on_filter2_select())
        self.filter3_frame.bind(
            "<<ValueSelected>>", lambda e: self.update_confirm_button()
        )

        # Bind keyboard navigation
        self._bind_keyboard_shortcuts()

    def _bind_keyboard_shortcuts(self) -> None:
        """Bind keyboard shortcuts for improved navigation and accessibility."""
        # Tab navigation between filters
        self.filter1_frame.entry.bind("<Tab>", self._handle_filter1_tab)
        self.filter2_frame.entry.bind("<Tab>", self._handle_filter2_tab)
        self.filter3_frame.entry.bind("<Tab>", self._handle_filter3_tab)
        self.filter1_frame.listbox.bind("<Tab>", self._handle_filter1_tab)
        self.filter2_frame.listbox.bind("<Tab>", self._handle_filter2_tab)
        self.filter3_frame.listbox.bind("<Tab>", self._handle_filter3_tab)

        # Bind shortcuts to the main frame
        shortcuts = {
            "<Return>": self._handle_return_key,
            "<Control-n>": lambda e: self.load_next_pdf(move_to_skipped=True),
            "<Control-N>": lambda e: self.load_next_pdf(move_to_skipped=True),
            "<Control-plus>": lambda e: self.pdf_viewer.zoom_in(),
            "<Control-minus>": lambda e: self.pdf_viewer.zoom_out(),
            "<Control-r>": lambda e: self.rotate_clockwise(),
            "<Control-R>": lambda e: self.rotate_counterclockwise(),
        }

        # Bind all shortcuts to the main frame
        for key, callback in shortcuts.items():
            self.bind_all(key, callback)

    def _handle_filter1_tab(self, event: Event) -> str:
        """Handle tab key in filter1 to move focus to filter2."""
        if self.filter1_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item and move to filter2
            if self.filter1_frame.listbox.size() > 0:
                # Only select first item if nothing is currently selected
                if not self.filter1_frame.listbox.curselection():
                    self.filter1_frame.listbox.selection_clear(0, TkEND)
                    self.filter1_frame.listbox.selection_set(0)
                    self.filter1_frame._on_select(None)
        self.filter2_frame.entry.focus_set()
        return "break"

    def _handle_filter2_tab(self, event: Event) -> str:
        """Handle tab key in filter2 to move focus to filter3."""
        if self.filter2_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item
            if self.filter2_frame.listbox.size() > 0:
                # Only select first item if nothing is currently selected
                if not self.filter2_frame.listbox.curselection():
                    self.filter2_frame.listbox.selection_clear(0, TkEND)
                    self.filter2_frame.listbox.selection_set(0)
                    self.filter2_frame._on_select(None)
        self.filter3_frame.entry.focus_set()
        return "break"

    def _handle_filter3_tab(self, event: Event) -> str:
        """Handle tab key in filter3 to move focus to confirm button."""
        if self.filter3_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item
            if self.filter3_frame.listbox.size() > 0:
                # Only select first item if nothing is currently selected
                if not self.filter3_frame.listbox.curselection():
                    self.filter3_frame.listbox.selection_clear(0, TkEND)
                    self.filter3_frame.listbox.selection_set(0)
                    self.filter3_frame._on_select(None)
        self.confirm_button.focus_set()
        return "break"

    def _handle_return_key(self, event: Event) -> str:
        """Handle Return key press to process the current file."""
        if str(self.confirm_button["state"]) != "disabled":
            self.process_current_file()
        return "break"

    def handle_config_change(self) -> None:
        """Handle configuration changes by reloading the current PDF if one is loaded."""
        if self.current_pdf:
            self.pdf_viewer.display_pdf(self.current_pdf)
            self.update_queue_display()

    def reload_excel_data_and_update_ui(self) -> None:
        try:
            config = self.config_manager.get_config()
            if not all(
                [
                    config["excel_file"],
                    config["excel_sheet"],
                    config["filter1_column"],
                    config["filter2_column"],
                    config["filter3_column"],
                ]
            ):
                print("Missing configuration values")
                return

            self.excel_manager.load_excel_data(
                config["excel_file"], config["excel_sheet"]
            )

            # Cache hyperlinks for filter2 column only
            self.excel_manager.cache_hyperlinks_for_column(
                config["excel_file"], config["excel_sheet"], config["filter2_column"]
            )

            self.filter1_label["text"] = config["filter1_column"]
            self.filter2_label["text"] = config["filter2_column"]
            self.filter3_label["text"] = config["filter3_column"]

            df = self.excel_manager.excel_data

            # Convert all values to strings to ensure consistent type handling
            def safe_convert_to_str(val):
                if pd.isna(val):  # Handle NaN/None values
                    return ""
                return str(val).strip()

            # Convert column values to strings
            self.all_values1 = sorted(
                df[config["filter1_column"]].astype(str).unique().tolist()
            )
            self.all_values2 = sorted(
                df[config["filter2_column"]].astype(str).unique().tolist()
            )
            self.all_values3 = sorted(
                df[config["filter3_column"]].astype(str).unique().tolist()
            )

            # Strip whitespace and ensure string type
            self.all_values1 = [safe_convert_to_str(x) for x in self.all_values1]
            self.all_values2 = [safe_convert_to_str(x) for x in self.all_values2]
            self.all_values3 = [safe_convert_to_str(x) for x in self.all_values3]

            self.filter1_frame.set_values(self.all_values1)

        except Exception as e:
            print(f"[DEBUG] Error in reload_excel_data_and_update_ui:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")

    def _format_filter2_value(
        self, value: str, row_idx: int, has_hyperlink: bool = False
    ) -> str:
        """Format filter2 value with row number and checkmark if hyperlinked."""
        prefix = "✓ " if has_hyperlink else ""
        return f"{prefix}{value} ⟨Excel Row: {row_idx + 2}⟩"  # +2 because Excel is 1-based and has header

    def _parse_filter2_value(self, formatted_value: str) -> tuple[str, int]:
        """Parse filter2 value to get original value and row number."""
        import re

        # Remove checkmark if present
        formatted_value = formatted_value.replace("✓ ", "", 1)

        match = re.match(r"(.*?)\s*⟨Excel Row:\s*(\d+)⟩", formatted_value)
        if match:
            value = match.group(1).strip()
            row_num = int(match.group(2))
            return value, row_num - 2  # Convert back to 0-based index
        return formatted_value, -1

    def on_filter1_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value = str(self.filter1_frame.get()).strip()

                df = self.excel_manager.excel_data
                # Convert column to string for comparison
                df[config["filter1_column"]] = df[config["filter1_column"]].astype(str)
                filtered_df = df[
                    df[config["filter1_column"]].str.strip() == selected_value
                ]

                # Create list of tuples with values and row indices using cached hyperlink info
                filter2_values = []
                for idx, row in filtered_df.iterrows():
                    value = str(row[config["filter2_column"]]).strip()
                    has_hyperlink = self.excel_manager.has_hyperlink(idx)
                    formatted_value = self._format_filter2_value(
                        value, idx, has_hyperlink
                    )
                    filter2_values.append(formatted_value)

                self.filter2_frame.clear()
                self.filter2_frame.set_values(sorted(filter2_values))

                # Clear filter3 since no filter2 value is selected yet
                self.filter3_frame.clear()
                self.filter3_frame.set_values([])

        except Exception as e:
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def on_filter2_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value1 = str(self.filter1_frame.get()).strip()
                selected_value2_formatted = str(self.filter2_frame.get()).strip()

                # Parse the selected value to get original value and row number
                selected_value2, row_idx = self._parse_filter2_value(
                    selected_value2_formatted
                )

                df = self.excel_manager.excel_data
                # Convert columns to string for comparison
                df[config["filter1_column"]] = df[config["filter1_column"]].astype(str)
                df[config["filter2_column"]] = df[config["filter2_column"]].astype(str)

                # Filter based on row index if available
                if row_idx >= 0:
                    filtered_df = df.iloc[[row_idx]]
                else:
                    # Fallback to old behavior if row parsing fails
                    filtered_df = df[
                        (df[config["filter1_column"]].str.strip() == selected_value1)
                        & (df[config["filter2_column"]].str.strip() == selected_value2)
                    ]

                filtered_values3 = sorted(
                    filtered_df[config["filter3_column"]].astype(str).tolist()
                )
                filtered_values3 = [str(x).strip() for x in filtered_values3]
                self.filter3_frame.clear()
                self.filter3_frame.set_values(filtered_values3)

        except Exception as e:
            import traceback

            print(f"[DEBUG] Error in on_filter2_select:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def update_confirm_button(self) -> None:
        """Update the confirm button state based on filter selections."""
        if (
            self.filter1_frame.get()
            and self.filter2_frame.get()
            and self.filter3_frame.get()
        ):
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
                    self._update_status(f"File skipped and moved to archive")
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
                self._move_to_skipped_folder(current_file)

            # Reload Excel data to ensure we have fresh data
            if config["excel_file"] and config["excel_sheet"]:
                self.reload_excel_data_and_update_ui()

            if not config["source_folder"]:
                self._update_status("Source folder not configured")
                return

            # Clear current display if no PDF is loaded
            if not path.exists(config["source_folder"]):
                self.file_info["text"] = "Source folder not found"
                self._update_status("Source folder does not exist")
                ErrorDialog(
                    self, "Error", f"Source folder not found: {config['source_folder']}"
                )
                return

            next_pdf = self.pdf_manager.get_next_pdf(config["source_folder"])
            if next_pdf:
                self.current_pdf = next_pdf
                self.file_info["text"] = path.basename(next_pdf)
                self.pdf_viewer.display_pdf(next_pdf, 1)
                self.rotation_label.config(text="0°")
                self.zoom_label.config(text="100%")

                # Clear all filters
                self.filter1_frame.clear()
                self.filter2_frame.clear()
                self.filter3_frame.clear()

                # Reset available values for dependent filters
                self.filter1_frame.set_values(self.all_values1)

                # Focus the first fuzzy search entry
                self.filter1_frame.entry.focus_set()

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
            import traceback

            print(f"[DEBUG] Error in load_next_pdf:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")

    def process_current_file(self) -> None:
        """Process the current PDF file with selected filters."""
        if not self.current_pdf:
            self._update_status("No file loaded")
            ErrorDialog(self, "Error", "No PDF file loaded")
            return

        if not path.exists(self.current_pdf):
            self._update_status("File no longer exists")
            ErrorDialog(self, "Error", f"File no longer exists: {self.current_pdf}")
            self.load_next_pdf(move_to_skipped=False)  # Don't move to skipped if file doesn't exist
            return

        try:
            value1 = self.filter1_frame.get()
            value2_formatted = self.filter2_frame.get()
            value3 = self.filter3_frame.get()

            if not value1 or not value2_formatted or not value3:
                self._update_status("Select all filters")
                return

            # Parse the actual value and row index from the formatted filter2 value
            value2, row_idx = self._parse_filter2_value(value2_formatted)

            # Verify the combination exists in Excel
            config = self.config_manager.get_config()
            self.excel_manager.find_matching_row(
                config["filter1_column"],
                config["filter2_column"],
                config["filter3_column"],
                value1,
                value2,
                value3,
            )
            # Create task using the parsed values
            task = PDFTask(
                task_id=PDFTask.generate_id(),
                pdf_path=self.current_pdf,
                value1=value1,
                value2=value2,
                value3=value3,
                row_idx=row_idx,
            )

            # Add to processing queue (this will mark changes and trigger update)
            self.pdf_queue.add_task(task)

            # Update status and load next file
            self._update_status("File queued for processing")
            self.load_next_pdf(move_to_skipped=False)  # Don't move to skipped since we're processing it

            # Clear all filters
            self.filter1_frame.clear()
            self.filter2_frame.clear()
            self.filter3_frame.clear()

            # Reset available values for dependent filters
            self.filter1_frame.set_values(self.all_values1)

            # Focus the first fuzzy search entry
            self.filter1_frame.entry.focus_set()

        except Exception as e:
            self._update_status("Processing error")
            ErrorDialog(self, "Error", str(e))

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
                pending = sum(
                    1 for t in tasks.values() if t.status in ["pending", "processing"]
                )

            # Update the display
            self.queue_display.update_display(tasks)

            # Update statistics display
            if total > 0:
                self.queue_stats.configure(
                    text=f"Queue: {total} total ({completed} completed, {failed} failed, {pending} pending)"
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
        except:
            pass
