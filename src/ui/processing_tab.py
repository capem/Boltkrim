from __future__ import annotations
from tkinter import (
    Canvas,
    END,
    Event as TkEvent,
    Widget,
    Toplevel,
    Frame as TkFrame,
)
from tkinter.ttk import (
    Frame,
    Scrollbar,
    Label,
    Button,
    Style,
    LabelFrame,
    Treeview,
)
from PIL.ImageTk import PhotoImage as PILPhotoImage
from os import path
from typing import Optional, Any, Dict, List, Callable
from threading import Thread, Lock, Event
from dataclasses import dataclass
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from ..utils import ConfigManager, ExcelManager, PDFManager


# Data Models
@dataclass
class PDFTask:
    pdf_path: str
    value1: str
    value2: str
    value3: str
    status: str = "pending"  # pending, processing, failed, completed
    error_msg: str = ""
    row_idx: int = -1  # Add row index field


# Queue Management
class ProcessingQueue:
    def __init__(self, config_manager: ConfigManager, excel_manager: ExcelManager, pdf_manager: PDFManager):
        self.tasks: Dict[str, PDFTask] = {}
        self.lock = Lock()
        self.processing_thread: Optional[Thread] = None
        try:
            self.stop_event = Event()
        except (AttributeError, RuntimeError):
            # Fallback for PyInstaller if Event initialization fails
            self.stop_event = Event()

        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        self._callbacks: List[Callable] = []

    def add_task(self, task: PDFTask) -> None:
        with self.lock:
            self.tasks[task.pdf_path] = task
        self._notify_status_change()
        self._ensure_processing()

    def _notify_status_change(self) -> None:
        for callback in self._callbacks:
            try:
                callback()
            except Exception:
                pass

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

    def clear_completed(self) -> None:
        with self.lock:
            self.tasks = {
                k: v for k, v in self.tasks.items() if v.status not in ["completed"]
            }

    def retry_failed(self) -> None:
        with self.lock:
            for task in self.tasks.values():
                if task.status == "failed":
                    task.status = "pending"
                    task.error_msg = ""
        self._ensure_processing()

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
                            template_data[column] = datetime.strptime(str(value), date_format)
                            break
                        except ValueError:
                            continue
                        except Exception as e:
                            print(f"[DEBUG] Failed to parse date in column {column}: {str(e)}")
                            break

                # Add processed_folder to template data and process PDF
                template_data["processed_folder"] = config["processed_folder"]
                processed_path = self.pdf_manager.generate_output_path(
                    config["output_template"], template_data
                )

                self.pdf_manager.process_pdf(
                    task_to_process.pdf_path,
                    template_data,
                    config["processed_folder"],
                    config["output_template"],
                )

                # Update Excel with the new path using the stored row index
                excel_manager.update_pdf_link(
                    config["excel_file"],
                    config["excel_sheet"],
                    task_to_process.row_idx,  # Use the stored row index
                    processed_path,
                    config["filter2_column"],
                )

                with self.lock:
                    task_to_process.status = "completed"
                    self._notify_status_change()

            except Exception as e:
                with self.lock:
                    task_to_process.status = "failed"
                    task_to_process.error_msg = str(e)
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
                        self._notify_status_change()
                        print(
                            "[DEBUG] Task marked as failed due to timeout or unexpected state"
                        )


# UI Components
class PDFViewer(Frame):
    """A modernized PDF viewer widget with zoom and scroll capabilities."""

    def __init__(self, master: Widget, pdf_manager: Any):
        super().__init__(master)
        self.pdf_manager = pdf_manager
        self.current_image: Optional[PILPhotoImage] = None
        self.current_pdf: Optional[str] = None
        self.zoom_level = 1.0

        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_ui()

    def setup_ui(self) -> None:
        """Setup the PDF viewer interface."""
        # Create a container frame with fixed padding for scrollbars
        self.container_frame = Frame(self)
        self.container_frame.grid(row=0, column=0, sticky="nsew")
        self.container_frame.grid_columnconfigure(0, weight=1)
        self.container_frame.grid_rowconfigure(0, weight=1)

        # Create canvas with modern styling
        self.canvas = Canvas(
            self.container_frame,
            bg="#f8f9fa",  # Light gray background
            highlightthickness=0,  # Remove border
            width=20,  # Minimum width to prevent collapse
            height=20,  # Minimum height to prevent collapse
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Modern scrollbars - always create them to reserve space
        self.h_scrollbar = Scrollbar(
            self.container_frame, orient="horizontal", command=self.canvas.xview
        )
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.v_scrollbar = Scrollbar(
            self.container_frame, orient="vertical", command=self.canvas.yview
        )
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure canvas scrolling
        self.canvas.configure(
            xscrollcommand=self._on_x_scroll, yscrollcommand=self._on_y_scroll
        )

        # Initially hide scrollbars but keep their space reserved
        self.h_scrollbar.grid_remove()
        self.v_scrollbar.grid_remove()

        # Create a frame for the loading message that won't affect layout
        self.loading_frame = Frame(self.container_frame)
        self.loading_frame.place(relx=0.5, rely=0.5, anchor="center")

        self._bind_events()

    def _on_x_scroll(self, *args) -> None:
        """Handle horizontal scrolling and scrollbar visibility."""
        self.h_scrollbar.set(*args)
        self._update_scrollbar_visibility()

    def _on_y_scroll(self, *args) -> None:
        """Handle vertical scrolling and scrollbar visibility."""
        self.v_scrollbar.set(*args)
        self._update_scrollbar_visibility()

    def _update_scrollbar_visibility(self) -> None:
        """Update scrollbar visibility based on content size."""
        if not self.current_image:
            self.h_scrollbar.grid_remove()
            self.v_scrollbar.grid_remove()
            return

        # Get the scroll region and canvas size
        x1, y1, x2, y2 = (
            self.canvas.bbox("all") if self.canvas.find_all() else (0, 0, 0, 0)
        )
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Show/hide horizontal scrollbar
        if x2 - x1 > canvas_width:
            self.h_scrollbar.grid()
        else:
            self.h_scrollbar.grid_remove()

        # Show/hide vertical scrollbar
        if y2 - y1 > canvas_height:
            self.v_scrollbar.grid()
        else:
            self.v_scrollbar.grid_remove()

    def _bind_events(self) -> None:
        """Bind mouse and keyboard events."""

        def _on_mousewheel(event: Event) -> None:
            if event.state & 4:  # Ctrl key
                if event.delta > 0:
                    self.zoom_in()
                else:
                    self.zoom_out()
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event: Event) -> None:
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event: Event) -> None:
            self.canvas.unbind_all("<MouseWheel>")

        # Bind mousewheel only when mouse is over the PDF viewer area
        self.canvas.bind("<Enter>", _bind_mousewheel)
        self.canvas.bind("<Leave>", _unbind_mousewheel)
        self.v_scrollbar.bind("<Enter>", _bind_mousewheel)
        self.v_scrollbar.bind("<Leave>", _unbind_mousewheel)

        # Pan functionality
        self.canvas.bind("<Button-1>", self._start_drag)
        self.canvas.bind("<B1-Motion>", self._do_drag)
        self.canvas.bind("<ButtonRelease-1>", self._stop_drag)

        # Window resize handling
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Key>", self._on_key)

    def _start_drag(self, event: Event) -> None:
        """Start panning the view."""
        self.canvas.scan_mark(event.x, event.y)
        self.canvas.configure(cursor="fleur")

    def _do_drag(self, event: Event) -> None:
        """Continue panning the view."""
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _stop_drag(self, event: Event) -> None:
        """Stop panning the view."""
        self.canvas.configure(cursor="")

    def _on_key(self, event: Event) -> None:
        """Handle keyboard navigation."""
        key = event.keysym
        shift_pressed = event.state & 0x1

        if key == "Up":
            self.canvas.yview_scroll(-1 * (5 if shift_pressed else 1), "units")
        elif key == "Down":
            self.canvas.yview_scroll(1 * (5 if shift_pressed else 1), "units")
        elif key == "Left":
            self.canvas.xview_scroll(-1 * (5 if shift_pressed else 1), "units")
        elif key == "Right":
            self.canvas.xview_scroll(1 * (5 if shift_pressed else 1), "units")
        elif key == "Prior":  # Page Up
            self.canvas.yview_scroll(-1, "pages")
        elif key == "Next":  # Page Down
            self.canvas.yview_scroll(1, "pages")
        elif key == "Home":
            self.canvas.yview_moveto(0)
        elif key == "End":
            self.canvas.yview_moveto(1)

    def _on_resize(self, event: Event) -> None:
        """Handle window resize events."""
        if event.widget == self.canvas:
            self._center_image()
            self._update_scrollbar_visibility()

    def _center_image(self) -> None:
        """Center the PDF image in the canvas."""
        if not self.current_image:
            return

        # Get dimensions
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        image_width = self.current_image.width()
        image_height = self.current_image.height()

        # Calculate centering offsets
        x = max(0, (canvas_width - image_width) // 2)
        y = max(0, (canvas_height - image_height) // 2)

        # Set scroll region to image bounds plus padding
        scroll_width = max(canvas_width, image_width + x * 2)
        scroll_height = max(canvas_height, image_height + y * 2)

        self.canvas.configure(scrollregion=(0, 0, scroll_width, scroll_height))

        # Clear and redraw image
        self.canvas.delete("all")
        image_x = (scroll_width - image_width) // 2
        image_y = (scroll_height - image_height) // 2
        self.canvas.create_image(
            image_x, image_y, anchor="nw", image=self.current_image
        )

        # Update scrollbar visibility
        self._update_scrollbar_visibility()

    def display_pdf(
        self, pdf_path: str, zoom: float = 1.0, show_loading: bool = True
    ) -> None:
        """Display a PDF file with the specified zoom level."""
        try:
            self.current_pdf = pdf_path
            self.zoom_level = zoom

            # Show loading message using place geometry manager
            loading_label = None
            if show_loading:
                loading_label = Label(
                    self.loading_frame, text="Loading PDF...", font=("Segoe UI", 10)
                )
                loading_label.pack(pady=20)
                self.loading_frame.lift()  # Bring loading message to front
                self.update()

            # Render PDF
            image = self.pdf_manager.render_pdf_page(pdf_path, zoom=zoom)

            if loading_label:
                loading_label.destroy()
                self.loading_frame.place_forget()  # Hide the loading frame

            self.current_image = PILPhotoImage(image)
            self._center_image()
            self.canvas.focus_set()

        except Exception as e:
            if loading_label:
                loading_label.destroy()
                self.loading_frame.place_forget()  # Hide the loading frame in case of error
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")

    def zoom_in(self, step: float = 0.2) -> None:
        """Zoom in the PDF view."""
        if self.current_pdf:
            self.zoom_level = min(3.0, self.zoom_level + step)
            self.display_pdf(self.current_pdf, self.zoom_level, show_loading=False)

    def zoom_out(self, step: float = 0.2) -> None:
        """Zoom out the PDF view."""
        if self.current_pdf:
            self.zoom_level = max(0.2, self.zoom_level - step)
            self.display_pdf(self.current_pdf, self.zoom_level, show_loading=False)


class QueueDisplay(Frame):
    def __init__(self, master: Widget):
        super().__init__(master)
        self.setup_ui()

    def setup_ui(self) -> None:
        # Configure grid weights for responsive layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)  # For buttons

        # Create a frame for the table and scrollbar with a light background
        table_frame = Frame(self, style="Card.TFrame")
        table_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Setup table with more informative columns
        columns = ("filename", "values", "status", "time")
        self.table = Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",  # Single selection mode
            style="Queue.Treeview"
        )

        # Configure modern style for the treeview
        style = Style()
        style.configure(
            "Queue.Treeview",
            background="#ffffff",
            foreground="#333333",
            rowheight=30,
            fieldbackground="#ffffff",
            borderwidth=0,
            font=("Segoe UI", 9)
        )
        style.configure(
            "Queue.Treeview.Heading",
            background="#f0f0f0",
            foreground="#333333",
            relief="flat",
            font=("Segoe UI", 9, "bold")
        )
        style.map(
            "Queue.Treeview",
            background=[("selected", "#e7f3ff")],
            foreground=[("selected", "#000000")]
        )

        # Configure headings with sort functionality and modern look
        self.table.heading("filename", text="File", anchor="w", command=lambda: self._sort_column("filename"))
        self.table.heading("values", text="Selected Values", anchor="w", command=lambda: self._sort_column("values"))
        self.table.heading("status", text="Status", anchor="w", command=lambda: self._sort_column("status"))
        self.table.heading("time", text="Time", anchor="w", command=lambda: self._sort_column("time"))

        # Configure column properties
        self.table.column("filename", width=250, minwidth=150, stretch=True)
        self.table.column("values", width=250, minwidth=150, stretch=True)
        self.table.column("status", width=100, minwidth=80, stretch=False)
        self.table.column("time", width=80, minwidth=80, stretch=False)

        # Add modern-looking scrollbars
        style.configure("Queue.Vertical.TScrollbar", background="#ffffff", troughcolor="#f0f0f0", width=10)
        style.configure("Queue.Horizontal.TScrollbar", background="#ffffff", troughcolor="#f0f0f0", width=10)

        v_scrollbar = Scrollbar(
            table_frame,
            orient="vertical",
            command=self.table.yview,
            style="Queue.Vertical.TScrollbar"
        )
        self.table.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.table.xview,
            style="Queue.Horizontal.TScrollbar"
        )
        self.table.configure(xscrollcommand=h_scrollbar.set)

        # Grid table and scrollbars
        self.table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure status colors
        self.table.tag_configure("pending", foreground="#666666")
        self.table.tag_configure("processing", foreground="#007bff")
        self.table.tag_configure("completed", foreground="#28a745")
        self.table.tag_configure("failed", foreground="#dc3545")

        # Add status icons
        self.status_icons = {
            "pending": "⋯",  # Three dots
            "processing": "↻",  # Rotating arrow
            "completed": "✓",  # Checkmark
            "failed": "✗"  # X mark
        }

        # Bind events for interactivity
        self.table.bind("<Double-1>", self._show_task_details)
        self.table.bind("<Return>", self._show_task_details)

    def _sort_column(self, column: str) -> None:
        """Sort treeview column."""
        # Get all items in the treeview
        items = [(self.table.set(item, column), item) for item in self.table.get_children("")]
        
        # Check if we're reversing the sort
        reverse = False
        if hasattr(self, "_last_sort") and self._last_sort == (column, False):
            reverse = True
        
        # Store the sort state
        self._last_sort = (column, reverse)
        
        # Sort the items
        items.sort(reverse=reverse)
        
        # Rearrange items in sorted positions
        for index, (_, item) in enumerate(items):
            self.table.move(item, "", index)

        # Update the heading to show sort direction
        for col in self.table["columns"]:
            if col == column:
                self.table.heading(col, text=f"{self.table.heading(col)['text']} {'↓' if reverse else '↑'}")
            else:
                # Remove sort indicator from other columns
                self.table.heading(col, text=self.table.heading(col)['text'].split(' ')[0])

    def _format_values_display(self, values_str: str) -> str:
        """Format the values string for better readability."""
        if not values_str or " | " not in values_str:
            return values_str
        
        parts = values_str.split(" | ")
        return " → ".join(parts)  # Using arrow for better visual flow

    def update_display(self, tasks: Dict[str, PDFTask]) -> None:
        """Update the queue display with current tasks."""
        selection = self.table.selection()
        selected_paths = [self.table.item(item)["values"][0] for item in selection]

        current_items = {}
        for item in self.table.get_children():
            path_value = self.table.item(item)["values"][0]
            current_items[path_value] = item

        for task_path, task in tasks.items():
            # Format the values string for better readability
            values_str = self._format_values_display(f"{task.value1} | {task.value2} | {task.value3}")

            # Get current time for processing tasks
            time_str = datetime.now().strftime("%H:%M:%S") if task.status == "processing" else ""

            # Add status icon to status text
            status_display = f"{self.status_icons[task.status]} {task.status.capitalize()}"

            # Create display values
            display_values = (
                path.basename(task_path),  # Show only filename in table
                values_str,
                status_display,
                time_str,
            )

            if task_path in current_items:
                # Update existing item
                item_id = current_items[task_path]
                for idx, value in enumerate(display_values):
                    self.table.set(item_id, column=idx, value=value)
                self.table.item(item_id, tags=(task.status,))
                current_items.pop(task_path)
            else:
                # Insert new item
                item = self.table.insert(
                    "",
                    0,
                    values=display_values,
                    tags=(task.status,),
                )
                if task_path in selected_paths:
                    self.table.selection_add(item)

        # Remove items that no longer exist
        for item_id in current_items.values():
            self.table.delete(item_id)

    def _show_task_details(self, event: TkEvent) -> None:
        """Show task details in a dialog when double-clicking a row."""
        item = self.table.identify("item", event.x, event.y)
        if not item:
            return

        # Get the task values
        values = self.table.item(item)["values"]
        if not values:
            return

        # Create a dialog window
        dialog = Toplevel(self)
        dialog.title("Task Details")
        dialog.transient(self)  # Make dialog modal
        dialog.grab_set()  # Make dialog modal

        # Calculate position to center the dialog
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (400 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")

        # Create a frame with padding
        frame = TkFrame(dialog, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        # Add details with proper formatting and spacing
        details = [
            ("File:", path.basename(values[0])),
            ("Full Path:", values[0]),
            ("Selected Values:", values[1]),
            ("Status:", values[2]),
            ("Time:", values[3] if len(values) > 3 else ""),
        ]

        # Add each detail with proper styling
        for i, (label, value) in enumerate(details):
            Label(frame, text=label, font=("Segoe UI", 10, "bold")).grid(
                row=i, column=0, sticky="w", pady=(0, 10)
            )
            Label(frame, text=value, wraplength=250).grid(
                row=i, column=1, sticky="w", padx=(10, 0), pady=(0, 10)
            )

        # Add error message if status is failed
        if values[2].lower() == "failed":
            Label(frame, text="Error Message:", font=("Segoe UI", 10, "bold")).grid(
                row=len(details), column=0, sticky="w", pady=(10, 0)
            )

            # Get error message from the task
            task_path = values[0]
            error_msg = "No error message available"

            # Get the parent ProcessingTab instance
            processing_tab = self.master.master
            if hasattr(processing_tab, "pdf_queue"):
                with processing_tab.pdf_queue.lock:
                    task = processing_tab.pdf_queue.tasks.get(task_path)
                    if task and task.error_msg:
                        error_msg = task.error_msg

            error_label = Label(
                frame, text=error_msg, wraplength=300, justify="left", foreground="red"
            )
            error_label.grid(
                row=len(details), column=1, sticky="w", padx=(10, 0), pady=(10, 0)
            )

        # Add close button at the bottom
        Button(frame, text="Close", command=dialog.destroy).grid(
            row=len(details) + 1, column=0, columnspan=2, pady=(20, 0)
        )

        # Make dialog resizable
        dialog.resizable(True, True)

        # Focus the dialog
        dialog.focus_set()


class ProcessingTab(Frame):
    """A modernized tab for processing PDF files with Excel data integration."""

    def __init__(
        self,
        master: Widget,
        config_manager: ConfigManager,
        excel_manager: ExcelManager,
        pdf_manager: PDFManager,
        error_handler: Callable[[Exception, str], None],
        status_handler: Callable[[str], None]
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

    def _create_left_panel(self, parent: Widget) -> Frame:
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

    def _create_center_panel(self, parent: Widget) -> Frame:
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

    def _create_right_panel(self, parent: Widget) -> Frame:
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
            command=self.load_next_pdf,
            style="Primary.TButton",
        )
        self.skip_button.pack(fill="x")

        return panel

    def _setup_filters(self, parent: Widget) -> None:
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
            "<Control-n>": lambda e: self.load_next_pdf(),
            "<Control-N>": lambda e: self.load_next_pdf(),
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
                    self.filter1_frame.listbox.selection_clear(0, END)
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
                    self.filter2_frame.listbox.selection_clear(0, END)
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
                    self.filter3_frame.listbox.selection_clear(0, END)
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
                config["excel_file"],
                config["excel_sheet"],
                config["filter2_column"]
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
            import traceback
            print(f"[DEBUG] Error in reload_excel_data_and_update_ui:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")

    def _format_filter2_value(self, value: str, row_idx: int, has_hyperlink: bool = False) -> str:
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
                    formatted_value = self._format_filter2_value(value, idx, has_hyperlink)
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

    def load_next_pdf(self) -> None:
        """Load the next PDF file from the source folder."""
        try:
            config = self.config_manager.get_config()
            
            # Reload Excel data to ensure we have fresh data
            if config["excel_file"] and config["excel_sheet"]:
                self.reload_excel_data_and_update_ui()
                
            if not config["source_folder"]:
                self._update_status("Source folder not configured")
                return

            # Clear current display if no PDF is loaded
            if not path.exists(config["source_folder"]):
                self.current_pdf = None
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
                self.current_pdf = None
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
            self.load_next_pdf()  # Try to load the next file
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
                pdf_path=self.current_pdf,
                value1=value1,
                value2=value2,
                value3=value3,
                row_idx=row_idx,
            )

            # Add to queue display
            self.queue_display.table.insert(
                "", 0, values=(task.pdf_path, "Pending"), tags=("pending",)
            )

            # Add to processing queue
            self.pdf_queue.add_task(task)

            # Update status and load next file
            self._update_status("File queued for processing")
            self.load_next_pdf()

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
        self.update_queue_display()
        self._update_status("Completed tasks cleared")

    def _retry_failed(self) -> None:
        """Retry failed tasks in the queue."""
        self.pdf_queue.retry_failed()
        self.update_queue_display()
        self._update_status("Retrying failed tasks")

    def update_queue_display(self) -> None:
        """Update the queue display with current tasks."""
        try:
            with self.pdf_queue.lock:
                tasks = self.pdf_queue.tasks.copy()
            self.queue_display.update_display(tasks)

            # Update queue statistics
            total = len(tasks)
            completed = sum(1 for t in tasks.values() if t.status == "completed")
            failed = sum(1 for t in tasks.values() if t.status == "failed")
            pending = sum(
                1 for t in tasks.values() if t.status in ["pending", "processing"]
            )

            if total > 0:
                self.queue_stats.configure(
                    text=f"Queue: {total} total ({completed} completed, {failed} failed, {pending} pending)"
                )
            else:
                self.queue_stats.configure(text="Queue: 0 total")

        except Exception as e:
            print(f"Error updating queue display: {str(e)}")

    def _periodic_update(self) -> None:
        """Periodically update the queue display."""
        self.update_queue_display()
        self.after(100, self._periodic_update)

    def __del__(self) -> None:
        """Clean up resources when the tab is destroyed."""
        try:
            if hasattr(self, "pdf_queue"):
                self.pdf_queue.stop()
        except:
            pass
