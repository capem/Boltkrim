from __future__ import annotations
from tkinter import Canvas, END, Event as TkEvent, Widget, ttk, StringVar, NSEW, PhotoImage, Toplevel
from tkinter.ttk import Frame, Scrollbar, Label, Button, Style, LabelFrame, Treeview, Notebook
from PIL.ImageTk import PhotoImage as PILPhotoImage
from os import path
from typing import Optional, Any, Dict, List, Callable
from threading import Thread, Lock, Event
from dataclasses import dataclass
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog
import pythoncom
from datetime import datetime
import pandas as pd

# Data Models
@dataclass
class PDFTask:
    pdf_path: str
    value1: str
    value2: str
    value3: str
    status: str = 'pending'  # pending, processing, failed, completed
    error_msg: str = ''

# Queue Management
class ProcessingQueue:
    def __init__(self, config_manager: Any, excel_manager: Any, pdf_manager: Any):
        self.tasks: Dict[str, PDFTask] = {}
        self.lock = Lock()
        self.processing_thread: Optional[Thread] = None
        try:
            self.stop_event = Event()
        except (AttributeError, RuntimeError):
            # Fallback for PyInstaller if Event initialization fails
            import threading
            self.stop_event = threading.Event()
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
            result = {
                'pending': [],
                'processing': [],
                'failed': [],
                'completed': []
            }
            for task in self.tasks.values():
                result[task.status].append(task)
            return result
    
    def clear_completed(self) -> None:
        with self.lock:
            self.tasks = {k: v for k, v in self.tasks.items() 
                         if v.status not in ['completed']}

    def retry_failed(self) -> None:
        with self.lock:
            for task in self.tasks.values():
                if task.status == 'failed':
                    task.status = 'pending'
                    task.error_msg = ''
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
                    if task.status == 'pending':
                        task.status = 'processing'
                        task_to_process = task
                        self._notify_status_change()
                        break
            
            if task_to_process is None:
                self.stop_event.wait(0.1)
                continue
                
            try:
                pythoncom.CoInitialize()
                config = self.config_manager.get_config()
                excel_manager = type(self.excel_manager)()
                excel_manager.load_excel_data(config['excel_file'], config['excel_sheet'])
                
                row_data, row_idx = excel_manager.find_matching_row(
                    config['filter1_column'],
                    config['filter2_column'],
                    config['filter3_column'],
                    task_to_process.value1,
                    task_to_process.value2,
                    task_to_process.value3
                )
                
                if row_data is None:
                    raise Exception("Selected combination not found in Excel sheet")
                
                # Prepare template data with consistent field names
                template_data = {
                    'filter1': row_data[config['filter1_column']],
                    'filter2': row_data[config['filter2_column']],
                    'filter_1': row_data[config['filter1_column']],  # For backward compatibility
                    'filter_2': row_data[config['filter2_column']]   # For backward compatibility
                }
                
                # Add DATE FACTURE from Excel if it exists
                if 'DATE FACTURE' in row_data:
                    # Try to parse the date from Excel
                    try:
                        if isinstance(row_data['DATE FACTURE'], datetime):
                            template_data['DATE FACTURE'] = row_data['DATE FACTURE']
                        else:
                            # Try to parse the date string (add more formats if needed)
                            for date_format in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d']:
                                try:
                                    template_data['DATE FACTURE'] = datetime.strptime(str(row_data['DATE FACTURE']), date_format)
                                    break
                                except ValueError:
                                    continue
                            if 'DATE FACTURE' not in template_data:
                                raise ValueError(f"Could not parse date: {row_data['DATE FACTURE']}")
                    except Exception as e:
                        raise Exception(f"Error processing DATE FACTURE: {str(e)}")
                
                # Process the PDF with template-based naming
                if self.pdf_manager.process_pdf(
                    task_to_process.pdf_path,
                    template_data,
                    config['processed_folder'],
                    config['output_template']
                ):
                    # Get the actual output path for Excel update
                    output_path = self.pdf_manager.generate_output_path(
                        config['output_template'],
                        template_data
                    )
                    
                    excel_manager.update_pdf_link(
                        config['excel_file'],
                        config['excel_sheet'],
                        row_idx,
                        output_path
                    )
                    
                    with self.lock:
                        task_to_process.status = 'completed'
                        self._notify_status_change()
                    
            except Exception as e:
                with self.lock:
                    task_to_process.status = 'failed'
                    task_to_process.error_msg = str(e)
                    self._notify_status_change()
            finally:
                try:
                    if 'excel_manager' in locals():
                        excel_manager.close()
                except:
                    pass
                pythoncom.CoUninitialize()

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
        self.container_frame.grid(row=0, column=0, sticky='nsew')
        self.container_frame.grid_columnconfigure(0, weight=1)
        self.container_frame.grid_rowconfigure(0, weight=1)
        
        # Create canvas with modern styling
        self.canvas = Canvas(
            self.container_frame,
            bg='#f8f9fa',  # Light gray background
            highlightthickness=0,  # Remove border
            width=20,  # Minimum width to prevent collapse
            height=20   # Minimum height to prevent collapse
        )
        self.canvas.grid(row=0, column=0, sticky='nsew')
        
        # Modern scrollbars - always create them to reserve space
        self.h_scrollbar = Scrollbar(
            self.container_frame,
            orient='horizontal',
            command=self.canvas.xview
        )
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        self.v_scrollbar = Scrollbar(
            self.container_frame,
            orient='vertical',
            command=self.canvas.yview
        )
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        
        # Configure canvas scrolling
        self.canvas.configure(
            xscrollcommand=self._on_x_scroll,
            yscrollcommand=self._on_y_scroll
        )
        
        # Initially hide scrollbars but keep their space reserved
        self.h_scrollbar.grid_remove()
        self.v_scrollbar.grid_remove()
        
        # Create a frame for the loading message that won't affect layout
        self.loading_frame = Frame(self.container_frame)
        self.loading_frame.place(relx=0.5, rely=0.5, anchor='center')
        
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
        x1, y1, x2, y2 = self.canvas.bbox("all") if self.canvas.find_all() else (0, 0, 0, 0)
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
        self.canvas.bind('<Enter>', _bind_mousewheel)
        self.canvas.bind('<Leave>', _unbind_mousewheel)
        self.v_scrollbar.bind('<Enter>', _bind_mousewheel)
        self.v_scrollbar.bind('<Leave>', _unbind_mousewheel)
        
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
        self.canvas.create_image(image_x, image_y, anchor='nw', image=self.current_image)
        
        # Update scrollbar visibility
        self._update_scrollbar_visibility()

    def display_pdf(self, pdf_path: str, zoom: float = 1.0, show_loading: bool = True) -> None:
        """Display a PDF file with the specified zoom level."""
        try:
            self.current_pdf = pdf_path
            self.zoom_level = zoom
            
            # Show loading message using place geometry manager
            loading_label = None
            if show_loading:
                loading_label = Label(
                    self.loading_frame,
                    text="Loading PDF...",
                    font=('Segoe UI', 10)
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

        # Create a frame for the table and scrollbar
        table_frame = Frame(self)
        table_frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Setup table with more columns
        columns = ('filename', 'values', 'status', 'time')
        self.table = Treeview(table_frame, columns=columns, show='headings', height=5)
        
        # Configure headings with sort functionality
        self.table.heading('filename', text='File')
        self.table.heading('values', text='Selected Values')
        self.table.heading('status', text='Status')
        self.table.heading('time', text='Time')
        
        # Configure column widths and weights
        self.table.column('filename', width=120, minwidth=80)
        self.table.column('values', width=150, minwidth=120)
        self.table.column('status', width=70, minwidth=70)
        self.table.column('time', width=60, minwidth=60)

        # Add vertical scrollbar
        v_scrollbar = Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=v_scrollbar.set)

        # Add horizontal scrollbar
        h_scrollbar = Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(xscrollcommand=h_scrollbar.set)

        # Grid table and scrollbars
        self.table.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        # Configure status colors and tooltips
        self.table.tag_configure('pending', foreground='black')
        self.table.tag_configure('processing', foreground='blue')
        self.table.tag_configure('completed', foreground='green')
        self.table.tag_configure('failed', foreground='red')

        # Bind tooltip events
        self.table.bind('<Motion>', self._show_tooltip)
        self._tooltip_label = None

        # Create button frame at the bottom
        btn_frame = Frame(self)
        btn_frame.grid(row=1, column=0, sticky='ew', padx=5, pady=(0, 5))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        # Style for modern looking buttons
        style = Style()
        style.configure("Action.TButton", padding=5)
        
        # Add buttons with equal width and spacing
        self.clear_btn = Button(
            btn_frame,
            text="Clear Completed",
            style="Action.TButton",
            width=20  # Reduced from 15
        )
        self.clear_btn.grid(row=0, column=0, padx=(0, 2), sticky='e')
        
        self.retry_btn = Button(
            btn_frame,
            text="Retry Failed",
            style="Action.TButton",
            width=20  # Reduced from 15
        )
        self.retry_btn.grid(row=0, column=1, padx=(2, 0), sticky='w')

    def _show_tooltip(self, event) -> None:
        """Show tooltip with full path when hovering over truncated filename."""
        # Get the item under cursor
        item = self.table.identify_row(event.y)
        if not item:
            self._hide_tooltip()
            return

        # Get the column under cursor
        column = self.table.identify_column(event.x)
        if column != '#1':  # Only show tooltip for filename column
            self._hide_tooltip()
            return

        # Get full path from the item
        values = self.table.item(item)['values']
        if not values:
            self._hide_tooltip()
            return

        full_path = values[0]
        displayed_text = self.table.item(item, 'text')

        # Only show tooltip if text is truncated
        cell_box = self.table.bbox(item, column)
        if not cell_box:
            self._hide_tooltip()
            return

        if self._tooltip_label is None:
            self._tooltip_label = Label(
                self,
                text=full_path,
                background='#ffffe0',
                relief='solid',
                borderwidth=1
            )

        # Position tooltip below the cell
        x = self.table.winfo_rootx() + cell_box[0]
        y = self.table.winfo_rooty() + cell_box[1] + cell_box[3]
        self._tooltip_label.place(x=x, y=y)

    def _hide_tooltip(self) -> None:
        """Hide the tooltip."""
        if self._tooltip_label:
            self._tooltip_label.place_forget()

    def _get_truncated_path(self, path_str: str, max_length: int = 30) -> str:
        """Truncate path while keeping filename."""
        filename = path.basename(path_str)
        if len(filename) <= max_length:
            return filename
        
        # If filename is too long, truncate the middle
        half = (max_length - 3) // 2
        return f"{filename[:half]}...{filename[-half:]}"

    def update_display(self, tasks: Dict[str, PDFTask]) -> None:
        selection = self.table.selection()
        selected_paths = [self.table.item(item)['values'][0] for item in selection]
        
        current_items = {}
        for item in self.table.get_children():
            path_value = self.table.item(item)['values'][0]
            current_items[path_value] = item
        
        for task_path, task in tasks.items():
            # Format the values string
            values_str = f"{task.value1} | {task.value2} | {task.value3}"
            
            # Get current time for processing tasks
            time_str = datetime.now().strftime("%H:%M:%S") if task.status == 'processing' else ""
            
            # Create display values
            display_values = (
                task_path,  # Store full path for tooltip
                values_str,
                task.status.capitalize(),
                time_str
            )
            
            if task_path in current_items:
                # Update all columns with new values
                for idx, col in enumerate(['filename', 'values', 'status', 'time']):
                    self.table.set(current_items[task_path], column=col, value=display_values[idx])
                self.table.item(current_items[task_path],
                              text=self._get_truncated_path(task_path),
                              tags=(task.status,))
                current_items.pop(task_path)
            else:
                item = self.table.insert('', 'end',
                                       text=self._get_truncated_path(task_path),
                                       values=display_values,
                                       tags=(task.status,))
                if task_path in selected_paths:
                    self.table.selection_add(item)
        
        for item_id in current_items.values():
            self.table.delete(item_id)

class ProcessingTab(Frame):
    """A modernized tab for processing PDF files with Excel data integration."""
    
    def __init__(self, master: Widget, config_manager: Any,
                 excel_manager: Any, pdf_manager: Any) -> None:
        super().__init__(master)
        self.master = master
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        
        self.pdf_queue = ProcessingQueue(config_manager, excel_manager, pdf_manager)
        self.current_pdf: Optional[str] = None
        self.all_values2: List[str] = []
        
        # Configure styles
        self._setup_styles()
        
        # Setup main layout
        self.setup_ui()
        self.update_queue_display()
        self.after(100, self._periodic_update)
        
        # Register for config changes
        if hasattr(config_manager, 'add_change_callback'):
            config_manager.add_change_callback(self.handle_config_change)

    def _setup_styles(self) -> None:
        """Configure custom styles for the interface."""
        style = Style()
        
        # Configure main theme settings
        style.configure(".", font=('Segoe UI', 10))
        style.configure("Title.TLabel", font=('Segoe UI', 12, 'bold'))
        style.configure("Header.TLabel", font=('Segoe UI', 11, 'bold'))
        
        # Configure button styles
        style.configure("Primary.TButton",
                       padding=10,
                       font=('Segoe UI', 10, 'bold'),
                       background="#007bff")
        style.configure("Secondary.TButton",
                       padding=10,
                                 font=('Segoe UI', 10))
        style.configure("Success.TButton",
                       padding=10,
                       font=('Segoe UI', 10, 'bold'),
                       background="#28a745")
        
        # Configure frame styles
        style.configure("Card.TFrame",
                       background="#ffffff",
                       relief="solid",
                       borderwidth=1)
        
        # Configure Treeview
        style.configure("Treeview",
                       font=('Segoe UI', 10),
                       rowheight=25)
        style.configure("Treeview.Heading",
                       font=('Segoe UI', 10, 'bold'))

    def setup_ui(self) -> None:
        """Setup the main user interface with a modern, clean layout."""
        self.configure(padding=20)
        
        # Configure grid weights for the main frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Create main container with three columns
        main_container = Frame(self)
        main_container.grid(row=0, column=0, sticky='nsew')
        
        # Configure column weights for the container
        main_container.grid_columnconfigure(0, weight=1)  # Left panel (20%)
        main_container.grid_columnconfigure(1, weight=8)  # Center panel (70%)
        main_container.grid_columnconfigure(2, weight=1)  # Right panel (10%)
        main_container.grid_rowconfigure(0, weight=1)
        
        # Left Panel (10% width)
        left_panel = self._create_left_panel(main_container)
        left_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
        # Center Panel (80% width)
        center_panel = self._create_center_panel(main_container)
        center_panel.grid(row=0, column=1, sticky='nsew', padx=10)
        
        # Right Panel (10% width)
        right_panel = self._create_right_panel(main_container)
        right_panel.grid(row=0, column=2, sticky='nsew', padx=(10, 0))

    def _create_left_panel(self, parent: Widget) -> Frame:
        """Create the left panel containing file information and queue."""
        panel = Frame(parent)
        
        # File Information Section
        info_frame = LabelFrame(panel, text="File Information", padding=10)
        info_frame.pack(fill='x', pady=(0, 10))
        
        # Create a fixed-width container for the file info
        file_info_container = Frame(info_frame)  # Removed fixed width from container
        file_info_container.pack(fill='x', pady=5)
        
        self.file_info = Label(file_info_container, text="No file loaded",
                             style="Header.TLabel", width=25)  # Set fixed width on label instead
        self.file_info.pack(fill='x', pady=5)
        
        # Create tooltip for full filename
        self.file_info_tooltip = None
        def show_tooltip(event):
            if self.file_info_tooltip or self.file_info['text'] == "No file loaded":
                return
            text = self.file_info['text']
            if len(text) > 50:  # Only show tooltip for long filenames
                x, y, _, _ = self.file_info.bbox("insert")
                x += self.file_info.winfo_rootx() + 25
                y += self.file_info.winfo_rooty() + 25
                self.file_info_tooltip = Toplevel(self.file_info)
                self.file_info_tooltip.wm_overrideredirect(True)
                self.file_info_tooltip.wm_geometry(f"+{x}+{y}")
                tooltip_label = Label(self.file_info_tooltip, text=text, 
                                   justify='left', background="#ffffe0", 
                                   relief='solid', borderwidth=1)
                tooltip_label.pack()

        def hide_tooltip(event):
            if self.file_info_tooltip:
                self.file_info_tooltip.destroy()
                self.file_info_tooltip = None

        self.file_info.bind('<Enter>', show_tooltip)
        self.file_info.bind('<Leave>', hide_tooltip)
        
        self.status_var = StringVar(value="Status: Ready")
        status_label = Label(info_frame, textvariable=self.status_var)
        status_label.pack(fill='x')
        
        # Processing Queue Section
        queue_frame = LabelFrame(panel, text="Processing Queue", padding=10)
        queue_frame.pack(fill='both', expand=True)
        
        self.queue_display = QueueDisplay(queue_frame)
        self.queue_display.pack(fill='both', expand=True)
        
        return panel

    def _create_center_panel(self, parent: Widget) -> Frame:
        """Create the center panel containing the PDF viewer."""
        panel = Frame(parent)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(0, weight=1)
        
        # PDF Viewer Section
        viewer_frame = LabelFrame(panel, text="PDF Viewer", padding=10)
        viewer_frame.grid(row=0, column=0, sticky='nsew')
        viewer_frame.grid_columnconfigure(0, weight=1)
        viewer_frame.grid_rowconfigure(0, weight=1)
        
        self.pdf_viewer = PDFViewer(viewer_frame, self.pdf_manager)
        self.pdf_viewer.grid(row=0, column=0, sticky='nsew')
        
        # Viewer Controls
        controls_frame = Frame(viewer_frame)
        controls_frame.grid(row=1, column=0, sticky='ew', pady=(10, 0))
        
        # Zoom Controls
        zoom_frame = Frame(controls_frame)
        zoom_frame.pack(side='left')
        
        Button(zoom_frame, text="−", width=3,
                  command=self.pdf_viewer.zoom_out).pack(side='left', padx=2)
        self.zoom_label = Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side='left', padx=5)
        Button(zoom_frame, text="+", width=3,
                  command=self.pdf_viewer.zoom_in).pack(side='left', padx=2)
        
        # Rotation Controls
        rotation_frame = Frame(controls_frame)
        rotation_frame.pack(side='right')
        
        Button(rotation_frame, text="↶", width=3,
                  command=self.rotate_counterclockwise).pack(side='left', padx=2)
        self.rotation_label = Label(rotation_frame, text="0°", width=6)
        self.rotation_label.pack(side='left', padx=5)
        Button(rotation_frame, text="↷", width=3,
                  command=self.rotate_clockwise).pack(side='left', padx=2)

        return panel

    def _create_right_panel(self, parent: Widget) -> Frame:
        """Create the right panel containing filters and actions."""
        panel = Frame(parent)
        
        # Filters Frame (replacing notebook)
        filters_frame = LabelFrame(panel, text="Filters", padding=10)
        filters_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self._setup_filters(filters_frame)
        
        # Action Buttons
        actions_frame = Frame(panel)
        actions_frame.pack(fill='x', pady=(10, 0))
        
        self.confirm_button = Button(
            actions_frame,
            text="Process File (Enter)",
            command=self.process_current_file,
            style="Success.TButton"
        )
        self.confirm_button.pack(fill='x', pady=(0, 5))
        
        self.skip_button = Button(
            actions_frame,
            text="Next File (Ctrl+N)",
            command=self.load_next_pdf,
            style="Primary.TButton"
        )
        self.skip_button.pack(fill='x')
        
        return panel

    def _setup_filters(self, parent: Widget) -> None:
        """Setup the filter controls with improved styling."""
        self.filter1_label = Label(parent, text="",
                                 style="Header.TLabel")
        self.filter1_label.pack(pady=(0, 5))
        self.filter1_frame = FuzzySearchFrame(parent, width=30,
                                            identifier='processing_filter1')
        self.filter1_frame.pack(fill='x', pady=(0, 15))
        
        self.filter2_label = Label(parent, text="",
                                 style="Header.TLabel")
        self.filter2_label.pack(pady=(0, 5))
        self.filter2_frame = FuzzySearchFrame(parent, width=30,
                                            identifier='processing_filter2')
        self.filter2_frame.pack(fill='x', pady=(0, 15))

        self.filter3_label = Label(parent, text="",
                                 style="Header.TLabel")
        self.filter3_label.pack(pady=(0, 5))
        self.filter3_frame = FuzzySearchFrame(parent, width=30,
                                            identifier='processing_filter3')
        self.filter3_frame.pack(fill='x')
        
        # Bind events
        self.filter1_frame.bind('<<ValueSelected>>', lambda e: self.on_filter1_select())
        self.filter2_frame.bind('<<ValueSelected>>', lambda e: self.on_filter2_select())
        self.filter3_frame.bind('<<ValueSelected>>', lambda e: self.update_confirm_button())

        # Bind keyboard navigation
        self._bind_keyboard_shortcuts()

    def _bind_keyboard_shortcuts(self) -> None:
        """Bind keyboard shortcuts for improved navigation and accessibility."""
        # Tab navigation between filters
        self.filter1_frame.entry.bind('<Tab>', self._handle_filter1_tab)
        self.filter2_frame.entry.bind('<Tab>', self._handle_filter2_tab)
        self.filter3_frame.entry.bind('<Tab>', self._handle_filter3_tab)
        self.filter1_frame.listbox.bind('<Tab>', self._handle_filter1_tab)
        self.filter2_frame.listbox.bind('<Tab>', self._handle_filter2_tab)
        self.filter3_frame.listbox.bind('<Tab>', self._handle_filter3_tab)
        
        # Global shortcuts
        shortcuts = {
            '<Return>': self._handle_return_key,
            '<Control-n>': lambda e: self.load_next_pdf(),
            '<Control-N>': lambda e: self.load_next_pdf(),
            '<Control-plus>': lambda e: self.pdf_viewer.zoom_in(),
            '<Control-minus>': lambda e: self.pdf_viewer.zoom_out(),
            '<Control-r>': lambda e: self.rotate_clockwise(),
            '<Control-R>': lambda e: self.rotate_counterclockwise()
        }
        
        # Bind shortcuts recursively to all widgets
        def _bind_recursive(widget: Widget) -> None:
            for key, callback in shortcuts.items():
                widget.bind(key, callback)
            for child in widget.winfo_children():
                _bind_recursive(child)
        
        _bind_recursive(self)
        
        # Also bind to the main frame
        for key, callback in shortcuts.items():
            self.bind_all(key, callback)

    def _handle_filter1_tab(self, event: Event) -> str:
        """Handle tab key in filter1 to move focus to filter2."""
        if self.filter1_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item and move to filter2
            if self.filter1_frame.listbox.size() > 0:
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
                self.filter3_frame.listbox.selection_clear(0, END)
                self.filter3_frame.listbox.selection_set(0)
                self.filter3_frame._on_select(None)
        self.confirm_button.focus_set()
        return "break"

    def _handle_return_key(self, event: Event) -> str:
        """Handle Return key press to process the current file."""
        if str(self.confirm_button['state']) != 'disabled':
            self.process_current_file()
        return "break"

    def handle_config_change(self) -> None:
        """Handle configuration changes by reloading the current PDF if one is loaded."""
        if self.current_pdf:
            self.pdf_viewer.display_pdf(self.current_pdf)
            self.update_queue_display()

    def load_excel_data(self) -> None:
        try:
            config = self.config_manager.get_config()
            if not all([config['excel_file'], config['excel_sheet'],
                       config['filter1_column'], config['filter2_column'],
                       config['filter3_column']]):
                print("Missing configuration values")
                return
                
            self.excel_manager.load_excel_data(config['excel_file'],
                                             config['excel_sheet'])
            
            self.filter1_label['text'] = config['filter1_column']
            self.filter2_label['text'] = config['filter2_column']
            self.filter3_label['text'] = config['filter3_column']
            
            df = self.excel_manager.excel_data
            
            # Convert all values to strings to ensure consistent type handling
            def safe_convert_to_str(val):
                if pd.isna(val):  # Handle NaN/None values
                    return ""
                return str(val).strip()
            
            # Convert column values to strings
            self.all_values1 = sorted(df[config['filter1_column']].astype(str).unique().tolist())
            self.all_values2 = sorted(df[config['filter2_column']].astype(str).unique().tolist())
            self.all_values3 = sorted(df[config['filter3_column']].astype(str).unique().tolist())
            
            # Strip whitespace and ensure string type
            self.all_values1 = [safe_convert_to_str(x) for x in self.all_values1]
            self.all_values2 = [safe_convert_to_str(x) for x in self.all_values2]
            self.all_values3 = [safe_convert_to_str(x) for x in self.all_values3]
            
            self.filter1_frame.set_values(self.all_values1)
            self.filter2_frame.set_values(self.all_values2)
            self.filter3_frame.set_values(self.all_values3)
            
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in load_excel_data:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")

    def on_filter1_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value = str(self.filter1_frame.get()).strip()
                
                df = self.excel_manager.excel_data
                # Convert column to string for comparison
                df[config['filter1_column']] = df[config['filter1_column']].astype(str)
                filtered_df = df[df[config['filter1_column']].str.strip() == selected_value]
                
                filtered_values2 = sorted(filtered_df[config['filter2_column']].astype(str).unique().tolist())
                filtered_values2 = [str(x).strip() for x in filtered_values2]
                self.filter2_frame.set_values(filtered_values2)
                
                # Clear filter3 since it depends on filter2
                self.filter3_frame.set('')
                self.filter3_frame.set_values(self.all_values3)
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in on_filter1_select:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def on_filter2_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value1 = str(self.filter1_frame.get()).strip()
                selected_value2 = str(self.filter2_frame.get()).strip()
                
                df = self.excel_manager.excel_data
                # Convert columns to string for comparison
                df[config['filter1_column']] = df[config['filter1_column']].astype(str)
                df[config['filter2_column']] = df[config['filter2_column']].astype(str)
                
                filtered_df = df[
                    (df[config['filter1_column']].str.strip() == selected_value1) &
                    (df[config['filter2_column']].str.strip() == selected_value2)
                ]
                
                filtered_values3 = sorted(filtered_df[config['filter3_column']].astype(str).unique().tolist())
                filtered_values3 = [str(x).strip() for x in filtered_values3]
                self.filter3_frame.set_values(filtered_values3)
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in on_filter2_select:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def update_confirm_button(self) -> None:
        """Update the confirm button state based on filter selections."""
        if self.filter1_frame.get() and self.filter2_frame.get() and self.filter3_frame.get():
            self.confirm_button.state(['!disabled'])
            self.status_var.set("Status: Ready to process")
        else:
            self.confirm_button.state(['disabled'])
            self.status_var.set("Status: Select all filters")

    def rotate_clockwise(self) -> None:
        """Rotate the PDF view clockwise."""
        self.pdf_manager.rotate_page(clockwise=True)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(self.current_pdf, self.pdf_viewer.zoom_level, show_loading=False)

    def rotate_counterclockwise(self) -> None:
        """Rotate the PDF view counterclockwise."""
        self.pdf_manager.rotate_page(clockwise=False)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(self.current_pdf, self.pdf_viewer.zoom_level, show_loading=False)

    def load_next_pdf(self) -> None:
        """Load the next PDF file from the source folder."""
        try:
            config = self.config_manager.get_config()
            if not config['source_folder']:
                self.status_var.set("Status: Source folder not configured")
                return
                
            # Clear current display if no PDF is loaded
            if not path.exists(config['source_folder']):
                self.current_pdf = None
                self.file_info['text'] = "Source folder not found"
                self.status_var.set("Status: Source folder does not exist")
                ErrorDialog(self, "Error", f"Source folder not found: {config['source_folder']}")
                return
                
            next_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
            if next_pdf:
                self.current_pdf = next_pdf
                self.file_info['text'] = path.basename(next_pdf)
                self.pdf_viewer.display_pdf(next_pdf, 1.0)
                self.rotation_label.config(text="0°")
                self.zoom_label.config(text="100%")
                self.filter1_frame.focus_set()
                self.status_var.set("Status: New file loaded")
            else:
                self.current_pdf = None
                self.file_info['text'] = "No PDF files found"
                self.status_var.set("Status: No files to process")
                # Clear the PDF viewer
                if hasattr(self.pdf_viewer, 'canvas'):
                    self.pdf_viewer.canvas.delete("all")
                # Disable the confirm button since there's no file to process
                self.confirm_button.state(['disabled'])
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in load_next_pdf:")
            print(traceback.format_exc())
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")

    def process_current_file(self) -> None:
        """Process the current PDF file with selected filters."""
        if not self.current_pdf:
            self.status_var.set("Status: No file loaded")
            ErrorDialog(self, "Error", "No PDF file loaded")
            return
            
        if not path.exists(self.current_pdf):
            self.status_var.set("Status: File no longer exists")
            ErrorDialog(self, "Error", f"File no longer exists: {self.current_pdf}")
            self.load_next_pdf()  # Try to load the next file
            return
            
        try:
            value1 = self.filter1_frame.get()
            value2 = self.filter2_frame.get()
            value3 = self.filter3_frame.get()
            
            if not value1 or not value2 or not value3:
                self.status_var.set("Status: Select all filters")
                return

            # Create task
            task = PDFTask(
                pdf_path=self.current_pdf,
                value1=value1,
                value2=value2,
                value3=value3
            )
            
            # Add to queue display
            self.queue_display.table.insert('', 'end',
                                          values=(task.pdf_path, 'Pending'),
                                          tags=('pending',))
            
            # Add to processing queue
            self.pdf_queue.add_task(task)
            
            # Update status and load next file
            self.status_var.set("Status: File queued for processing")
            self.load_next_pdf()
            
            # Clear filters
            self.filter1_frame.set('')
            self.filter2_frame.set('')
            self.filter3_frame.set('')
            self.filter2_frame.set_values(self.all_values2)
            self.filter3_frame.set_values(self.all_values3)
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in process_current_file:")
            print(traceback.format_exc())
            self.status_var.set("Status: Processing error")
            ErrorDialog(self, "Error", str(e))

    def _show_error_details(self, event: TkEvent) -> None:
        """Show error details for failed tasks."""
        selection = self.queue_display.table.selection()
        if not selection:
            return
            
        item = selection[0]
        task_path = self.queue_display.table.item(item)['values'][0]
        
        with self.pdf_queue.lock:
            task = self.pdf_queue.tasks.get(task_path)
            if task and task.status == 'failed' and task.error_msg:
                ErrorDialog(self, "Processing Error",
                          f"Error processing {path.basename(task_path)}:\n{task.error_msg}")

    def _clear_completed(self) -> None:
        """Clear completed tasks from the queue."""
        self.pdf_queue.clear_completed()
        self.update_queue_display()
        self.status_var.set("Status: Completed tasks cleared")

    def _retry_failed(self) -> None:
        """Retry failed tasks in the queue."""
        self.pdf_queue.retry_failed()
        self.update_queue_display()
        self.status_var.set("Status: Retrying failed tasks")

    def update_queue_display(self) -> None:
        """Update the queue display with current tasks."""
        try:
            with self.pdf_queue.lock:
                tasks = self.pdf_queue.tasks.copy()
            self.queue_display.update_display(tasks)
            
            # Update status with queue statistics
            total = len(tasks)
            completed = sum(1 for t in tasks.values() if t.status == 'completed')
            failed = sum(1 for t in tasks.values() if t.status == 'failed')
            pending = sum(1 for t in tasks.values() if t.status in ['pending', 'processing'])
            
            if total > 0:
                self.status_var.set(
                    f"Status: {completed} completed, {failed} failed, {pending} pending"
                )
            
        except Exception as e:
            print(f"Error updating queue display: {str(e)}")

    def _periodic_update(self) -> None:
        """Periodically update the queue display."""
        self.update_queue_display()
        self.after(100, self._periodic_update)

    def __del__(self) -> None:
        """Clean up resources when the tab is destroyed."""
        try:
            if hasattr(self, 'pdf_queue'):
                self.pdf_queue.stop()
        except:
            pass
