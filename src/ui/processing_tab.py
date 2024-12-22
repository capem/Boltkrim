from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import os
from typing import Optional, Any, Dict, List, Tuple
from queue import Queue
from threading import Thread, Event, Lock
from dataclasses import dataclass
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog
import pythoncom

@dataclass
class PDFTask:
    pdf_path: str
    value1: str
    value2: str
    status: str = 'pending'  # pending, processing, failed, completed
    error_msg: str = ''

class ProcessingQueue:
    def __init__(self, config_manager: Any, excel_manager: Any, pdf_manager: Any):
        self.tasks: Dict[str, PDFTask] = {}
        self.lock = Lock()
        self.processing_thread: Optional[Thread] = None
        self.stop_event = Event()
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        self._callbacks: List[callable] = []  # Callbacks for status changes

    def add_task(self, task: PDFTask) -> None:
        with self.lock:
            self.tasks[task.pdf_path] = task
        self._notify_status_change()
        self._ensure_processing()

    def _notify_status_change(self) -> None:
        """Notify listeners of status changes."""
        for callback in self._callbacks:
            try:
                callback()
            except Exception:
                pass

    def _ensure_processing(self) -> None:
        """Ensure processing thread is running."""
        if self.processing_thread is None or not self.processing_thread.is_alive():
            self.stop_event.clear()
            self.processing_thread = Thread(target=self._process_queue, daemon=True)
            self.processing_thread.start()
    
    def get_task_status(self) -> Dict[str, List[PDFTask]]:
        """Get tasks grouped by status."""
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
        """Remove completed tasks from queue."""
        with self.lock:
            self.tasks = {k: v for k, v in self.tasks.items() 
                         if v.status not in ['completed']}

    def retry_failed(self) -> None:
        """Retry all failed tasks."""
        with self.lock:
            for task in self.tasks.values():
                if task.status == 'failed':
                    task.status = 'pending'
                    task.error_msg = ''
        self._ensure_processing()

    def stop(self) -> None:
        """Stop processing queue."""
        self.stop_event.set()
        if self.processing_thread:
            self.processing_thread.join(timeout=1)

    def _process_queue(self) -> None:
        """Process tasks in the queue."""
        while not self.stop_event.is_set():
            task_to_process = None
            
            # Find next pending task
            with self.lock:
                for task in self.tasks.values():
                    if task.status == 'pending':
                        task.status = 'processing'
                        task_to_process = task
                        self._notify_status_change()
                        break
            
            if task_to_process is None:
                # No pending tasks, sleep for a bit
                self.stop_event.wait(0.1)  # Shorter sleep for responsiveness
                continue
                
            try:
                # Initialize COM for Excel operations
                pythoncom.CoInitialize()
                
                config = self.config_manager.get_config()
                
                # Create a new Excel manager instance for this task
                excel_manager = type(self.excel_manager)()  # Create without args
                excel_manager.load_excel_data(config['excel_file'], config['excel_sheet'])
                
                row_data, row_idx = excel_manager.find_matching_row(
                    config['filter1_column'],
                    config['filter2_column'],
                    task_to_process.value1,
                    task_to_process.value2
                )
                
                if row_data is None:
                    raise Exception("Selected combination not found in Excel sheet")
                    
                new_filename = f"{row_data[config['filter1_column']]} - {row_data[config['filter2_column']]}"
                
                # Clean filename
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    new_filename = new_filename.replace(char, '_')
                    
                if not new_filename.lower().endswith('.pdf'):
                    new_filename += '.pdf'
                    
                new_filepath = os.path.join(config['processed_folder'], new_filename)
                
                if self.pdf_manager.process_pdf(task_to_process.pdf_path, new_filepath, 
                                              config['processed_folder']):
                    excel_manager.update_pdf_link(
                        config['excel_file'],
                        config['excel_sheet'],
                        row_idx,
                        new_filepath
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
                    # Clean up Excel manager
                    if 'excel_manager' in locals():
                        excel_manager.close()
                except:
                    pass
                pythoncom.CoUninitialize()

class ProcessingTab(ttk.Frame):
    """A tab for processing PDF files with Excel data integration.
    
    This tab provides functionality to:
    - View PDF files from a source directory
    - Apply filters based on Excel data
    - Process files with new names based on filter selections
    - Navigate through files with keyboard shortcuts
    
    Attributes:
        config_manager: Manages application configuration
        excel_manager: Handles Excel file operations
        pdf_manager: Handles PDF file operations
    """
    
    # Constants
    ZOOM_STEP: float = 0.2
    MIN_ZOOM: float = 0.2
    MAX_ZOOM: float = 3.0
    INITIAL_ZOOM: float = 1.0
    
    def __init__(self, master: tk.Widget, config_manager: Any, 
                 excel_manager: Any, pdf_manager: Any) -> None:
        """Initialize the ProcessingTab.
        
        Args:
            master: Parent widget
            config_manager: Configuration management instance
            excel_manager: Excel operations instance
            pdf_manager: PDF operations instance
        """
        super().__init__(master)
        self.master = master
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        
        # Initialize state
        self.pdf_queue = ProcessingQueue(config_manager, excel_manager, pdf_manager)
        self.current_zoom = self.INITIAL_ZOOM
        self.current_pdf: Optional[str] = None
        self.all_values2: List[str] = []
        
        # Setup UI
        self.setup_ui()
        
        # Start status updates
        self.update_queue_display()
        self.after(1000, self._periodic_update)
        
    def setup_ui(self) -> None:
        """Create and setup the processing UI components."""
        self._setup_styles()
        container = self._create_main_container()
        self._setup_header(container)
        content_frame = self._setup_content_area(container)
        self._setup_pdf_viewer(content_frame)
        self._setup_controls(content_frame)
        self._bind_keyboard_shortcuts()
        
    def _setup_styles(self) -> None:
        """Configure ttk styles for the tab."""
        style = ttk.Style()
        style.configure("Action.TButton", padding=5)
        style.configure("Zoom.TButton", padding=2)
        
    def _create_main_container(self) -> ttk.Frame:
        """Create the main container frame with padding."""
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=20, pady=20)
        return container
        
    def _setup_header(self, container: ttk.Frame) -> None:
        """Setup the header area with file information."""
        header_frame = ttk.Frame(container)
        header_frame.pack(fill='x', pady=(0, 10))
        
        self.file_info = ttk.Label(header_frame, text="No file loaded", 
                                 font=('Segoe UI', 10))
        self.file_info.pack(side='left')
        
    def _setup_content_area(self, container: ttk.Frame) -> ttk.Frame:
        """Setup the main content area with PDF viewer and controls."""
        content_frame = ttk.Frame(container)
        content_frame.pack(fill='both', expand=True)
        return content_frame
        
    def _setup_pdf_viewer(self, content_frame: ttk.Frame) -> None:
        """Setup the PDF viewer area with scrollbars."""
        viewer_frame = ttk.LabelFrame(content_frame, text="PDF Viewer", padding=10)
        viewer_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        self.pdf_frame = ttk.Frame(viewer_frame)
        self.pdf_frame.pack(fill='both', expand=True)
        
    def _setup_controls(self, content_frame: ttk.Frame) -> None:
        """Setup the controls panel with zoom, filters, and action buttons."""
        controls_frame = ttk.LabelFrame(content_frame, text="Controls", padding=10)
        controls_frame.configure(width=250)
        controls_frame.pack(side='right', fill='y')
        controls_frame.pack_propagate(False)
        
        self._setup_zoom_controls(controls_frame)
        self._setup_queue_status(controls_frame)
        self._setup_filters(controls_frame)
        self._setup_action_buttons(controls_frame)
        
    def _setup_zoom_controls(self, controls_frame: ttk.Frame) -> None:
        """Setup zoom control buttons and label."""
        zoom_frame = ttk.Frame(controls_frame)
        zoom_frame.pack(fill='x', pady=(0, 15))
        
        # Zoom controls
        ttk.Label(zoom_frame, text="Zoom:").pack(side='left')
        ttk.Button(zoom_frame, text="−", width=3, style="Zoom.TButton",
                  command=self.zoom_out).pack(side='left', padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side='left', padx=5)
        ttk.Button(zoom_frame, text="+", width=3, style="Zoom.TButton",
                  command=self.zoom_in).pack(side='left', padx=2)
        
        # Rotation controls
        rotation_frame = ttk.Frame(controls_frame)
        rotation_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(rotation_frame, text="Rotate:").pack(side='left')
        ttk.Button(rotation_frame, text="↶", width=3, style="Zoom.TButton",
                  command=self.rotate_counterclockwise).pack(side='left', padx=2)
        self.rotation_label = ttk.Label(rotation_frame, text="0°", width=6)
        self.rotation_label.pack(side='left', padx=5)
        ttk.Button(rotation_frame, text="↷", width=3, style="Zoom.TButton",
                  command=self.rotate_clockwise).pack(side='left', padx=2)
        
    def _setup_queue_status(self, controls_frame: ttk.Frame) -> None:
        """Setup the queue status panel with a table view."""
        status_frame = ttk.LabelFrame(controls_frame, text="Processing Queue", padding=10)
        status_frame.pack(fill='x', pady=(0, 15))

        # Create table
        columns = ('filename', 'status')
        self.queue_table = ttk.Treeview(status_frame, columns=columns, show='headings', height=5)
        
        # Configure columns
        self.queue_table.heading('filename', text='File')
        self.queue_table.heading('status', text='Status')
        
        # Column widths (adjust based on control frame width)
        self.queue_table.column('filename', width=150)
        self.queue_table.column('status', width=80)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.queue_table.yview)
        self.queue_table.configure(yscrollcommand=scrollbar.set)

        # Pack table and scrollbar
        self.queue_table.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Status color tags
        self.queue_table.tag_configure('pending', foreground='black')
        self.queue_table.tag_configure('processing', foreground='blue')
        self.queue_table.tag_configure('completed', foreground='green')
        self.queue_table.tag_configure('failed', foreground='red')

        # Action buttons
        btn_frame = ttk.Frame(status_frame)
        btn_frame.pack(fill='x', pady=(5, 0))
        
        ttk.Button(btn_frame, text="Clear Completed", 
                  command=self._clear_completed).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Retry Failed",
                  command=self._retry_failed).pack(side='right', padx=2)

        # Bind double-click to show error message for failed items
        self.queue_table.bind('<Double-1>', self._show_error_details)

    def _show_error_details(self, event) -> None:
        """Show error details for failed tasks when double-clicked."""
        item = self.queue_table.selection()[0]
        task_path = self.queue_table.item(item)['values'][0]
        
        with self.pdf_queue.lock:
            task = self.pdf_queue.tasks.get(task_path)
            if task and task.status == 'failed' and task.error_msg:
                ErrorDialog(self, "Processing Error", 
                          f"Error processing {os.path.basename(task_path)}:\n{task.error_msg}")

    def _clear_completed(self) -> None:
        """Clear completed tasks from queue."""
        self.pdf_queue.clear_completed()
        self.update_queue_display()

    def _retry_failed(self) -> None:
        """Retry failed tasks."""
        self.pdf_queue.retry_failed()
        self.update_queue_display()

    def _setup_filters(self, controls_frame: ttk.Frame) -> None:
        """Setup filter controls with labels and fuzzy search frames."""
        filters_frame = ttk.LabelFrame(controls_frame, text="Filters", padding=10)
        filters_frame.pack(fill='x', pady=(0, 15))
        
        # First filter
        self.filter1_label = ttk.Label(filters_frame, text="", 
                                     font=('Segoe UI', 9, 'bold'))
        self.filter1_label.pack(pady=(0, 5))
        self.filter1_frame = FuzzySearchFrame(filters_frame, width=30, 
                                            identifier='processing_filter1')
        self.filter1_frame.pack(fill='x', pady=(0, 10))
        
        # Second filter
        self.filter2_label = ttk.Label(filters_frame, text="", 
                                     font=('Segoe UI', 9, 'bold'))
        self.filter2_label.pack(pady=(0, 5))
        self.filter2_frame = FuzzySearchFrame(filters_frame, width=30, 
                                            identifier='processing_filter2')
        self.filter2_frame.pack(fill='x')
        
        # Bind filter events
        self.filter1_frame.bind('<<ValueSelected>>', 
                              lambda e: self.on_filter1_select())
        self.filter2_frame.bind('<<ValueSelected>>', 
                              lambda e: self.update_confirm_button())
        
    def _setup_action_buttons(self, controls_frame: ttk.Frame) -> None:
        """Setup action buttons for processing and skipping files."""
        actions_frame = ttk.Frame(controls_frame)
        actions_frame.pack(fill='x', pady=(0, 10))
        
        self.confirm_button = ttk.Button(
            actions_frame, 
            text="Process File (Enter)",
            command=self.process_current_file, 
            style="Action.TButton"
        )
        self.confirm_button.pack(fill='x', pady=(0, 5))
        
        self.skip_button = ttk.Button(
            actions_frame, 
            text="Skip File (→)",
            command=self.load_next_pdf, 
            style="Action.TButton"
        )
        self.skip_button.pack(fill='x')
        
    def _bind_keyboard_shortcuts(self) -> None:
        """Bind keyboard shortcuts for common actions."""
        self.bind_all('<Return>', lambda e: self.handle_return_key(e))
        self.bind_all('<Right>', lambda e: self.load_next_pdf())
        self.bind_all('<Control-plus>', lambda e: self.zoom_in())
        self.bind_all('<Control-minus>', lambda e: self.zoom_out())
        self.bind_all('<Control-r>', lambda e: self.rotate_clockwise())
        self.bind_all('<Control-Shift-R>', lambda e: self.rotate_counterclockwise())
        
    def handle_return_key(self, event: tk.Event) -> str:
        """Handle Return key press for processing current file.
        
        Args:
            event: Keyboard event information
            
        Returns:
            str: 'break' to prevent event propagation
        """
        if str(self.confirm_button['state']) != 'disabled':
            self.process_current_file()
        return "break"
        
    def load_excel_data(self) -> None:
        """Load and prepare Excel data for filtering."""
        try:
            config = self.config_manager.get_config()
            if not all([config['excel_file'], config['excel_sheet'],
                       config['filter1_column'], config['filter2_column']]):
                print("Missing configuration values")
                return
                
            self.excel_manager.load_excel_data(config['excel_file'], 
                                             config['excel_sheet'])
            
            # Update filter labels and values
            self.filter1_label['text'] = config['filter1_column']
            self.filter2_label['text'] = config['filter2_column']
            
            df = self.excel_manager.excel_data
            self.all_values1 = sorted(df[config['filter1_column']].unique().tolist())
            self.all_values2 = sorted(df[config['filter2_column']].unique().tolist())
            
            self.filter1_frame.set_values(self.all_values1)
            self.filter2_frame.set_values(self.all_values2)
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")
            
    def on_filter1_select(self) -> None:
        """Handle selection in first filter and update second filter options."""
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value = self.filter1_frame.get()
                
                df = self.excel_manager.excel_data
                filtered_df = df[df[config['filter1_column']] == selected_value]
                
                filtered_values = sorted(filtered_df[config['filter2_column']].unique().tolist())
                self.filter2_frame.set_values(filtered_values)
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")
            
    def update_confirm_button(self) -> None:
        """Update confirm button state based on filter selections."""
        if self.filter1_frame.get() and self.filter2_frame.get():
            self.confirm_button.state(['!disabled'])
        else:
            self.confirm_button.state(['disabled'])
            
    def zoom_in(self) -> None:
        """Increase zoom level and refresh display."""
        self.zoom_level = min(self.MAX_ZOOM, self.zoom_level + self.ZOOM_STEP)
        self.display_pdf()
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        
    def zoom_out(self) -> None:
        """Decrease zoom level and refresh display."""
        self.zoom_level = max(self.MIN_ZOOM, self.zoom_level - self.ZOOM_STEP)
        self.display_pdf()
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        
    def rotate_clockwise(self) -> None:
        """Rotate the PDF clockwise by 90 degrees."""
        self.pdf_manager.rotate_page(clockwise=True)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.display_pdf()
        
    def rotate_counterclockwise(self) -> None:
        """Rotate the PDF counterclockwise by 90 degrees."""
        self.pdf_manager.rotate_page(clockwise=False)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.display_pdf()
        
    def load_next_pdf(self) -> None:
        """Load the next PDF file from the source folder."""
        try:
            config = self.config_manager.get_config()
            if not config['source_folder']:
                return
                
            next_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
            if next_pdf:
                self.current_pdf = next_pdf
                self.file_info['text'] = os.path.basename(next_pdf)
                self.zoom_level = self.INITIAL_ZOOM
                self.zoom_label.config(text="100%")
                self.rotation_label.config(text="0°")  # Reset rotation label
                self.display_pdf()
                self.filter1_frame.focus_set()
            else:
                self.file_info['text'] = "No PDF files found"
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")
            
    def display_pdf(self) -> None:
        """Display the current PDF with zoom level and scrollbars."""
        try:
            # Clear previous display
            for widget in self.pdf_frame.winfo_children():
                widget.destroy()
                
            if not self.current_pdf:
                self.file_info.config(text="No file loaded")
                return
                
            filename = os.path.basename(self.current_pdf)
            self.file_info.config(text=f"Current file: {filename}")
            
            # Show loading indicator
            loading_label = ttk.Label(self.pdf_frame, text="Loading PDF...", 
                                    font=('Segoe UI', 10))
            loading_label.pack(pady=20)
            self.update()
            
            # Render PDF page
            image = self.pdf_manager.render_pdf_page(self.current_pdf, 
                                                   zoom=self.zoom_level)
            loading_label.destroy()
            
            self.current_image = ImageTk.PhotoImage(image)
            
            # Create frame to hold canvas and scrollbars
            canvas_frame = ttk.Frame(self.pdf_frame)
            canvas_frame.pack(fill='both', expand=True)
            
            # Create scrollable canvas
            canvas = tk.Canvas(
                canvas_frame,
                width=min(self.current_image.width(), 800),  # Limit initial width
                height=min(self.current_image.height(), 800),  # Limit initial height
                bg='#f0f0f0'
            )
            
            # Add scrollbars
            h_scrollbar = ttk.Scrollbar(canvas_frame, orient='horizontal', 
                                      command=canvas.xview)
            v_scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical', 
                                      command=canvas.yview)
            
            # Configure canvas scrolling
            canvas.configure(xscrollcommand=h_scrollbar.set, 
                           yscrollcommand=v_scrollbar.set)
            
            # Pack scrollbars and canvas
            h_scrollbar.pack(side='bottom', fill='x')
            v_scrollbar.pack(side='right', fill='y')
            canvas.pack(side='left', fill='both', expand=True)
            
            def center_image():
                """Center the image in the canvas."""
                # Get canvas size
                canvas_width = canvas.winfo_width()
                canvas_height = canvas.winfo_height()
                
                # Calculate center position
                x = max(0, (canvas_width - self.current_image.width()) // 2)
                y = max(0, (canvas_height - self.current_image.height()) // 2)
                
                # Create image with padding for centering
                padding_width = max(canvas_width, self.current_image.width())
                padding_height = max(canvas_height, self.current_image.height())
                
                # Update canvas scrollregion with padding
                canvas.configure(scrollregion=(
                    -x,  # Left padding
                    -y,  # Top padding
                    padding_width + x,  # Right edge
                    padding_height + y   # Bottom edge
                ))
                
                # Move image to center position
                canvas.delete("all")  # Clear any existing images
                canvas.create_image(x, y, anchor='nw', image=self.current_image)
            
            def on_resize(event: tk.Event) -> None:
                """Handle canvas resize events."""
                if event.widget == canvas:
                    center_image()
            
            def on_mousewheel(event: tk.Event) -> None:
                if event.state & 4:  # Check if Ctrl key is pressed
                    if event.delta > 0:
                        self.zoom_in()
                    else:
                        self.zoom_out()
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            
            def start_drag(event: tk.Event) -> None:
                canvas.scan_mark(event.x, event.y)
                canvas.configure(cursor="fleur")
                
            def do_drag(event: tk.Event) -> None:
                canvas.scan_dragto(event.x, event.y, gain=1)
                
            def stop_drag(event: tk.Event) -> None:
                canvas.configure(cursor="")
                
            def on_key(event: tk.Event) -> None:
                key = event.keysym
                shift_pressed = event.state & 0x1  # Check if Shift is pressed
                
                if key == "Up":
                    canvas.yview_scroll(-1 * (5 if shift_pressed else 1), "units")
                elif key == "Down":
                    canvas.yview_scroll(1 * (5 if shift_pressed else 1), "units")
                elif key == "Left":
                    canvas.xview_scroll(-1 * (5 if shift_pressed else 1), "units")
                elif key == "Right":
                    canvas.xview_scroll(1 * (5 if shift_pressed else 1), "units")
                elif key == "Prior":  # Page Up
                    canvas.yview_scroll(-1, "pages")
                elif key == "Next":  # Page Down
                    canvas.yview_scroll(1, "pages")
                elif key == "Home":
                    canvas.yview_moveto(0)
                elif key == "End":
                    canvas.yview_moveto(1)
            
            # Bind events
            canvas.bind("<MouseWheel>", on_mousewheel)
            canvas.bind("<Button-1>", start_drag)  # Left mouse button
            canvas.bind("<B1-Motion>", do_drag)
            canvas.bind("<ButtonRelease-1>", stop_drag)
            canvas.bind("<Configure>", on_resize)  # Bind resize event
            canvas.bind_all("<Key>", on_key)  # Keyboard navigation
            
            # Initial centering (after a brief delay to ensure canvas is ready)
            self.after(100, center_image)
            
            # Set focus to canvas for keyboard navigation
            canvas.focus_set()
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")
            
    def process_current_file(self) -> None:
        """Queue the current PDF file for processing."""
        if not self.current_pdf or not os.path.exists(self.current_pdf):
            ErrorDialog(self, "Error", "No PDF file loaded")
            return
            
        try:
            value1 = self.filter1_frame.get()
            value2 = self.filter2_frame.get()
            
            if not value1 or not value2:
                return

            # Create and queue the task
            task = PDFTask(
                pdf_path=self.current_pdf,
                value1=value1,
                value2=value2
            )
            
            # Update UI immediately before adding to queue
            self.queue_table.insert('', 'end', 
                                  values=(task.pdf_path, 'Pending'),
                                  tags=('pending',))
            
            # Add to processing queue
            self.pdf_queue.add_task(task)
            
            # Load next PDF and reset filters immediately
            self.load_next_pdf()
            self.filter1_frame.set('')
            self.filter2_frame.set('')
            self.filter2_frame.set_values(self.all_values2)
                
        except Exception as e:
            ErrorDialog(self, "Error", str(e))

    def update_queue_display(self) -> None:
        """Update the queue status display table."""
        try:
            # Get current selection
            selection = self.queue_table.selection()
            selected_paths = [self.queue_table.item(item)['values'][0] for item in selection]
            
            # Get current items in table
            current_items = {}
            for item in self.queue_table.get_children():
                path = self.queue_table.item(item)['values'][0]
                current_items[path] = item
            
            # Update existing items and add new ones
            with self.pdf_queue.lock:
                tasks = self.pdf_queue.tasks.copy()  # Make a copy to minimize lock time
            
            for task_path, task in tasks.items():
                if task_path in current_items:
                    # Update existing item
                    self.queue_table.set(current_items[task_path], 
                                       'status', task.status.capitalize())
                    self.queue_table.item(current_items[task_path], 
                                        tags=(task.status,))
                    current_items.pop(task_path)  # Remove from current items
                else:
                    # Add new item
                    item = self.queue_table.insert('', 'end', 
                                                 values=(task_path, task.status.capitalize()),
                                                 tags=(task.status,))
                    if task_path in selected_paths:
                        self.queue_table.selection_add(item)
            
            # Remove items that no longer exist in the queue
            for item_id in current_items.values():
                self.queue_table.delete(item_id)
                
        except Exception as e:
            # Don't show error dialog for display updates
            print(f"Error updating queue display: {str(e)}")

    def _periodic_update(self) -> None:
        """Update queue display periodically."""
        self.update_queue_display()
        # Use a shorter interval for more responsive updates
        self.after(100, self._periodic_update)

    def __del__(self) -> None:
        """Cleanup resources on deletion."""
        try:
            if hasattr(self, 'pdf_queue'):
                self.pdf_queue.stop()
        except:
            pass
