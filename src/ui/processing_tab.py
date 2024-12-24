from __future__ import annotations
from tkinter import Canvas, END, Event as TkEvent, Widget
from tkinter.ttk import Frame, Scrollbar, Label, Button, Style, LabelFrame, Treeview
from PIL.ImageTk import PhotoImage
from os import path
from typing import Optional, Any, Dict, List, Callable
from threading import Thread, Lock, Event
from dataclasses import dataclass
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog
import pythoncom
from datetime import datetime

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
    def __init__(self, master: Widget, pdf_manager: Any):
        super().__init__(master)
        self.pdf_manager = pdf_manager
        self.current_image: Optional[PhotoImage] = None
        self.current_pdf: Optional[str] = None
        self.zoom_level = 1.0
        self.setup_ui()

    def setup_ui(self) -> None:
        self.canvas_frame = Frame(self)
        self.canvas_frame.pack(fill='both', expand=True)
        
        self.canvas = Canvas(
            self.canvas_frame,
            bg='#f0f0f0'
        )
        
        self.h_scrollbar = Scrollbar(self.canvas_frame, orient='horizontal',
                                   command=self.canvas.xview)
        self.v_scrollbar = Scrollbar(self.canvas_frame, orient='vertical',
                                   command=self.canvas.yview)
        
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set,
                            yscrollcommand=self.v_scrollbar.set)
        
        self.h_scrollbar.pack(side='bottom', fill='x')
        self.v_scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        
        self._bind_events()

    def _bind_events(self) -> None:
        def _on_mousewheel(event):
            if event.state & 4:  # Ctrl key
                if event.delta > 0:
                    self.zoom_in()
                else:
                    self.zoom_out()
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
        def _unbind_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")

        # Bind mousewheel only when mouse is over the PDF viewer area
        self.canvas.bind('<Enter>', _bind_mousewheel)
        self.canvas.bind('<Leave>', _unbind_mousewheel)
        self.v_scrollbar.bind('<Enter>', _bind_mousewheel)
        self.v_scrollbar.bind('<Leave>', _unbind_mousewheel)
        
        self.canvas.bind("<Button-1>", self._start_drag)
        self.canvas.bind("<B1-Motion>", self._do_drag)
        self.canvas.bind("<ButtonRelease-1>", self._stop_drag)
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Key>", self._on_key)

    def _on_mousewheel(self, event: Event) -> None:
        if event.state & 4:  # Ctrl key
            if event.delta > 0:
                self.zoom_in()
            else:
                self.zoom_out()
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _start_drag(self, event: Event) -> None:
        self.canvas.scan_mark(event.x, event.y)
        self.canvas.configure(cursor="fleur")

    def _do_drag(self, event: Event) -> None:
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _stop_drag(self, event: Event) -> None:
        self.canvas.configure(cursor="")

    def _on_key(self, event: Event) -> None:
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
        if event.widget == self.canvas:
            self._center_image()

    def _center_image(self) -> None:
        if not self.current_image:
            return

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        x = max(0, (canvas_width - self.current_image.width()) // 2)
        y = max(0, (canvas_height - self.current_image.height()) // 2)
        
        padding_width = max(canvas_width, self.current_image.width())
        padding_height = max(canvas_height, self.current_image.height())
        
        self.canvas.configure(scrollregion=(
            -x,
            -y,
            padding_width + x,
            padding_height + y
        ))
        
        self.canvas.delete("all")
        self.canvas.create_image(x, y, anchor='nw', image=self.current_image)

    def display_pdf(self, pdf_path: str, zoom: float = 1.0) -> None:
        try:
            self.current_pdf = pdf_path
            self.zoom_level = zoom
            
            for widget in self.canvas_frame.winfo_children():
                if isinstance(widget, Label):
                    widget.destroy()
            
            loading_label = Label(self.canvas_frame, text="Loading PDF...",
                                    font=('Segoe UI', 10))
            loading_label.pack(pady=20)
            self.update()
            
            image = self.pdf_manager.render_pdf_page(pdf_path, zoom=zoom)
            loading_label.destroy()
            
            self.current_image = PhotoImage(image)
            self.after(100, self._center_image)
            self.canvas.focus_set()
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")

    def zoom_in(self, step: float = 0.2) -> None:
        if self.current_pdf:
            self.zoom_level = min(3.0, self.zoom_level + step)
            self.display_pdf(self.current_pdf, self.zoom_level)

    def zoom_out(self, step: float = 0.2) -> None:
        if self.current_pdf:
            self.zoom_level = max(0.2, self.zoom_level - step)
            self.display_pdf(self.current_pdf, self.zoom_level)

class QueueDisplay(Frame):
    def __init__(self, master: Widget):
        super().__init__(master)
        self.setup_ui()

    def setup_ui(self) -> None:
        columns = ('filename', 'status')
        self.table = Treeview(self, columns=columns, show='headings', height=5)
        
        self.table.heading('filename', text='File')
        self.table.heading('status', text='Status')
        
        self.table.column('filename', width=150)
        self.table.column('status', width=80)

        scrollbar = Scrollbar(self, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scrollbar.set)

        self.table.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.table.tag_configure('pending', foreground='black')
        self.table.tag_configure('processing', foreground='blue')
        self.table.tag_configure('completed', foreground='green')
        self.table.tag_configure('failed', foreground='red')

        btn_frame = Frame(self)
        btn_frame.pack(fill='x', pady=(5, 0))
        
        self.clear_btn = Button(btn_frame, text="Clear Completed")
        self.clear_btn.pack(side='left', padx=2)
        
        self.retry_btn = Button(btn_frame, text="Retry Failed")
        self.retry_btn.pack(side='right', padx=2)

    def update_display(self, tasks: Dict[str, PDFTask]) -> None:
        selection = self.table.selection()
        selected_paths = [self.table.item(item)['values'][0] for item in selection]
        
        current_items = {}
        for item in self.table.get_children():
            path = self.table.item(item)['values'][0]
            current_items[path] = item
        
        for task_path, task in tasks.items():
            if task_path in current_items:
                self.table.set(current_items[task_path],
                             'status', task.status.capitalize())
                self.table.item(current_items[task_path],
                              tags=(task.status,))
                current_items.pop(task_path)
            else:
                item = self.table.insert('', 'end',
                                       values=(task_path, task.status.capitalize()),
                                       tags=(task.status,))
                if task_path in selected_paths:
                    self.table.selection_add(item)
        
        for item_id in current_items.values():
            self.table.delete(item_id)

class ProcessingTab(Frame):
    """A tab for processing PDF files with Excel data integration."""
    
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
        
        self.setup_ui()
        self.update_queue_display()
        self.after(100, self._periodic_update)
        
        # Register for config changes
        if hasattr(config_manager, 'add_change_callback'):
            config_manager.add_change_callback(self.handle_config_change)

    def handle_config_change(self) -> None:
        """Handle configuration changes by reloading the current PDF if one is loaded."""
        if self.current_pdf:
            self.pdf_viewer.display_pdf(self.current_pdf)
            self.update_queue_display()

    def setup_ui(self) -> None:
        style = Style()
        style.configure("Action.TButton", padding=5)
        style.configure("Zoom.TButton", padding=2)
        
        container = Frame(self)
        container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Header
        header_frame = Frame(container)
        header_frame.pack(fill='x', pady=(0, 10))
        
        self.file_info = Label(header_frame, text="No file loaded",
                                 font=('Segoe UI', 10))
        self.file_info.pack(side='left')
        
        # Content area
        content_frame = Frame(container)
        content_frame.pack(fill='both', expand=True)
        
        # PDF Viewer
        viewer_frame = LabelFrame(content_frame, text="PDF Viewer", padding=10)
        viewer_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        self.pdf_viewer = PDFViewer(viewer_frame, self.pdf_manager)
        self.pdf_viewer.pack(fill='both', expand=True)
        
        # Controls - Add scrollable frame
        controls_outer_frame = Frame(content_frame)
        controls_outer_frame.pack(side='right', fill='y')
        
        # Create canvas and scrollbar for controls
        controls_canvas = Canvas(controls_outer_frame, width=250)
        controls_scrollbar = Scrollbar(controls_outer_frame, orient="vertical", command=controls_canvas.yview)
        
        # Controls frame
        controls_frame = LabelFrame(controls_canvas, text="Controls", padding=10)
        controls_frame.configure(width=250)
        
        # Configure canvas
        controls_canvas.configure(yscrollcommand=controls_scrollbar.set)
        controls_canvas.pack(side='left', fill='both', expand=True)
        controls_scrollbar.pack(side='right', fill='y')
        
        # Create window in canvas for controls frame
        canvas_frame = controls_canvas.create_window((0, 0), window=controls_frame, anchor='nw', width=230)
        
        self._setup_zoom_controls(controls_frame)
        self._setup_queue_display(controls_frame)
        self._setup_filters(controls_frame)
        self._setup_action_buttons(controls_frame)
        
        # Configure scrolling
        def _configure_canvas(event):
            controls_canvas.configure(scrollregion=controls_canvas.bbox("all"))
            controls_canvas.itemconfig(canvas_frame, width=controls_canvas.winfo_width())
        
        controls_frame.bind('<Configure>', _configure_canvas)
        controls_canvas.bind('<Configure>', lambda e: controls_canvas.itemconfig(canvas_frame, width=controls_canvas.winfo_width()))
        
        # Enable mousewheel scrolling only when mouse is over controls
        def _on_mousewheel(event):
            controls_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def _bind_mousewheel(event):
            controls_canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
        def _unbind_mousewheel(event):
            controls_canvas.unbind_all("<MouseWheel>")
        
        # Bind mousewheel only when mouse enters/leaves the controls area
        controls_canvas.bind('<Enter>', _bind_mousewheel)
        controls_canvas.bind('<Leave>', _unbind_mousewheel)
        controls_scrollbar.bind('<Enter>', _bind_mousewheel)
        controls_scrollbar.bind('<Leave>', _unbind_mousewheel)
        
        self._bind_keyboard_shortcuts()

    def _setup_zoom_controls(self, parent: Frame) -> None:
        zoom_frame = Frame(parent)
        zoom_frame.pack(fill='x', pady=(0, 15))
        
        Label(zoom_frame, text="Zoom:").pack(side='left')
        Button(zoom_frame, text="−", width=3, style="Zoom.TButton",
                  command=self.pdf_viewer.zoom_out).pack(side='left', padx=2)
        self.zoom_label = Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side='left', padx=5)
        Button(zoom_frame, text="+", width=3, style="Zoom.TButton",
                  command=self.pdf_viewer.zoom_in).pack(side='left', padx=2)
        
        rotation_frame = Frame(parent)
        rotation_frame.pack(fill='x', pady=(0, 15))
        
        Label(rotation_frame, text="Rotate:").pack(side='left')
        Button(rotation_frame, text="↶", width=3, style="Zoom.TButton",
                  command=self.rotate_counterclockwise).pack(side='left', padx=2)
        self.rotation_label = Label(rotation_frame, text="0°", width=6)
        self.rotation_label.pack(side='left', padx=5)
        Button(rotation_frame, text="↷", width=3, style="Zoom.TButton",
                  command=self.rotate_clockwise).pack(side='left', padx=2)

    def _setup_queue_display(self, parent: Frame) -> None:
        queue_frame = LabelFrame(parent, text="Processing Queue", padding=10)
        queue_frame.pack(fill='x', pady=(0, 15))
        
        self.queue_display = QueueDisplay(queue_frame)
        self.queue_display.pack(fill='both', expand=True)
        
        self.queue_display.clear_btn.configure(command=self._clear_completed)
        self.queue_display.retry_btn.configure(command=self._retry_failed)
        self.queue_display.table.bind('<Double-1>', self._show_error_details)

    def _setup_filters(self, parent: Frame) -> None:
        filters_frame = LabelFrame(parent, text="Filters", padding=10)
        filters_frame.pack(fill='x', pady=(0, 15))
        
        self.filter1_label = Label(filters_frame, text="",
                                     font=('Segoe UI', 9, 'bold'))
        self.filter1_label.pack(pady=(0, 5))
        self.filter1_frame = FuzzySearchFrame(filters_frame, width=30,
                                            identifier='processing_filter1')
        self.filter1_frame.pack(fill='x', pady=(0, 10))
        
        self.filter2_label = Label(filters_frame, text="",
                                     font=('Segoe UI', 9, 'bold'))
        self.filter2_label.pack(pady=(0, 5))
        self.filter2_frame = FuzzySearchFrame(filters_frame, width=30,
                                            identifier='processing_filter2')
        self.filter2_frame.pack(fill='x', pady=(0, 10))

        self.filter3_label = Label(filters_frame, text="",
                                     font=('Segoe UI', 9, 'bold'))
        self.filter3_label.pack(pady=(0, 5))
        self.filter3_frame = FuzzySearchFrame(filters_frame, width=30,
                                            identifier='processing_filter3')
        self.filter3_frame.pack(fill='x')
        
        self.filter1_frame.bind('<<ValueSelected>>', lambda e: self.on_filter1_select())
        self.filter2_frame.bind('<<ValueSelected>>', lambda e: self.on_filter2_select())
        self.filter3_frame.bind('<<ValueSelected>>', lambda e: self.update_confirm_button())

        # Bind tab navigation
        self.filter1_frame.entry.bind('<Tab>', self._handle_filter1_tab)
        self.filter2_frame.entry.bind('<Tab>', self._handle_filter2_tab)
        self.filter3_frame.entry.bind('<Tab>', self._handle_filter3_tab)
        self.filter1_frame.listbox.bind('<Tab>', self._handle_filter1_tab)
        self.filter2_frame.listbox.bind('<Tab>', self._handle_filter2_tab)
        self.filter3_frame.listbox.bind('<Tab>', self._handle_filter3_tab)

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
        """Handle tab key in filter2 to move focus to next widget."""
        if self.filter2_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item
            if self.filter2_frame.listbox.size() > 0:
                self.filter2_frame.listbox.selection_clear(0, END)
                self.filter2_frame.listbox.selection_set(0)
                self.filter2_frame._on_select(None)
        self.confirm_button.focus_set()
        return "break"

    def _handle_filter3_tab(self, event: Event) -> str:
        """Handle tab key in filter3 to move focus to next widget."""
        if self.filter3_frame.listbox.winfo_ismapped():
            # If listbox is visible, select first item
            if self.filter3_frame.listbox.size() > 0:
                self.filter3_frame.listbox.selection_clear(0, END)
                self.filter3_frame.listbox.selection_set(0)
                self.filter3_frame._on_select(None)
        self.confirm_button.focus_set()
        return "break"

    def _setup_action_buttons(self, parent: Frame) -> None:
        actions_frame = Frame(parent)
        actions_frame.pack(fill='x', pady=(0, 10))
        
        self.confirm_button = Button(
            actions_frame,
            text="Process File (Enter)",
            command=self.process_current_file,
            style="Action.TButton"
        )
        self.confirm_button.pack(fill='x', pady=(0, 5))
        
        self.skip_button = Button(
            actions_frame,
            text="Next File (Ctrl+N)",
            command=self.load_next_pdf,
            style="Action.TButton"
        )
        self.skip_button.pack(fill='x', pady=(0, 10))

    def _bind_keyboard_shortcuts(self) -> None:
        """Bind keyboard shortcuts specific to the processing tab."""
        shortcuts = {
            '<Return>': self.handle_return_key,
            '<Control-n>': lambda e: self.load_next_pdf(),
            '<Control-N>': lambda e: self.load_next_pdf(),
            '<Control-plus>': lambda e: self.pdf_viewer.zoom_in(),
            '<Control-minus>': lambda e: self.pdf_viewer.zoom_out(),
            '<Control-r>': lambda e: self.rotate_clockwise(),
            '<Control-Shift-R>': lambda e: self.rotate_counterclockwise()
        }
        
        def _bind_recursive(widget):
            for key, callback in shortcuts.items():
                widget.bind(key, callback)
            for child in widget.winfo_children():
                _bind_recursive(child)
        
        _bind_recursive(self)
        
        # Also bind to the main frame for good measure
        for key, callback in shortcuts.items():
            self.bind_all(key, callback)

    def handle_return_key(self, event: Event) -> str:
        if str(self.confirm_button['state']) != 'disabled':
            self.process_current_file()
        return "break"

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
            self.all_values1 = sorted(df[config['filter1_column']].unique().tolist())
            self.all_values2 = sorted(df[config['filter2_column']].unique().tolist())
            self.all_values3 = sorted(df[config['filter3_column']].unique().tolist())
            
            self.filter1_frame.set_values(self.all_values1)
            self.filter2_frame.set_values(self.all_values2)
            self.filter3_frame.set_values(self.all_values3)
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")

    def on_filter1_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value = self.filter1_frame.get()
                
                df = self.excel_manager.excel_data
                filtered_df = df[df[config['filter1_column']] == selected_value]
                
                filtered_values2 = sorted(filtered_df[config['filter2_column']].unique().tolist())
                self.filter2_frame.set_values(filtered_values2)
                
                # Clear filter3 since it depends on filter2
                self.filter3_frame.set('')
                self.filter3_frame.set_values(self.all_values3)
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def on_filter2_select(self) -> None:
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                selected_value1 = self.filter1_frame.get()
                selected_value2 = self.filter2_frame.get()
                
                df = self.excel_manager.excel_data
                filtered_df = df[
                    (df[config['filter1_column']] == selected_value1) &
                    (df[config['filter2_column']] == selected_value2)
                ]
                
                filtered_values3 = sorted(filtered_df[config['filter3_column']].unique().tolist())
                self.filter3_frame.set_values(filtered_values3)
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")

    def update_confirm_button(self) -> None:
        if self.filter1_frame.get() and self.filter2_frame.get() and self.filter3_frame.get():
            self.confirm_button.state(['!disabled'])
        else:
            self.confirm_button.state(['disabled'])

    def rotate_clockwise(self) -> None:
        self.pdf_manager.rotate_page(clockwise=True)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(self.current_pdf, self.pdf_viewer.zoom_level)

    def rotate_counterclockwise(self) -> None:
        self.pdf_manager.rotate_page(clockwise=False)
        self.rotation_label.config(text=f"{self.pdf_manager.get_rotation()}°")
        self.pdf_viewer.display_pdf(self.current_pdf, self.pdf_viewer.zoom_level)

    def load_next_pdf(self) -> None:
        try:
            config = self.config_manager.get_config()
            if not config['source_folder']:
                return
                
            next_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
            if next_pdf:
                self.current_pdf = next_pdf
                self.file_info['text'] = path.basename(next_pdf)
                self.pdf_viewer.display_pdf(next_pdf, 1.0)
                self.rotation_label.config(text="0°")
                self.filter1_frame.focus_set()
            else:
                self.file_info['text'] = "No PDF files found"
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")

    def process_current_file(self) -> None:
        if not self.current_pdf or not path.exists(self.current_pdf):
            ErrorDialog(self, "Error", "No PDF file loaded")
            return
            
        try:
            value1 = self.filter1_frame.get()
            value2 = self.filter2_frame.get()
            value3 = self.filter3_frame.get()
            
            if not value1 or not value2 or not value3:
                return

            task = PDFTask(
                pdf_path=self.current_pdf,
                value1=value1,
                value2=value2,
                value3=value3
            )
            
            self.queue_display.table.insert('', 'end',
                                          values=(task.pdf_path, 'Pending'),
                                          tags=('pending',))
            
            self.pdf_queue.add_task(task)
            
            self.load_next_pdf()
            self.filter1_frame.set('')
            self.filter2_frame.set('')
            self.filter3_frame.set('')
            self.filter2_frame.set_values(self.all_values2)
            self.filter3_frame.set_values(self.all_values3)
                
        except Exception as e:
            ErrorDialog(self, "Error", str(e))

    def _show_error_details(self, event: Event) -> None:
        item = self.queue_display.table.selection()[0]
        task_path = self.queue_display.table.item(item)['values'][0]
        
        with self.pdf_queue.lock:
            task = self.pdf_queue.tasks.get(task_path)
            if task and task.status == 'failed' and task.error_msg:
                ErrorDialog(self, "Processing Error",
                          f"Error processing {path.basename(task_path)}:\n{task.error_msg}")

    def _clear_completed(self) -> None:
        self.pdf_queue.clear_completed()
        self.update_queue_display()

    def _retry_failed(self) -> None:
        self.pdf_queue.retry_failed()
        self.update_queue_display()

    def update_queue_display(self) -> None:
        try:
            with self.pdf_queue.lock:
                tasks = self.pdf_queue.tasks.copy()
            self.queue_display.update_display(tasks)
        except Exception as e:
            print(f"Error updating queue display: {str(e)}")

    def _periodic_update(self) -> None:
        self.update_queue_display()
        self.after(100, self._periodic_update)

    def __del__(self) -> None:
        try:
            if hasattr(self, 'pdf_queue'):
                self.pdf_queue.stop()
        except:
            pass
