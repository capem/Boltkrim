from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import os
from typing import Optional, Any
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog

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
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        
        # Initialize state
        self.current_pdf: Optional[str] = None
        self.zoom_level: float = self.INITIAL_ZOOM
        self.current_image: Optional[ImageTk.PhotoImage] = None
        self.all_values1: list[str] = []
        self.all_values2: list[str] = []
        
        self.setup_ui()
        
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
        self._setup_filters(controls_frame)
        self._setup_action_buttons(controls_frame)
        
    def _setup_zoom_controls(self, controls_frame: ttk.Frame) -> None:
        """Setup zoom control buttons and label."""
        zoom_frame = ttk.Frame(controls_frame)
        zoom_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(zoom_frame, text="Zoom:").pack(side='left')
        ttk.Button(zoom_frame, text="−", width=3, style="Zoom.TButton",
                  command=self.zoom_out).pack(side='left', padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side='left', padx=5)
        ttk.Button(zoom_frame, text="+", width=3, style="Zoom.TButton",
                  command=self.zoom_in).pack(side='left', padx=2)
        
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
        
    def load_next_pdf(self) -> None:
        """Load and display the next PDF file from source folder."""
        try:
            config = self.config_manager.get_config()
            if not config['source_folder']:
                ErrorDialog(self, "Error", "Source folder not configured")
                return
            
            self.current_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
            
            if not self.current_pdf:
                self.file_info.config(text="No PDF files found in source folder")
                self.confirm_button.state(['disabled'])
                self.skip_button.state(['disabled'])
                return
                
            self.skip_button.state(['!disabled'])
            self.confirm_button.state(['disabled'])
            self.display_pdf()
            
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
            
            # Create scrollable canvas
            canvas = tk.Canvas(
                self.pdf_frame,
                width=self.current_image.width(),
                height=self.current_image.height(),
                bg='#f0f0f0'
            )
            canvas.pack(fill='both', expand=True)
            
            # Add scrollbars
            h_scrollbar = ttk.Scrollbar(self.pdf_frame, orient='horizontal', 
                                      command=canvas.xview)
            v_scrollbar = ttk.Scrollbar(self.pdf_frame, orient='vertical', 
                                      command=canvas.yview)
            
            canvas.configure(xscrollcommand=h_scrollbar.set, 
                           yscrollcommand=v_scrollbar.set)
            
            def on_mousewheel(event: tk.Event) -> None:
                if event.state & 4:  # Check if Ctrl key is pressed
                    if event.delta > 0:
                        self.zoom_in()
                    else:
                        self.zoom_out()
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            
            canvas.bind("<MouseWheel>", on_mousewheel)
            
            h_scrollbar.pack(side='bottom', fill='x')
            v_scrollbar.pack(side='right', fill='y')
            
            canvas.create_image(0, 0, anchor='nw', image=self.current_image)
            canvas.configure(scrollregion=canvas.bbox('all'))
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")
            
    def process_current_file(self) -> None:
        """Process the current PDF file with selected filter values."""
        if not self.current_pdf or not os.path.exists(self.current_pdf):
            ErrorDialog(self, "Error", "No PDF file loaded")
            return
            
        try:
            value1 = self.filter1_frame.get()
            value2 = self.filter2_frame.get()
            
            if not value1 or not value2:
                return
                
            config = self.config_manager.get_config()
            
            row_data, row_idx = self.excel_manager.find_matching_row(
                config['filter1_column'],
                config['filter2_column'],
                value1,
                value2
            )
            
            if row_data is None:
                ErrorDialog(self, "Error", "Selected combination not found in Excel sheet")
                return
                
            new_filename = f"{row_data[config['filter1_column']]} - {row_data[config['filter2_column']]}"
            
            # Clean filename
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                new_filename = new_filename.replace(char, '_')
                
            if not new_filename.lower().endswith('.pdf'):
                new_filename += '.pdf'
                
            new_filepath = os.path.join(config['processed_folder'], new_filename)
            
            if self.pdf_manager.process_pdf(self.current_pdf, new_filepath, 
                                          config['processed_folder']):
                self.excel_manager.update_pdf_link(
                    config['excel_file'],
                    config['excel_sheet'],
                    row_idx,
                    new_filepath
                )
                
                messagebox.showinfo("Success", f"File processed successfully")
                
                # Load next PDF and reset filters
                self.load_next_pdf()
                self.filter1_frame.set('')
                self.filter2_frame.set('')
                self.filter2_frame.set_values(self.all_values2)
                
        except Exception as e:
            ErrorDialog(self, "Error", str(e))
