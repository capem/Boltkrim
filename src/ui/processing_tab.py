import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageTk
import os
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog

class ProcessingTab(ttk.Frame):
    def __init__(self, master, config_manager, excel_manager, pdf_manager):
        super().__init__(master)
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.pdf_manager = pdf_manager
        
        # Initialize state
        self.current_pdf = None
        self.zoom_level = 1.0
        self.current_image = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create and setup the processing UI."""
        style = ttk.Style()
        style.configure("Action.TButton", padding=5)
        style.configure("Zoom.TButton", padding=2)
        
        # Create main container with padding
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Create header frame
        header_frame = ttk.Frame(container)
        header_frame.pack(fill='x', pady=(0, 10))
        
        # Add file info label
        self.file_info = ttk.Label(header_frame, text="No file loaded", font=('Segoe UI', 10))
        self.file_info.pack(side='left')
        
        # Create main content frame with 3:1 ratio
        content_frame = ttk.Frame(container)
        content_frame.pack(fill='both', expand=True)
        
        # Create left frame for PDF viewer with scrollbars
        viewer_frame = ttk.LabelFrame(content_frame, text="PDF Viewer", padding=10)
        viewer_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        self.pdf_frame = ttk.Frame(viewer_frame)
        self.pdf_frame.pack(fill='both', expand=True)
        
        # Create right frame for controls
        controls_frame = ttk.LabelFrame(content_frame, text="Controls", padding=10)
        controls_frame.configure(width=250)  # Set width using configure
        controls_frame.pack(side='right', fill='y')
        controls_frame.pack_propagate(False)  # Prevent the frame from shrinking
        
        # Add zoom controls with modern styling
        zoom_frame = ttk.Frame(controls_frame)
        zoom_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(zoom_frame, text="Zoom:").pack(side='left')
        ttk.Button(zoom_frame, text="−", width=3, style="Zoom.TButton",
                  command=self.zoom_out).pack(side='left', padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%", width=6)
        self.zoom_label.pack(side='left', padx=5)
        ttk.Button(zoom_frame, text="+", width=3, style="Zoom.TButton",
                  command=self.zoom_in).pack(side='left', padx=2)
        
        # Add search filters section
        filters_frame = ttk.LabelFrame(controls_frame, text="Filters", padding=10)
        filters_frame.pack(fill='x', pady=(0, 15))
        
        # First filter
        self.filter1_label = ttk.Label(filters_frame, text="", font=('Segoe UI', 9, 'bold'))
        self.filter1_label.pack(pady=(0, 5))
        self.filter1_frame = FuzzySearchFrame(filters_frame, width=30, identifier='processing_filter1')
        self.filter1_frame.pack(fill='x', pady=(0, 10))
        
        # Second filter
        self.filter2_label = ttk.Label(filters_frame, text="", font=('Segoe UI', 9, 'bold'))
        self.filter2_label.pack(pady=(0, 5))
        self.filter2_frame = FuzzySearchFrame(filters_frame, width=30, identifier='processing_filter2')
        self.filter2_frame.pack(fill='x')
        
        # Add action buttons
        actions_frame = ttk.Frame(controls_frame)
        actions_frame.pack(fill='x', pady=(0, 10))
        
        self.confirm_button = ttk.Button(actions_frame, text="Process File (Enter)", 
                                       command=self.process_current_file, style="Action.TButton")
        self.confirm_button.pack(fill='x', pady=(0, 5))
        
        self.skip_button = ttk.Button(actions_frame, text="Skip File (→)", 
                                    command=self.load_next_pdf, style="Action.TButton")
        self.skip_button.pack(fill='x')
        
        # Add keyboard shortcuts
        self.bind_all('<Return>', lambda e: self.handle_return_key(e))
        self.bind_all('<Right>', lambda e: self.load_next_pdf())
        self.bind_all('<Control-plus>', lambda e: self.zoom_in())
        self.bind_all('<Control-minus>', lambda e: self.zoom_out())
        
        # Bind filter selection events
        self.filter1_frame.bind('<<ValueSelected>>', lambda e: self.on_filter1_select())
        self.filter2_frame.bind('<<ValueSelected>>', lambda e: self.update_confirm_button())
        
    def handle_return_key(self, event):
        """Handle Return key press"""
        if str(self.confirm_button['state']) != 'disabled':
            self.process_current_file()
        return "break"  # Prevent event propagation

    def load_excel_data(self):
        """Load data from Excel file."""
        try:
            config = self.config_manager.get_config()
            if not all([config['excel_file'], config['excel_sheet'],
                       config['filter1_column'], config['filter2_column']]):
                print("Missing configuration values")
                return
                
            # Load Excel data
            self.excel_manager.load_excel_data(config['excel_file'], config['excel_sheet'])
            
            # Update filter labels
            self.filter1_label['text'] = config['filter1_column']
            self.filter2_label['text'] = config['filter2_column']
            
            # Get unique values for filters
            df = self.excel_manager.excel_data
            self.all_values1 = sorted(df[config['filter1_column']].unique().tolist())
            self.all_values2 = sorted(df[config['filter2_column']].unique().tolist())
            
            # Update filters with values
            self.filter1_frame.set_values(self.all_values1)
            self.filter2_frame.set_values(self.all_values2)
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}")
            
    def on_filter1_select(self):
        """Handle selection in first filter."""
        try:
            config = self.config_manager.get_config()
            if self.excel_manager.excel_data is not None:
                # Get selected value from first filter
                selected_value = self.filter1_frame.get()
                
                # Filter second filter based on first selection
                df = self.excel_manager.excel_data
                filtered_df = df[df[config['filter1_column']] == selected_value]
                
                # Update second filter values
                filtered_values = sorted(filtered_df[config['filter2_column']].unique().tolist())
                self.filter2_frame.set_values(filtered_values)
                
        except Exception as e:
            ErrorDialog(self, "Error", f"Error updating filters: {str(e)}")
            
    def update_confirm_button(self):
        """Update confirm button state based on filter selections"""
        if self.filter1_frame.get() and self.filter2_frame.get():
            self.confirm_button.state(['!disabled'])
        else:
            self.confirm_button.state(['disabled'])
            
    def zoom_in(self):
        """Increase zoom level."""
        self.zoom_level = min(3.0, self.zoom_level + 0.2)
        self.display_pdf()
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        
    def zoom_out(self):
        """Decrease zoom level."""
        self.zoom_level = max(0.2, self.zoom_level - 0.2)
        self.display_pdf()
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        
    def load_next_pdf(self):
        """Load the next PDF file from source folder."""
        try:
            config = self.config_manager.get_config()
            if not config['source_folder']:
                ErrorDialog(self, "Error", "Source folder not configured")
                return
            
            # Get next PDF
            self.current_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
            
            if not self.current_pdf:
                self.file_info.config(text="No PDF files found in source folder")
                # Disable buttons when no files are available
                self.confirm_button.state(['disabled'])
                self.skip_button.state(['disabled'])
                return
                
            # Enable skip button when files are available
            self.skip_button.state(['!disabled'])
            # Confirm button stays disabled until filters are selected
            self.confirm_button.state(['disabled'])
            
            # Display the PDF
            self.display_pdf()
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error loading next PDF: {str(e)}")
            
    def display_pdf(self):
        """Display the current PDF."""
        try:
            # Clear previous image
            for widget in self.pdf_frame.winfo_children():
                widget.destroy()
                
            if not self.current_pdf:
                self.file_info.config(text="No file loaded")
                return
                
            # Update file info
            filename = os.path.basename(self.current_pdf)
            self.file_info.config(text=f"Current file: {filename}")
                
            # Show loading indicator
            loading_label = ttk.Label(self.pdf_frame, text="Loading PDF...", font=('Segoe UI', 10))
            loading_label.pack(pady=20)
            self.update()
            
            # Render PDF page
            image = self.pdf_manager.render_pdf_page(self.current_pdf, zoom=self.zoom_level)
            
            # Remove loading indicator
            loading_label.destroy()
            
            # Convert to PhotoImage and keep a reference
            self.current_image = ImageTk.PhotoImage(image)
            
            # Create canvas for image with light gray background
            canvas = tk.Canvas(self.pdf_frame, 
                             width=self.current_image.width(),
                             height=self.current_image.height(),
                             bg='#f0f0f0')
            canvas.pack(fill='both', expand=True)
            
            # Add scrollbars with modern style
            h_scrollbar = ttk.Scrollbar(self.pdf_frame, orient='horizontal', command=canvas.xview)
            v_scrollbar = ttk.Scrollbar(self.pdf_frame, orient='vertical', command=canvas.yview)
            
            # Configure canvas scrolling
            canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
            
            # Pack scrollbars
            h_scrollbar.pack(side='bottom', fill='x')
            v_scrollbar.pack(side='right', fill='y')
            
            # Display image
            canvas.create_image(0, 0, anchor='nw', image=self.current_image)
            canvas.configure(scrollregion=canvas.bbox('all'))
            
        except Exception as e:
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")
            
    def process_current_file(self):
        """Process the current PDF file."""
        if not self.current_pdf or not os.path.exists(self.current_pdf):
            ErrorDialog(self, "Error", "No PDF file loaded")
            return
            
        try:
            # Get selected values
            value1 = self.filter1_frame.get()
            value2 = self.filter2_frame.get()
            
            if not value1 or not value2:
                return  # Silently return if filters are not selected
                
            config = self.config_manager.get_config()
            
            # Find the matching row in Excel
            row_data, row_idx = self.excel_manager.find_matching_row(
                config['filter1_column'],
                config['filter2_column'],
                value1,
                value2
            )
            
            if row_data is None:
                ErrorDialog(self, "Error", "Selected combination not found in Excel sheet")
                return
                
            # Generate new filename
            new_filename = f"{row_data[config['filter1_column']]} - {row_data[config['filter2_column']]}"
            
            # Clean the filename of invalid characters
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                new_filename = new_filename.replace(char, '_')
                
            # Add .pdf extension if needed
            if not new_filename.lower().endswith('.pdf'):
                new_filename += '.pdf'
                
            # Generate full path for new file
            new_filepath = os.path.join(config['processed_folder'], new_filename)
            
            try:
                # Process the PDF file
                if self.pdf_manager.process_pdf(self.current_pdf, new_filepath, config['processed_folder']):
                    # Update Excel with link
                    self.excel_manager.update_pdf_link(
                        config['excel_file'],
                        config['excel_sheet'],
                        row_idx,
                        new_filepath
                    )
                    
                    messagebox.showinfo("Success", f"File processed successfully")
                    
                    # Load next PDF first
                    self.load_next_pdf()
                    
                    # Reset filter selections and update second filter values
                    self.filter1_frame.set('')
                    self.filter2_frame.set('')
                    self.filter2_frame.set_values(self.all_values2)  # Reset second filter to show all values
                    
            except Exception as e:
                ErrorDialog(self, "Error", str(e))
            
        except Exception as e:
            ErrorDialog(self, "Error", str(e))
