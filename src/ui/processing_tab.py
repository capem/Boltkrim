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
        # Create main container
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create left frame for PDF viewer with scrollbars
        self.pdf_frame = ttk.Frame(container)
        self.pdf_frame.pack(side='left', fill='both', expand=True)
        
        # Create right frame for controls
        controls_frame = ttk.Frame(container)
        controls_frame.pack(side='right', fill='y', padx=10)
        
        # Add zoom controls
        zoom_frame = ttk.Frame(controls_frame)
        zoom_frame.pack(fill='x', pady=5)
        ttk.Button(zoom_frame, text="-", width=3,
                  command=self.zoom_out).pack(side='left', padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side='left', padx=5)
        ttk.Button(zoom_frame, text="+", width=3,
                  command=self.zoom_in).pack(side='left', padx=2)
        
        # Add search filters with labels
        self.filter1_label = ttk.Label(controls_frame, text="")
        self.filter1_label.pack(pady=5)
        self.filter1_frame = FuzzySearchFrame(controls_frame, width=30, identifier='processing_filter1')
        self.filter1_frame.pack(pady=5)
        
        self.filter2_label = ttk.Label(controls_frame, text="")
        self.filter2_label.pack(pady=5)
        self.filter2_frame = FuzzySearchFrame(controls_frame, width=30, identifier='processing_filter2')
        self.filter2_frame.pack(pady=5)
        
        # Bind filter selection events
        self.filter1_frame.bind('<<ValueSelected>>', lambda e: self.on_filter1_select())
        self.filter2_frame.bind('<<ValueSelected>>', lambda e: None)
        
        # Add confirm button
        ttk.Button(controls_frame, text="Confirm", 
                  command=self.process_current_file).pack(pady=20)
                  
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
        config = self.config_manager.get_config()
        self.current_pdf = self.pdf_manager.get_next_pdf(config['source_folder'])
        self.display_pdf()
        
    def display_pdf(self):
        """Display the current PDF."""
        try:
            # Clear previous image
            for widget in self.pdf_frame.winfo_children():
                widget.destroy()
                
            if not self.current_pdf:
                return
                
            # Render PDF page
            image = self.pdf_manager.render_pdf_page(self.current_pdf, zoom=self.zoom_level)
            
            # Convert to PhotoImage and keep a reference
            self.current_image = ImageTk.PhotoImage(image)
            
            # Create canvas for image
            canvas = tk.Canvas(self.pdf_frame, 
                             width=self.current_image.width(),
                             height=self.current_image.height())
            canvas.pack(fill='both', expand=True)
            
            # Add scrollbars
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
                ErrorDialog(self, "Error", "Please select values from both filters")
                return
                
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
            
            # Preview confirmation
            if not messagebox.askyesno("Confirm",
                f"Current file will be:\n" \
                f"Renamed to: {new_filename}\n" \
                f"Moved to: {config['processed_folder']}\n\n" \
                f"Do you want to proceed?"):
                return
                
            # Check if file already exists
            if os.path.exists(new_filepath):
                if not messagebox.askyesno("Warning", 
                    f"File {new_filename} already exists. Do you want to overwrite it?"):
                    return
                    
            # Process the PDF file
            if self.pdf_manager.process_pdf(self.current_pdf, new_filepath, config['processed_folder']):
                # Update Excel with link
                self.excel_manager.update_pdf_link(
                    config['excel_file'],
                    config['excel_sheet'],
                    row_idx,
                    new_filepath
                )
                
                messagebox.showinfo("Success", f"File processed successfully: {new_filename}")
                
                # Load next PDF
                self.load_next_pdf()
                
        except Exception as e:
            ErrorDialog(self, "Error", str(e))
