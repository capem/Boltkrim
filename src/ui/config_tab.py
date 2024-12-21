import tkinter as tk
from tkinter import ttk, filedialog
from .fuzzy_search import FuzzySearchFrame

class ConfigTab(ttk.Frame):
    def __init__(self, master, config_manager, excel_manager):
        super().__init__(master)
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.config_change_callbacks = []
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        self.setup_ui()
        self.setup_styles()
        
    def setup_styles(self):
        """Setup custom styles for widgets"""
        style = ttk.Style()
        
        # Main title style
        style.configure("Title.TLabel", 
                       font=('Helvetica', 14, 'bold'),
                       padding=5)
        
        # Section title style
        style.configure("Section.TLabel",
                       font=('Helvetica', 10, 'bold'),
                       padding=2)
        
        # Custom frame style
        style.configure("Card.TFrame",
                       background='#f0f0f0',
                       relief='solid',
                       borderwidth=1)
        
        # Custom button style
        style.configure("Action.TButton",
                       padding=3,
                       font=('Helvetica', 9))
        
        # LabelFrame style
        style.configure("TLabelframe",
                       padding=5)
        style.configure("TLabelframe.Label",
                       font=('Helvetica', 9, 'bold'))
        
    def setup_ui(self):
        """Create and setup the configuration UI."""
        # Main Title
        title = ttk.Label(self, 
                         text="Configuration Settings",
                         style="Title.TLabel")
        title.grid(row=0, column=0, columnspan=2, pady=(10,5), sticky='n')
        
        # Left Column - Folder Settings
        folder_frame = ttk.LabelFrame(self, text="Folder Settings", padding=5)
        folder_frame.grid(row=1, column=0, padx=5, pady=5, sticky='nsew')
        folder_frame.grid_columnconfigure(1, weight=1)
        
        # Source Folder
        ttk.Label(folder_frame, text="Source:", style="Section.TLabel").grid(row=0, column=0, sticky='w', pady=2)
        self.source_folder_entry = ttk.Entry(folder_frame)
        self.source_folder_entry.grid(row=0, column=1, padx=2, sticky='ew')
        ttk.Button(folder_frame, text="Browse", style="Action.TButton", 
                  command=self.select_source_folder).grid(row=0, column=2)
        
        # Processed Folder
        ttk.Label(folder_frame, text="Processed:", style="Section.TLabel").grid(row=1, column=0, sticky='w', pady=2)
        self.processed_folder_entry = ttk.Entry(folder_frame)
        self.processed_folder_entry.grid(row=1, column=1, padx=2, sticky='ew')
        ttk.Button(folder_frame, text="Browse", style="Action.TButton",
                  command=self.select_processed_folder).grid(row=1, column=2)
        
        # Right Column - Excel Configuration
        excel_frame = ttk.LabelFrame(self, text="Excel Configuration", padding=5)
        excel_frame.grid(row=1, column=1, padx=5, pady=5, sticky='nsew')
        excel_frame.grid_columnconfigure(1, weight=1)
        
        # Excel File
        ttk.Label(excel_frame, text="File:", style="Section.TLabel").grid(row=0, column=0, sticky='w', pady=2)
        self.excel_file_entry = ttk.Entry(excel_frame)
        self.excel_file_entry.grid(row=0, column=1, padx=2, sticky='ew')
        ttk.Button(excel_frame, text="Browse", style="Action.TButton",
                  command=self.select_excel_file).grid(row=0, column=2)
        
        # Excel Sheet
        ttk.Label(excel_frame, text="Sheet:", style="Section.TLabel").grid(row=1, column=0, sticky='w', pady=2)
        self.sheet_combobox = ttk.Combobox(excel_frame, state='readonly')
        self.sheet_combobox.grid(row=1, column=1, columnspan=2, padx=2, sticky='ew')
        self.sheet_combobox.bind('<<ComboboxSelected>>', lambda e: self.update_column_lists())
        
        # Column Configuration - Full Width
        column_frame = ttk.LabelFrame(self, text="Column Configuration", padding=5)
        column_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')
        column_frame.grid_columnconfigure(1, weight=1)
        column_frame.grid_columnconfigure(3, weight=1)
        
        # First Column
        ttk.Label(column_frame, text="First:", style="Section.TLabel").grid(row=0, column=0, sticky='w', pady=2)
        self.filter1_frame = FuzzySearchFrame(column_frame, width=30, identifier='config_filter1')
        self.filter1_frame.grid(row=0, column=1, padx=2, sticky='ew')
        
        # Second Column
        ttk.Label(column_frame, text="Second:", style="Section.TLabel").grid(row=0, column=2, sticky='w', pady=2, padx=(10,0))
        self.filter2_frame = FuzzySearchFrame(column_frame, width=30, identifier='config_filter2')
        self.filter2_frame.grid(row=0, column=3, padx=2, sticky='ew')
        
        # Bottom Frame for Save Button and Status
        bottom_frame = ttk.Frame(self)
        bottom_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Save Button and Status Label side by side
        save_btn = ttk.Button(bottom_frame, 
                             text="Save Configuration",
                             style="Action.TButton",
                             command=self.save_config)
        save_btn.pack(side='left', padx=5)
        
        self.status_label = ttk.Label(bottom_frame, text="", foreground="green")
        self.status_label.pack(side='left', padx=5)
        
        # Load initial values
        self.load_current_config()
        
    def show_status_message(self, message, is_error=False):
        """Show a status message with color coding"""
        self.status_label.config(
            text=message,
            foreground="red" if is_error else "green"
        )
        # Clear the message after 3 seconds
        self.after(3000, lambda: self.status_label.config(text=""))

    def load_current_config(self):
        """Load current configuration into UI elements."""
        config = self.config_manager.get_config()
        self.source_folder_entry.insert(0, config['source_folder'])
        self.processed_folder_entry.insert(0, config['processed_folder'])
        self.excel_file_entry.insert(0, config['excel_file'])
        
        if config['excel_file']:
            self.update_sheet_list()
            # Update column lists after a short delay to ensure widgets are ready
            self.after(100, self.update_column_lists)
            
    def select_source_folder(self):
        """Open dialog to select source folder."""
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder_entry.delete(0, tk.END)
            self.source_folder_entry.insert(0, folder)
            
    def select_processed_folder(self):
        """Open dialog to select processed folder."""
        folder = filedialog.askdirectory()
        if folder:
            self.processed_folder_entry.delete(0, tk.END)
            self.processed_folder_entry.insert(0, folder)
            
    def select_excel_file(self):
        """Open dialog to select Excel file."""
        file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file)
            self.update_sheet_list()
            
    def update_sheet_list(self):
        """Update sheet list in combobox."""
        try:
            excel_file = self.excel_file_entry.get()
            if excel_file:
                sheet_names = self.excel_manager.get_sheet_names(excel_file)
                self.sheet_combobox['values'] = sheet_names
                
                config = self.config_manager.get_config()
                if config['excel_sheet'] in sheet_names:
                    self.sheet_combobox.set(config['excel_sheet'])
                else:
                    self.sheet_combobox.set(sheet_names[0])
                    
                self.update_column_lists()
                
        except Exception as e:
            from .error_dialog import ErrorDialog
            ErrorDialog(self, "Error", f"Error reading Excel file:\n\n{str(e)}")
            
    def update_column_lists(self):
        """Update column lists in fuzzy search frames."""
        try:
            excel_file = self.excel_file_entry.get()
            sheet_name = self.sheet_combobox.get()
            
            if not excel_file or not sheet_name:
                return
                
            # Load Excel data
            self.excel_manager.load_excel_data(excel_file, sheet_name)
            columns = self.excel_manager.get_column_names()
            
            # Update filters with column names
            self.filter1_frame.set_values(columns)
            self.filter2_frame.set_values(columns)
            
            # Set current values if they exist
            config = self.config_manager.get_config()
            if config['filter1_column'] in columns:
                self.filter1_frame.set(config['filter1_column'])
            if config['filter2_column'] in columns:
                self.filter2_frame.set(config['filter2_column'])
                
        except Exception as e:
            from .error_dialog import ErrorDialog
            ErrorDialog(self, "Error", f"Error updating column lists: {str(e)}")
            
    def add_config_change_callback(self, callback):
        """Add a callback to be called when config changes."""
        self.config_change_callbacks.append(callback)

    def save_config(self):
        """Save current configuration."""
        try:
            new_config = {
                'source_folder': self.source_folder_entry.get(),
                'processed_folder': self.processed_folder_entry.get(),
                'excel_file': self.excel_file_entry.get(),
                'excel_sheet': self.sheet_combobox.get(),
                'filter1_column': self.filter1_frame.get(),
                'filter2_column': self.filter2_frame.get(),
                'dropdown1_column': self.filter1_frame.get(),  # Keep backward compatibility
                'dropdown2_column': self.filter2_frame.get()   # Keep backward compatibility
            }
            
            # Basic validation
            if not all([new_config['source_folder'], new_config['processed_folder'], 
                       new_config['excel_file'], new_config['excel_sheet'],
                       new_config['filter1_column'], new_config['filter2_column']]):
                self.show_status_message("Please fill in all required fields", is_error=True)
                return
                
            self.config_manager.update_config(new_config)
            
            # Notify all callbacks about the config change
            for callback in self.config_change_callbacks:
                callback()
                
            self.show_status_message("Configuration saved successfully!")
            
        except Exception as e:
            self.show_status_message(f"Error saving configuration: {str(e)}", is_error=True)
