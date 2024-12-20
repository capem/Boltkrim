import tkinter as tk
from tkinter import ttk, filedialog
from .fuzzy_search import FuzzySearchFrame

class ConfigTab(ttk.Frame):
    def __init__(self, master, config_manager, excel_manager):
        super().__init__(master)
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.config_change_callbacks = []
        self.setup_ui()
        
    def setup_ui(self):
        """Create and setup the configuration UI."""
        # Title
        ttk.Label(self, text="Configuration", font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Source Folder
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="Source Folder:").pack(side='left')
        self.source_folder_entry = ttk.Entry(frame, width=50)
        self.source_folder_entry.pack(side='left', padx=5)
        ttk.Button(frame, text="Browse", command=self.select_source_folder).pack(side='left')
        
        # Processed Folder
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="Processed Folder:").pack(side='left')
        self.processed_folder_entry = ttk.Entry(frame, width=50)
        self.processed_folder_entry.pack(side='left', padx=5)
        ttk.Button(frame, text="Browse", command=self.select_processed_folder).pack(side='left')
        
        # Excel File
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="Excel File:").pack(side='left')
        self.excel_file_entry = ttk.Entry(frame, width=50)
        self.excel_file_entry.pack(side='left', padx=5)
        ttk.Button(frame, text="Browse", command=self.select_excel_file).pack(side='left')
        
        # Excel Sheet
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="Excel Sheet:").pack(side='left')
        self.sheet_combobox = ttk.Combobox(frame, width=47)
        self.sheet_combobox.pack(side='left', padx=5)
        self.sheet_combobox.bind('<<ComboboxSelected>>', lambda e: self.update_column_lists())
        
        # Column Configurations
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="First Column:").pack(side='left')
        self.filter1_frame = FuzzySearchFrame(frame, width=47, identifier='config_filter1')
        self.filter1_frame.pack(side='left', padx=5, fill='x', expand=True)
        
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(frame, text="Second Column:").pack(side='left')
        self.filter2_frame = FuzzySearchFrame(frame, width=47, identifier='config_filter2')
        self.filter2_frame.pack(side='left', padx=5, fill='x', expand=True)
        
        # Save Button
        ttk.Button(self, text="Save Configuration", 
                  command=self.save_config).pack(pady=20)
                  
        # Load initial values
        self.load_current_config()
        
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
        new_config = {
            'source_folder': self.source_folder_entry.get(),
            'processed_folder': self.processed_folder_entry.get(),
            'excel_file': self.excel_file_entry.get(),
            'excel_sheet': self.sheet_combobox.get(),
            'filter1_column': self.filter1_frame.get(),
            'filter2_column': self.filter2_frame.get()
        }
        self.config_manager.update_config(new_config)
        
        # Notify all callbacks about the config change
        for callback in self.config_change_callbacks:
            callback()
