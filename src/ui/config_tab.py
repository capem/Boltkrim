from tkinter import LEFT, END, StringVar
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
        self._bind_shortcuts()
        
    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts specific to config tab."""
        self.bind('<Control-s>', lambda e: self.save_config())
        
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
        # Configure grid weights for proper layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # Main Title
        title = ttk.Label(self, 
                         text="Configuration Settings",
                         style="Title.TLabel")
        title.grid(row=0, column=0, columnspan=2, pady=(10,5), sticky='n')
        
        # Create a frame for the main content to manage layout better
        main_content = ttk.Frame(self)
        main_content.grid(row=1, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        main_content.grid_columnconfigure(0, weight=1)
        main_content.grid_columnconfigure(1, weight=1)
        
        # Preset Configuration Frame
        preset_frame = ttk.LabelFrame(main_content, text="Preset Configurations", padding=5)
        preset_frame.grid(row=0, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        preset_frame.grid_columnconfigure(1, weight=1)
        
        # Preset Selector
        ttk.Label(preset_frame, text="Preset:", style="Section.TLabel").grid(row=0, column=0, sticky='w', pady=2)
        self.preset_var = StringVar()
        self.preset_combobox = ttk.Combobox(preset_frame, textvariable=self.preset_var, state='readonly')
        self.preset_combobox.grid(row=0, column=1, padx=2, sticky='ew')
        self.preset_combobox.bind('<<ComboboxSelected>>', self.load_preset)
        
        # Preset Buttons Frame
        preset_buttons_frame = ttk.Frame(preset_frame)
        preset_buttons_frame.grid(row=0, column=2, padx=5)
        
        ttk.Button(preset_buttons_frame, text="Save As Preset", 
                  style="Action.TButton",
                  command=self.save_as_preset).pack(side='left', padx=2)
        ttk.Button(preset_buttons_frame, text="Delete Preset",
                  style="Action.TButton",
                  command=self.delete_preset).pack(side='left', padx=2)
        
        # Left Column - Folder Settings
        folder_frame = ttk.LabelFrame(main_content, text="Folder Settings", padding=5)
        folder_frame.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)
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
        excel_frame = ttk.LabelFrame(main_content, text="Excel Configuration", padding=5)
        excel_frame.grid(row=1, column=1, sticky='nsew', padx=5, pady=5)
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
        column_frame = ttk.LabelFrame(main_content, text="Column Configuration", padding=5)
        column_frame.grid(row=2, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        column_frame.grid_columnconfigure(1, weight=1)
        column_frame.grid_columnconfigure(3, weight=1)
        column_frame.grid_columnconfigure(5, weight=1)
        
        # First Column
        ttk.Label(column_frame, text="First:", style="Section.TLabel").grid(row=0, column=0, sticky='w', pady=2)
        self.filter1_frame = FuzzySearchFrame(column_frame, width=30, identifier='config_filter1')
        self.filter1_frame.grid(row=0, column=1, padx=2, sticky='ew')
        
        # Second Column
        ttk.Label(column_frame, text="Second:", style="Section.TLabel").grid(row=0, column=2, sticky='w', pady=2, padx=(10,0))
        self.filter2_frame = FuzzySearchFrame(column_frame, width=30, identifier='config_filter2')
        self.filter2_frame.grid(row=0, column=3, padx=2, sticky='ew')
        
        # Third Column
        ttk.Label(column_frame, text="Third:", style="Section.TLabel").grid(row=0, column=4, sticky='w', pady=2, padx=(10,0))
        self.filter3_frame = FuzzySearchFrame(column_frame, width=30, identifier='config_filter3')
        self.filter3_frame.grid(row=0, column=5, padx=2, sticky='ew')
        
        # Template Configuration
        template_frame = ttk.LabelFrame(main_content, text="Output Template", padding=5)
        template_frame.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        template_frame.grid_columnconfigure(0, weight=1)
        
        # Template help text
        help_text = (
            "Template Syntax Guide:\n"
            "• Basic fields: {field_name}\n"
            "• Multiple operations: {field|operation1|operation2}\n\n"
            "Date Operations:\n"
            "• Extract year: {DATE FACTURE|date.year}\n"
            "• Extract month: {DATE FACTURE|date.month}\n"
            "• Year-Month: {DATE FACTURE|date.year_month}\n"
            "• Custom format: {DATE FACTURE|date.format:%Y/%m}\n\n"
            "String Operations:\n"
            "• Uppercase: {field|str.upper}\n"
            "• Lowercase: {field|str.lower}\n"
            "• Title Case: {field|str.title}\n"
            "• Replace: {field|str.replace:old:new}\n"
            "• Slice: {field|str.slice:0:4}"
        )
        
        help_label = ttk.Label(template_frame, text=help_text, justify=LEFT)
        help_label.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0,5))
        
        # Template Entry
        ttk.Label(template_frame, text="Template:", style="Section.TLabel").grid(row=1, column=0, sticky='w', pady=2)
        self.template_entry = ttk.Entry(template_frame)
        self.template_entry.grid(row=2, column=0, sticky='ew', pady=(0,5))
        
        # Bottom Frame for Save Button and Status
        bottom_frame = ttk.Frame(main_content)
        bottom_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
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
        self.update_preset_list()
        config = self.config_manager.get_config()
        self.source_folder_entry.insert(0, config['source_folder'])
        self.processed_folder_entry.insert(0, config['processed_folder'])
        self.excel_file_entry.insert(0, config['excel_file'])
        self.template_entry.insert(0, config.get('output_template', '{processed_folder}/{filter1|str.upper} - {filter2|str.upper}.pdf'))
        
        if config['excel_file']:
            self.update_sheet_list()
            # Update column lists after a short delay to ensure widgets are ready
            self.after(100, self.update_column_lists)
            
    def select_source_folder(self):
        """Open dialog to select source folder."""
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder_entry.delete(0, END)
            self.source_folder_entry.insert(0, folder)
            
    def select_processed_folder(self):
        """Open dialog to select processed folder."""
        folder = filedialog.askdirectory()
        if folder:
            self.processed_folder_entry.delete(0, END)
            self.processed_folder_entry.insert(0, folder)
            
    def select_excel_file(self):
        """Open dialog to select Excel file."""
        file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            self.excel_file_entry.delete(0, END)
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
                
            print(f"[DEBUG] Loading Excel data from file: {excel_file}, sheet: {sheet_name}")
            
            # Load Excel data
            self.excel_manager.load_excel_data(excel_file, sheet_name)
            columns = self.excel_manager.get_column_names()
            print(f"[DEBUG] Retrieved columns: {columns}")
            print(f"[DEBUG] Column types: {[type(col) for col in columns]}")
            
            # Update filters with column names
            self.filter1_frame.set_values(columns)
            self.filter2_frame.set_values(columns)
            self.filter3_frame.set_values(columns)
            
            # Set current values if they exist
            config = self.config_manager.get_config()
            print(f"[DEBUG] Current config filter values:")
            print(f"  filter1_column: {config.get('filter1_column')} (type: {type(config.get('filter1_column'))})")
            print(f"  filter2_column: {config.get('filter2_column')} (type: {type(config.get('filter2_column'))})")
            print(f"  filter3_column: {config.get('filter3_column')} (type: {type(config.get('filter3_column'))})")
            
            # Helper function to safely check column existence
            def safe_column_match(column_value, available_columns):
                if not column_value:
                    return False
                # Convert both to strings for comparison
                column_str = str(column_value).strip()
                print(f"[DEBUG] Comparing column value: '{column_str}' (type: {type(column_str)})")
                print(f"[DEBUG] Against available columns: {[str(col).strip() for col in available_columns]}")
                match = any(str(col).strip() == column_str for col in available_columns)
                print(f"[DEBUG] Match found: {match}")
                return match
            
            # Safely set column values
            if safe_column_match(config.get('filter1_column'), columns):
                print(f"[DEBUG] Setting filter1 to: {config['filter1_column']}")
                self.filter1_frame.set(config['filter1_column'])
            if safe_column_match(config.get('filter2_column'), columns):
                print(f"[DEBUG] Setting filter2 to: {config['filter2_column']}")
                self.filter2_frame.set(config['filter2_column'])
            if safe_column_match(config.get('filter3_column'), columns):
                print(f"[DEBUG] Setting filter3 to: {config['filter3_column']}")
                self.filter3_frame.set(config['filter3_column'])
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in update_column_lists:")
            print(traceback.format_exc())
            from .error_dialog import ErrorDialog
            ErrorDialog(self, "Error", f"Error updating column lists: {str(e)}\n\nFull traceback:\n{traceback.format_exc()}")
            
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
                'filter3_column': self.filter3_frame.get(),
                'dropdown1_column': self.filter1_frame.get(),  # Keep backward compatibility
                'dropdown2_column': self.filter2_frame.get(),   # Keep backward compatibility
                'output_template': self.template_entry.get()
            }
            
            # Basic validation
            if not all([new_config['source_folder'], new_config['processed_folder'], 
                       new_config['excel_file'], new_config['excel_sheet'],
                       new_config['filter1_column'], new_config['filter2_column'],
                       new_config['filter3_column'], new_config['output_template']]):
                self.show_status_message("Please fill in all required fields", is_error=True)
                return
                
            self.config_manager.update_config(new_config)
            
            # Notify all callbacks about the config change
            for callback in self.config_change_callbacks:
                callback()
                
            self.show_status_message("Configuration saved successfully!")
            
        except Exception as e:
            self.show_status_message(f"Error saving configuration: {str(e)}", is_error=True)

    def load_preset(self, event=None):
        """Load the selected preset configuration."""
        try:
            preset_name = self.preset_var.get()
            if not preset_name:
                return
                
            print(f"[DEBUG] Loading preset: {preset_name}")
            preset_config = self.config_manager.get_preset(preset_name)
            if not preset_config:
                print(f"[DEBUG] No preset found for name: {preset_name}")
                return
                
            print(f"[DEBUG] Preset config loaded: {preset_config}")
            
            # Clear existing values
            self.source_folder_entry.delete(0, END)
            self.processed_folder_entry.delete(0, END)
            self.excel_file_entry.delete(0, END)
            self.template_entry.delete(0, END)
            
            # Load preset values
            self.source_folder_entry.insert(0, preset_config.get('source_folder', ''))
            self.processed_folder_entry.insert(0, preset_config.get('processed_folder', ''))
            self.excel_file_entry.insert(0, preset_config.get('excel_file', ''))
            self.template_entry.insert(0, preset_config.get('output_template', ''))
            
            # Update Excel-related fields
            if preset_config.get('excel_file'):
                print(f"[DEBUG] Loading Excel file from preset: {preset_config['excel_file']}")
                try:
                    self.update_sheet_list()
                    if preset_config.get('excel_sheet'):
                        print(f"[DEBUG] Setting sheet to: {preset_config['excel_sheet']}")
                        self.sheet_combobox.set(preset_config['excel_sheet'])
                        
                        # Update column lists and wait for them to be populated
                        self.update_column_lists()
                        
                        # Use after() to ensure column lists are updated before setting filter values
                        def set_filter_values():
                            if preset_config.get('filter1_column'):
                                print(f"[DEBUG] Setting filter1 to: {preset_config['filter1_column']}")
                                self.filter1_frame.set(preset_config['filter1_column'])
                            if preset_config.get('filter2_column'):
                                print(f"[DEBUG] Setting filter2 to: {preset_config['filter2_column']}")
                                self.filter2_frame.set(preset_config['filter2_column'])
                            if preset_config.get('filter3_column'):
                                print(f"[DEBUG] Setting filter3 to: {preset_config['filter3_column']}")
                                self.filter3_frame.set(preset_config['filter3_column'])
                                
                        # Wait a short moment for the column lists to be populated
                        self.after(100, set_filter_values)
                        
                except Exception as e:
                    import traceback
                    print(f"[DEBUG] Error loading Excel data:")
                    print(traceback.format_exc())
                    from .error_dialog import ErrorDialog
                    ErrorDialog(self, "Error", f"Error loading Excel data: {str(e)}\n\nFull traceback:\n{traceback.format_exc()}")
                    return
                    
            self.show_status_message(f"Loaded preset: {preset_name}")
            
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error in load_preset:")
            print(traceback.format_exc())
            self.show_status_message(f"Error loading preset: {str(e)}", is_error=True)
        
    def save_as_preset(self):
        """Save current configuration as a new preset."""
        from tkinter import simpledialog
        
        preset_name = simpledialog.askstring("Save Preset", 
                                           "Enter a name for this preset configuration:",
                                           parent=self)
        if not preset_name:
            return
            
        current_config = {
            'source_folder': self.source_folder_entry.get(),
            'processed_folder': self.processed_folder_entry.get(),
            'excel_file': self.excel_file_entry.get(),
            'excel_sheet': self.sheet_combobox.get(),
            'filter1_column': self.filter1_frame.get(),
            'filter2_column': self.filter2_frame.get(),
            'filter3_column': self.filter3_frame.get(),
            'output_template': self.template_entry.get()
        }
        
        self.config_manager.save_preset(preset_name, current_config)
        self.update_preset_list()
        self.preset_var.set(preset_name)
        self.show_status_message(f"Saved preset: {preset_name}")
        
    def delete_preset(self):
        """Delete the currently selected preset."""
        preset_name = self.preset_var.get()
        if not preset_name:
            self.show_status_message("No preset selected", is_error=True)
            return
            
        from tkinter import messagebox
        if messagebox.askyesno("Delete Preset",
                              f"Are you sure you want to delete the preset '{preset_name}'?",
                              parent=self):
            self.config_manager.delete_preset(preset_name)
            self.update_preset_list()
            self.show_status_message(f"Deleted preset: {preset_name}")
            
    def update_preset_list(self):
        """Update the list of available presets in the combobox."""
        presets = self.config_manager.get_preset_names()
        self.preset_combobox['values'] = presets
        if not self.preset_var.get() in presets:
            self.preset_var.set('')
