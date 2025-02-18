from __future__ import annotations
from tkinter import (
    StringVar,
    filedialog,
    Widget,
    Canvas,
    END,
    messagebox,
    LEFT,
    simpledialog,
)
from tkinter.ttk import (
    Frame,
    Label,
    Entry,
    Button,
    Combobox,
    Scrollbar,
    Style,
    LabelFrame,
)
from ..utils import ConfigManager, ExcelManager  # Updated import
from .fuzzy_search import FuzzySearchFrame
from .error_dialog import ErrorDialog


class ConfigTab(Frame):
    """Configuration tab for managing application settings."""

    def __init__(
        self, master: Widget, config_manager: ConfigManager, excel_manager: ExcelManager
    ) -> None:
        """Initialize the configuration tab.

        Args:
            master: Parent widget
            config_manager: Manager for handling configuration
            excel_manager: Manager for Excel operations
        """
        super().__init__(master)
        self.config_manager = config_manager
        self.excel_manager = excel_manager
        self.filter_frames = []  # Store filter frames for dynamic handling

        # Configure base grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_scroll_infrastructure()
        self.setup_styles()
        self.setup_ui()
        self._bind_shortcuts()

    def setup_scroll_infrastructure(self):
        """Setup scrollable canvas infrastructure"""
        # Create canvas and scrollbar
        self.canvas = Canvas(self, bg="SystemButtonFace")
        self.scrollbar = Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)

        # Configure scrollable behavior
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw", width=self.winfo_width()
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Grid layout for canvas and scrollbar
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure scrollable frame grid weights
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        # Bind to resize events
        self.bind("<Configure>", self._on_frame_configure)

        # Bind mouse wheel to scrolling
        self.bind_mouse_wheel(self)

    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts specific to config tab."""
        self.bind("<Control-s>", lambda e: self.save_config())

    def bind_mouse_wheel(self, widget):
        """Bind mouse wheel to all widgets for scrolling"""
        widget.bind("<MouseWheel>", self._on_mouse_wheel)
        widget.bind("<Button-4>", self._on_mouse_wheel)
        widget.bind("<Button-5>", self._on_mouse_wheel)
        for child in widget.winfo_children():
            self.bind_mouse_wheel(child)

    def _on_mouse_wheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.delta < 0:
            self.canvas.yview_scroll(1, "units")
        elif event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        return "break"

    def setup_styles(self):
        """Setup custom styles for widgets"""
        style = Style()

        # Main title style - reduced size
        style.configure("Title.TLabel", font=("Helvetica", 12, "bold"), padding=3)

        # Section title style
        style.configure("Section.TLabel", font=("Helvetica", 10, "bold"), padding=2)

        # Custom frame style
        style.configure(
            "Card.TFrame", background="#f0f0f0", relief="solid", borderwidth=1
        )

        # Custom button style
        style.configure("Action.TButton", padding=3, font=("Helvetica", 9))

        # LabelFrame style
        style.configure("TLabelframe", padding=5)
        style.configure("TLabelframe.Label", font=("Helvetica", 9, "bold"))

    def setup_ui(self):
        """Create and setup the configuration UI."""
        # Configure grid weights for proper layout
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_columnconfigure(1, weight=1)

        # Main Title
        title = Label(
            self.scrollable_frame, text="Configuration Settings", style="Title.TLabel"
        )
        title.grid(row=0, column=0, columnspan=2, pady=(5, 3), sticky="n")

        # Create a frame for the main content to manage layout better
        main_content = Frame(self.scrollable_frame)
        main_content.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        main_content.grid_columnconfigure(0, weight=1)
        main_content.grid_columnconfigure(1, weight=1)

        # Preset Configuration Frame
        preset_frame = LabelFrame(main_content, text="Preset Configurations", padding=5)
        preset_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        preset_frame.grid_columnconfigure(1, weight=1)

        # Preset Selector
        Label(preset_frame, text="Preset:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.preset_var = StringVar()
        self.preset_combobox = Combobox(
            preset_frame, textvariable=self.preset_var, state="readonly"
        )
        self.preset_combobox.grid(row=0, column=1, padx=2, sticky="ew")
        self.preset_combobox.bind("<<ComboboxSelected>>", self.load_preset)

        # Preset Buttons Frame
        preset_buttons_frame = Frame(preset_frame)
        preset_buttons_frame.grid(row=0, column=2, padx=5)

        Button(
            preset_buttons_frame,
            text="Save As Preset",
            style="Action.TButton",
            command=self.save_as_preset,
        ).pack(side="left", padx=2)
        Button(
            preset_buttons_frame,
            text="Delete Preset",
            style="Action.TButton",
            command=self.delete_preset,
        ).pack(side="left", padx=2)

        # Left Column - Folder Settings
        folder_frame = LabelFrame(main_content, text="Folder Settings", padding=5)
        folder_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        folder_frame.grid_columnconfigure(1, weight=1)

        # Source Folder
        Label(folder_frame, text="Source:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.source_folder_entry = Entry(folder_frame)
        self.source_folder_entry.grid(row=0, column=1, padx=2, sticky="ew")
        Button(
            folder_frame,
            text="Browse",
            style="Action.TButton",
            command=lambda: self.select_folder(self.source_folder_entry),
        ).grid(row=0, column=2)

        # Processed Folder
        Label(folder_frame, text="Processed:", style="Section.TLabel").grid(
            row=1, column=0, sticky="w", pady=2
        )
        self.processed_folder_entry = Entry(folder_frame)
        self.processed_folder_entry.grid(row=1, column=1, padx=2, sticky="ew")
        Button(
            folder_frame,
            text="Browse",
            style="Action.TButton",
            command=lambda: self.select_folder(self.processed_folder_entry),
        ).grid(row=1, column=2)

        # Right Column - Excel Configuration
        excel_frame = LabelFrame(main_content, text="Excel Configuration", padding=5)
        excel_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        excel_frame.grid_columnconfigure(1, weight=1)

        # Excel File
        Label(excel_frame, text="File:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.excel_file_entry = Entry(excel_frame)
        self.excel_file_entry.grid(row=0, column=1, padx=2, sticky="ew")
        Button(
            excel_frame,
            text="Browse",
            style="Action.TButton",
            command=self.select_excel_file,
        ).grid(row=0, column=2)

        # Excel Sheet
        Label(excel_frame, text="Sheet:", style="Section.TLabel").grid(
            row=1, column=0, sticky="w", pady=2
        )
        self.sheet_combobox = Combobox(excel_frame, state="readonly")
        self.sheet_combobox.grid(row=1, column=1, columnspan=2, padx=2, sticky="ew")
        self.sheet_combobox.bind(
            "<<ComboboxSelected>>", lambda e: self.update_column_lists()
        )

        # Column Configuration Frame
        column_frame = LabelFrame(main_content, text="Column Configuration", padding=5)
        column_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        column_frame.grid_columnconfigure(0, weight=1)

        # Create container for filters
        self.filters_container = Frame(column_frame)
        self.filters_container.grid(row=0, column=0, sticky="nsew")
        self.filters_container.grid_columnconfigure(0, weight=1)

        # Add button for new filters
        add_filter_btn = Button(
            column_frame,
            text="Add Filter",
            style="Action.TButton",
            command=lambda: self._add_filter("", None)
        )
        add_filter_btn.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        # Template Configuration
        template_frame = LabelFrame(main_content, text="Output Template", padding=5)
        template_frame.grid(
            row=3, column=0, columnspan=2, sticky="nsew", padx=5, pady=5
        )
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

        help_label = Label(template_frame, text=help_text, justify=LEFT)
        help_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        # Template Entry
        Label(template_frame, text="Template:", style="Section.TLabel").grid(
            row=1, column=0, sticky="w", pady=2
        )
        self.template_entry = Entry(template_frame)
        self.template_entry.grid(row=2, column=0, sticky="ew", pady=(0, 5))

        # Bottom Frame for Save Button and Status
        bottom_frame = Frame(main_content)
        bottom_frame.grid(row=4, column=0, columnspan=2, pady=5)

        # Save Button and Status Label side by side
        save_btn = Button(
            bottom_frame,
            text="Save Configuration",
            style="Action.TButton",
            command=self.save_config,
        )
        save_btn.pack(side="left", padx=5)

        self.status_label = Label(bottom_frame, text="", foreground="green")
        self.status_label.pack(side="left", padx=5)

        # Load initial values
        self.load_current_config()

    def show_status_message(self, message, is_error=False):
        """Show a status message with color coding"""
        self.status_label.config(
            text=message, foreground="red" if is_error else "green"
        )
        # Clear the message after 3 seconds
        self.after(3000, lambda: self.status_label.config(text=""))

    def load_current_config(self):
        """Load current configuration into UI elements."""
        self.update_preset_list()
        config = self.config_manager.get_config()
        
        # Load basic settings
        self.source_folder_entry.insert(0, config["source_folder"])
        self.processed_folder_entry.insert(0, config["processed_folder"])
        self.excel_file_entry.insert(0, config["excel_file"])
        self.template_entry.insert(0, config["output_template"])

        # Clear existing filters
        for frame in self.filter_frames:
            frame['frame'].destroy()
        self.filter_frames.clear()

        # Load filters from config
        filter_columns = self._get_filter_columns_from_config(config)
        
        # Always ensure at least 3 filters
        while len(filter_columns) < 3:
            filter_columns.append("")

        # Create filter frames
        for i, column in enumerate(filter_columns, 1):
            self._add_filter(column, f"filter{i}")

        if config["excel_file"]:
            self.update_sheet_list()
            # Update column lists after a short delay to ensure widgets are ready
            self.after(100, self.update_column_lists)

    def select_folder(self, entry_widget):
        """Open dialog to select a folder and update the specified entry."""
        folder = filedialog.askdirectory()
        if folder:
            entry_widget.delete(0, END)
            entry_widget.insert(0, folder)

    def select_excel_file(self):
        """Open dialog to select Excel file."""
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            self.excel_file_entry.delete(0, END)
            self.excel_file_entry.insert(0, file)
            self.update_sheet_list()

    def update_sheet_list(self, config=None):
        """Update sheet list in combobox.

        Args:
            config (dict, optional): Configuration to use. If None, uses default config.
        """
        try:
            excel_file = self.excel_file_entry.get()
            if excel_file:
                sheet_names = self.excel_manager.get_sheet_names(excel_file)
                self.sheet_combobox["values"] = sheet_names
                if config is None:
                    config = self.config_manager.get_config()
                self.sheet_combobox.set(config["excel_sheet"])
                self.update_column_lists()

        except Exception as e:
            ErrorDialog(self, "Error Loading Excel", e)

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

            # Convert columns to strings for comparison
            str_columns = [str(col) for col in columns]

            config = self.config_manager.get_config()

            # Update each filter frame with available columns
            for frame in self.filter_frames:
                frame['fuzzy_frame'].set_values(columns)

            # Set selected values from config
            for i, frame in enumerate(self.filter_frames, 1):
                column_key = f"filter{i}_column"
                if str(config.get(column_key)) in str_columns:
                    frame['fuzzy_frame'].set(config[column_key])
                    frame['label']['text'] = config[column_key]

        except Exception as e:
            ErrorDialog(self, "Error Updating Columns", e)

    def save_config(self):
        """Save current configuration with dynamic filters."""
        try:
            # Create new config with only current values
            new_config = {
                "source_folder": self.source_folder_entry.get(),
                "processed_folder": self.processed_folder_entry.get(),
                "excel_file": self.excel_file_entry.get(),
                "excel_sheet": self.sheet_combobox.get(),
                "output_template": self.template_entry.get(),
            }

            # Add current filter columns to config
            for i, filter_dict in enumerate(self.filter_frames, 1):
                filter_value = filter_dict['fuzzy_frame'].get()
                new_config[f"filter{i}_column"] = filter_value

            # Basic validation - ensure at least 3 filters are configured
            required_fields = [
                "source_folder",
                "processed_folder",
                "excel_file",
                "excel_sheet",
                "output_template",
            ]
            required_fields.extend([f"filter{i}_column" for i in range(1, 4)])  # First 3 filters are required

            if not all(new_config.get(field) for field in required_fields):
                ErrorDialog(self, "Validation Error", "Please fill in all required fields (including at least 3 filters)")
                return

            # Reset config to defaults first, then update with new values
            self.config_manager.reset_config()
            self.config_manager.update_config(new_config)
            self.show_status_message("Configuration saved successfully!")

        except Exception as e:
            ErrorDialog(self, "Save Error", e)

    def load_preset(self, event=None):
        """Load the selected preset configuration."""
        try:
            preset_name = self.preset_var.get()
            if not preset_name:
                return

            print(f"[DEBUG] Loading preset: {preset_name}")
            preset_config = self.config_manager.get_preset(preset_name)
            if not preset_config:
                ErrorDialog(
                    self, "Preset Error", f"No preset found with name: {preset_name}"
                )
                return

            print(f"[DEBUG] Preset config loaded: {preset_config}")

            # Clear existing values
            self.source_folder_entry.delete(0, END)
            self.processed_folder_entry.delete(0, END)
            self.excel_file_entry.delete(0, END)
            self.template_entry.delete(0, END)

            # Load basic settings
            self.source_folder_entry.insert(0, preset_config.get("source_folder", ""))
            self.processed_folder_entry.insert(0, preset_config.get("processed_folder", ""))
            self.excel_file_entry.insert(0, preset_config.get("excel_file", ""))
            self.template_entry.insert(0, preset_config.get("output_template", ""))

            # Update sheet list first
            self.update_sheet_list(preset_config)

            # Clear existing filters
            for frame in self.filter_frames:
                frame['frame'].destroy()
            self.filter_frames.clear()

            # Load filters from preset
            filter_columns = self._get_filter_columns_from_config(preset_config)
            
            # Always ensure at least 3 filters
            while len(filter_columns) < 3:
                filter_columns.append("")

            # Create filter frames
            for i, column in enumerate(filter_columns, 1):
                self._add_filter(column, f"filter{i}")

            # Update column lists after a short delay to ensure widgets are ready
            self.after(100, lambda: self._update_columns_from_preset(preset_config))

            self.show_status_message(f"Loaded preset: {preset_name}")

        except Exception as e:
            ErrorDialog(self, "Load Preset Error", e)

    def _update_columns_from_preset(self, preset_config):
        """Update column lists with preset values."""
        try:
            excel_file = self.excel_file_entry.get()
            sheet_name = self.sheet_combobox.get()

            if not excel_file or not sheet_name:
                return

            # Load Excel data
            self.excel_manager.load_excel_data(excel_file, sheet_name)
            columns = self.excel_manager.get_column_names()

            # Convert columns to strings for comparison
            str_columns = [str(col) for col in columns]

            # Update each filter frame with available columns
            for frame in self.filter_frames:
                frame['fuzzy_frame'].set_values(columns)

            # Set selected values from preset
            for i, frame in enumerate(self.filter_frames, 1):
                column_key = f"filter{i}_column"
                if column_key in preset_config and str(preset_config[column_key]) in str_columns:
                    frame['fuzzy_frame'].set(preset_config[column_key])
                    frame['label']['text'] = preset_config[column_key]

            # Save configuration after updating
            self.after(100, self.save_config)

        except Exception as e:
            ErrorDialog(self, "Error Updating Columns", e)

    def save_as_preset(self):
        """Save current configuration as a new preset."""
        try:
            preset_name = simpledialog.askstring(
                "Save Preset",
                "Enter a name for this preset configuration:",
                parent=self,
            )
            if not preset_name:
                return

            # Get current configuration
            current_config = {
                "source_folder": self.source_folder_entry.get(),
                "processed_folder": self.processed_folder_entry.get(),
                "excel_file": self.excel_file_entry.get(),
                "excel_sheet": self.sheet_combobox.get(),
                "output_template": self.template_entry.get(),
            }

            # Add filter columns to config
            for i, filter_dict in enumerate(self.filter_frames, 1):
                filter_value = filter_dict['fuzzy_frame'].get()
                current_config[f"filter{i}_column"] = filter_value

            self.config_manager.save_preset(preset_name, current_config)
            self.update_preset_list()
            self.preset_var.set(preset_name)
            self.show_status_message(f"Saved preset: {preset_name}")

        except Exception as e:
            ErrorDialog(self, "Save Preset Error", e)

    def delete_preset(self):
        """Delete the currently selected preset."""
        try:
            preset_name = self.preset_var.get()
            if not preset_name:
                ErrorDialog(self, "Delete Error", "No preset selected")
                return

            if messagebox.askyesno(
                "Delete Preset",
                f"Are you sure you want to delete the preset '{preset_name}'?",
                parent=self,
            ):
                self.config_manager.delete_preset(preset_name)
                self.update_preset_list()
                self.show_status_message(f"Deleted preset: {preset_name}")

        except Exception as e:
            ErrorDialog(self, "Delete Preset Error", e)

    def update_preset_list(self):
        """Update the list of available presets in the combobox."""
        presets = self.config_manager.get_preset_names()
        self.preset_combobox["values"] = presets
        if self.preset_var.get() not in presets:
            self.preset_var.set("")

    def _on_frame_configure(self, event=None):
        """Handle frame resize"""
        # Update the canvas window width when the frame is resized
        self.canvas.itemconfig(
            "all", width=self.winfo_width() - self.scrollbar.winfo_width()
        )

    def _setup_filters(self, parent: Widget) -> None:
        """Setup dynamic filter controls with improved styling."""
        # Create a frame to hold all filter frames
        self.filters_container = Frame(parent)
        self.filters_container.pack(fill="x", expand=True)

        # Add button to add new filter
        add_filter_btn = Button(
            parent,
            text="Add Filter",
            style="Action.TButton",
            command=self._add_filter
        )
        add_filter_btn.pack(fill="x", pady=(10, 0))

        # Load initial filters from config
        config = self.config_manager.get_config()
        filter_columns = self._get_filter_columns_from_config(config)
        
        # Always ensure at least 3 filters
        while len(filter_columns) < 3:
            filter_columns.append("")

        # Create initial filters
        for i, column in enumerate(filter_columns, 1):
            self._add_filter(column, f"filter{i}")

    def _get_filter_columns_from_config(self, config: dict) -> list:
        """Extract filter columns from config."""
        filter_columns = []
        i = 1
        while True:
            filter_key = f"filter{i}_column"
            if filter_key not in config:
                break
            filter_columns.append(config[filter_key])
            i += 1
        return filter_columns

    def _add_filter(self, column_name: str = "", identifier: str = None) -> None:
        """Add a new filter frame to the configuration."""
        filter_num = len(self.filter_frames) + 1
        if identifier is None:
            identifier = f"filter{filter_num}"

        # Create frame for this filter
        filter_frame = Frame(self.filters_container)
        filter_frame.pack(fill="x", pady=(0, 15))

        # Label for the filter
        label = Label(filter_frame, text=column_name or f"Filter {filter_num}:", style="Header.TLabel")
        label.pack(pady=(0, 5))

        # Create fuzzy search frame
        fuzzy_frame = FuzzySearchFrame(
            filter_frame,
            width=30,
            identifier=f"config_{identifier}"
        )
        fuzzy_frame.pack(fill="x")

        # Add remove button if not one of the first three filters
        if filter_num > 3:
            remove_btn = Button(
                filter_frame,
                text="Remove",
                style="Action.TButton",
                command=lambda f=filter_frame: self._remove_filter(f)
            )
            remove_btn.pack(pady=(5, 0))

        # Store references
        self.filter_frames.append({
            'frame': filter_frame,
            'label': label,
            'fuzzy_frame': fuzzy_frame,
            'identifier': identifier
        })

        # Initialize with available columns if Excel file is loaded
        try:
            excel_file = self.excel_file_entry.get()
            sheet_name = self.sheet_combobox.get()
            if excel_file and sheet_name:
                # Load Excel data if not already loaded
                if self.excel_manager.excel_data is None:
                    self.excel_manager.load_excel_data(excel_file, sheet_name)
                
                # Get column names and set them as available values
                columns = self.excel_manager.get_column_names()
                fuzzy_frame.set_values(columns)

                # If column_name is provided and exists in columns, select it
                if column_name and column_name in columns:
                    fuzzy_frame.set(column_name)
                    label['text'] = column_name
        except Exception as e:
            print(f"[DEBUG] Error initializing filter values: {str(e)}")

        # Update the configuration after adding the filter
        self.after(100, self.save_config)

    def _remove_filter(self, filter_frame: Frame) -> None:
        """Remove a filter frame from the configuration."""
        # Find and remove the filter from our stored references
        for i, filter_dict in enumerate(self.filter_frames):
            if filter_dict['frame'] == filter_frame:
                self.filter_frames.pop(i)
                break
        
        # Destroy the frame
        filter_frame.destroy()

        # Update the remaining filter labels
        self._update_filter_labels()

    def _update_filter_labels(self) -> None:
        """Update the labels of all filters to maintain sequential numbering."""
        for i, filter_dict in enumerate(self.filter_frames, 1):
            if not filter_dict['label']['text']:  # Only update if no specific column name
                filter_dict['label']['text'] = f"Filter {i}:"
