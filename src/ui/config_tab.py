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
from typing import Optional, Callable
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

        # Column Configuration - Full Width
        column_frame = LabelFrame(main_content, text="Column Configuration", padding=5)
        column_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        column_frame.grid_columnconfigure(1, weight=1)
        column_frame.grid_columnconfigure(3, weight=1)
        column_frame.grid_columnconfigure(5, weight=1)

        # First Column
        Label(column_frame, text="First:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.filter1_frame = FuzzySearchFrame(
            column_frame, width=30, identifier="config_filter1"
        )
        self.filter1_frame.grid(row=0, column=1, padx=2, sticky="ew")

        # Second Column
        Label(column_frame, text="Second:", style="Section.TLabel").grid(
            row=0, column=2, sticky="w", pady=2, padx=2
        )
        self.filter2_frame = FuzzySearchFrame(
            column_frame, width=30, identifier="config_filter2"
        )
        self.filter2_frame.grid(row=0, column=3, padx=2, sticky="ew")

        # Third Column
        Label(column_frame, text="Third:", style="Section.TLabel").grid(
            row=0, column=4, sticky="w", pady=2, padx=2
        )
        self.filter3_frame = FuzzySearchFrame(
            column_frame, width=30, identifier="config_filter3"
        )
        self.filter3_frame.grid(row=0, column=5, padx=2, sticky="ew")

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
        self.source_folder_entry.insert(0, config["source_folder"])
        self.processed_folder_entry.insert(0, config["processed_folder"])
        self.excel_file_entry.insert(0, config["excel_file"])
        self.template_entry.insert(0, config["output_template"])

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

            # Update filters with original column values
            self.filter1_frame.set_values(columns)
            self.filter2_frame.set_values(columns)
            self.filter3_frame.set_values(columns)

            config = self.config_manager.get_config()

            # Compare using string versions
            if str(config.get("filter1_column")) in str_columns:
                self.filter1_frame.set(config["filter1_column"])
            if str(config.get("filter2_column")) in str_columns:
                self.filter2_frame.set(config["filter2_column"])
            if str(config.get("filter3_column")) in str_columns:
                self.filter3_frame.set(config["filter3_column"])

        except Exception as e:
            ErrorDialog(self, "Error Updating Columns", e)

    def save_config(self):
        """Save current configuration."""
        try:
            new_config = {
                "source_folder": self.source_folder_entry.get(),
                "processed_folder": self.processed_folder_entry.get(),
                "excel_file": self.excel_file_entry.get(),
                "excel_sheet": self.sheet_combobox.get(),
                "filter1_column": self.filter1_frame.get(),
                "filter2_column": self.filter2_frame.get(),
                "filter3_column": self.filter3_frame.get(),
                "dropdown1_column": self.filter1_frame.get(),  # Keep backward compatibility
                "dropdown2_column": self.filter2_frame.get(),  # Keep backward compatibility
                "output_template": self.template_entry.get(),
            }

            # Basic validation
            if not all(
                [
                    new_config["source_folder"],
                    new_config["processed_folder"],
                    new_config["excel_file"],
                    new_config["excel_sheet"],
                    new_config["filter1_column"],
                    new_config["filter2_column"],
                    new_config["filter3_column"],
                    new_config["output_template"],
                ]
            ):
                ErrorDialog(
                    self, "Validation Error", "Please fill in all required fields"
                )
                return

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
            self.sheet_combobox["values"] = ()
            self.sheet_combobox.set("")
            self.filter1_frame.clear()
            self.filter2_frame.clear()
            self.filter3_frame.clear()
            self.template_entry.delete(0, END)

            # Load preset values
            self.source_folder_entry.insert(0, preset_config.get("source_folder", ""))
            self.processed_folder_entry.insert(
                0, preset_config.get("processed_folder", "")
            )
            self.excel_file_entry.insert(0, preset_config.get("excel_file", ""))
            self.update_sheet_list(preset_config)
            self.template_entry.insert(0, preset_config.get("output_template", ""))
            self.filter1_frame.set(preset_config.get("filter1_column", ""))
            self.filter2_frame.set(preset_config.get("filter2_column", ""))
            self.filter3_frame.set(preset_config.get("filter3_column", ""))

            self.save_config()
            self.show_status_message(f"Loaded preset: {preset_name}")

        except Exception as e:
            ErrorDialog(self, "Load Preset Error", e)

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

            current_config = {
                "source_folder": self.source_folder_entry.get(),
                "processed_folder": self.processed_folder_entry.get(),
                "excel_file": self.excel_file_entry.get(),
                "excel_sheet": self.sheet_combobox.get(),
                "filter1_column": self.filter1_frame.get(),
                "filter2_column": self.filter2_frame.get(),
                "filter3_column": self.filter3_frame.get(),
                "output_template": self.template_entry.get(),
            }

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
        if not self.preset_var.get() in presets:
            self.preset_var.set("")

    def _on_frame_configure(self, event=None):
        """Handle frame resize"""
        # Update the canvas window width when the frame is resized
        self.canvas.itemconfig(
            "all", width=self.winfo_width() - self.scrollbar.winfo_width()
        )
