from __future__ import annotations
from tkinter import Tk, SUNKEN, DISABLED, NORMAL, X, BOTTOM, LEFT, RIGHT, W
from tkinter.ttk import Notebook, Frame, Label, Button
from tkinter.messagebox import showerror
from typing import Optional
from src.ui import ConfigTab, ProcessingTab
from src.utils import ConfigManager, ExcelManager, PDFManager

class FileOrganizerApp:
    """Main application class for the File Organizer application."""
    
    def __init__(self, root: Tk) -> None:
        """Initialize the application.
        
        Args:
            root: The root Tk window
        """
        self.root = root
        self.root.title("File Organizer")
        self.root.state('zoomed')
        
        # Initialize managers
        self.config_manager = ConfigManager()
        self.excel_manager = ExcelManager()
        self.pdf_manager = PDFManager()
        
        # Load configuration first
        self.config_manager.load_config()
        
        self._setup_ui()
        self._bind_shortcuts()
        
        # Schedule data loading after window is shown
        self.root.after(100, self.load_initial_data)
    
    def _setup_ui(self) -> None:
        """Setup the main UI components."""
        # Create notebook for tabs
        self.notebook = Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create tabs
        self.processing_tab = ProcessingTab(
            self.notebook,
            self.config_manager,
            self.excel_manager,
            self.pdf_manager
        )
        self.config_tab = ConfigTab(
            self.notebook,
            self.config_manager,
            self.excel_manager
        )
        
        # Register callback for config changes
        self.config_tab.add_config_change_callback(self.on_config_change)
        
        self.notebook.add(self.processing_tab, text='Processing')
        self.notebook.add(self.config_tab, text='Configuration')
        
        # Create status bar with retry button
        self._setup_status_bar()

    def _setup_status_bar(self) -> None:
        """Setup the status bar and retry button."""
        self.status_frame = Frame(self.root)
        self.status_frame.pack(side=BOTTOM, fill=X)
        
        self.status_bar = Label(
            self.status_frame,
            text="Loading...",
            relief=SUNKEN,
            anchor=W
        )
        self.status_bar.pack(side=LEFT, fill=X, expand=True)
        
        self.retry_button = Button(
            self.status_frame,
            text="Retry",
            command=self.retry_load_data,
            state=DISABLED
        )
        self.retry_button.pack(side=RIGHT, padx=5)

    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts to actions."""
        shortcuts = {
            '<F5>': lambda e: self.retry_load_data()
        }
        for key, callback in shortcuts.items():
            self.root.bind(key, callback)

    def _handle_error(self, error: Exception, operation: str) -> None:
        """Centralized error handling.
        
        Args:
            error: The exception that occurred
            operation: Description of the operation that failed
        """
        if "Network" in str(error):
            error_msg = "Network error: Cannot access files"
            detail_msg = "Cannot access network files. Please check your network connection and try again."
        else:
            error_msg = f"Error during {operation}: {str(error)}"
            detail_msg = str(error)
        
        self.status_bar['text'] = error_msg
        self.retry_button.configure(state=NORMAL)
        showerror("Error", detail_msg)

    def safe_load_next_pdf(self) -> None:
        """Safely load next PDF with error handling."""
        try:
            self.processing_tab.load_next_pdf()
            self.retry_button.configure(state=DISABLED)
            self.status_bar['text'] = "Ready"
        except Exception as e:
            self._handle_error(e, "loading PDF")
    
    def load_initial_data(self) -> None:
        """Load initial data asynchronously after window is shown."""
        try:
            config = self.config_manager.get_config()
            if config['excel_file'] and config['excel_sheet']:
                self.processing_tab.load_excel_data()
                
            if config['source_folder']:
                self.safe_load_next_pdf()
                
            self.status_bar['text'] = "Ready"
            self.retry_button.configure(state=DISABLED)
        except Exception as e:
            self._handle_error(e, "initial data load")

    def retry_load_data(self) -> None:
        """Retry loading data after an error."""
        self.status_bar['text'] = "Retrying..."
        self.retry_button.configure(state=DISABLED)
        self.root.after(100, self.load_initial_data)

    def on_config_change(self) -> None:
        """Handle configuration changes."""
        self.status_bar['text'] = "Loading..."
        self.retry_button.configure(state=DISABLED)
        self.root.after(100, self.load_initial_data)

def main() -> None:
    """Main entry point for the application."""
    root = Tk()
    app = FileOrganizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
