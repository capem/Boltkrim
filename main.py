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
        
        # Schedule data loading after window is shown
        self.root.after(100, self.processing_tab.load_initial_data)
    
    def _setup_ui(self) -> None:
        """Setup the main UI components."""
        # Create notebook for tabs
        self.notebook = Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create status bar first so it's available for error handling
        self._setup_status_bar()
        
        # Create tabs
        self.processing_tab = ProcessingTab(
            self.notebook,
            self.config_manager,
            self.excel_manager,
            self.pdf_manager,
            self._handle_error,  # Pass error handler to tab
            self.update_status   # Pass status handler to tab
        )
        self.config_tab = ConfigTab(
            self.notebook,
            self.config_manager,
            self.excel_manager
        )
        
        self.notebook.add(self.processing_tab, text='Processing')
        self.notebook.add(self.config_tab, text='Configuration')

    def _setup_status_bar(self) -> None:
        """Setup the status bar."""
        self.status_frame = Frame(self.root)
        self.status_frame.pack(side=BOTTOM, fill=X)
        
        self.status_bar = Label(
            self.status_frame,
            text="Loading...",
            relief=SUNKEN,
            anchor=W
        )
        self.status_bar.pack(side=LEFT, fill=X, expand=True)

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
        showerror("Error", detail_msg)

    def update_status(self, message: str) -> None:
        """Update the status bar message.
        
        Args:
            message: The message to display
        """
        self.status_bar['text'] = message

def main() -> None:
    """Main entry point for the application."""
    root = Tk()
    app = FileOrganizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
