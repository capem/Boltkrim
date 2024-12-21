import tkinter as tk
from tkinter import ttk, messagebox
import os
from src.ui import ConfigTab, ProcessingTab
from src.utils import ConfigManager, ExcelManager, PDFManager

class FileOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Organizer")
        self.root.state('zoomed')
        
        # Initialize managers
        self.config_manager = ConfigManager()
        self.excel_manager = ExcelManager()
        self.pdf_manager = PDFManager()
        
        # Load configuration first
        self.config_manager.load_config()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
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
        self.status_frame = ttk.Frame(root)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_bar = ttk.Label(
            self.status_frame,
            text="Loading...",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.retry_button = ttk.Button(
            self.status_frame,
            text="Retry",
            command=self.retry_load_data,
            state=tk.DISABLED
        )
        self.retry_button.pack(side=tk.RIGHT, padx=5)
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.config_tab.save_config())
        self.root.bind('<Control-n>', lambda e: self.safe_load_next_pdf())
        self.root.bind('<Return>', lambda e: self.processing_tab.process_current_file())
        self.root.bind('<Control-plus>', lambda e: self.processing_tab.zoom_in())
        self.root.bind('<Control-minus>', lambda e: self.processing_tab.zoom_out())
        self.root.bind('<F5>', lambda e: self.retry_load_data())
        
        # Schedule data loading after window is shown
        self.root.after(100, self.load_initial_data)
    
    def safe_load_next_pdf(self):
        """Safely load next PDF with error handling."""
        try:
            self.processing_tab.load_next_pdf()
            self.retry_button.configure(state=tk.DISABLED)
            self.status_bar['text'] = "Ready"
        except Exception as e:
            if "Network" in str(e):
                self.status_bar['text'] = "Network error: Cannot access PDF files"
                self.retry_button.configure(state=tk.NORMAL)
                messagebox.showerror("Network Error", 
                    "Cannot access network files. Please check your network connection and try again.")
            else:
                self.status_bar['text'] = f"Error: {str(e)}"
                self.retry_button.configure(state=tk.NORMAL)
    
    def load_initial_data(self):
        """Load initial data asynchronously after window is shown."""
        try:
            config = self.config_manager.get_config()
            if config['excel_file'] and config['excel_sheet']:
                self.processing_tab.load_excel_data()
                
            # Load initial PDF if source folder exists
            if config['source_folder']:
                self.safe_load_next_pdf()
                
            self.status_bar['text'] = "Ready"
            self.retry_button.configure(state=tk.DISABLED)
        except Exception as e:
            if "Network" in str(e):
                self.status_bar['text'] = "Network error: Cannot access files"
                self.retry_button.configure(state=tk.NORMAL)
                messagebox.showerror("Network Error", 
                    "Cannot access network files. Please check your network connection and try again.")
            else:
                self.status_bar['text'] = f"Error: {str(e)}"
                self.retry_button.configure(state=tk.NORMAL)

    def retry_load_data(self):
        """Retry loading data after a network error."""
        self.status_bar['text'] = "Retrying..."
        self.retry_button.configure(state=tk.DISABLED)
        self.root.after(100, self.load_initial_data)

    def on_config_change(self):
        """Handle configuration changes."""
        self.status_bar['text'] = "Loading..."
        self.retry_button.configure(state=tk.DISABLED)
        self.root.after(100, self.load_initial_data)

def main():
    root = tk.Tk()
    app = FileOrganizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
