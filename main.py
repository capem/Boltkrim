import tkinter as tk
from tkinter import ttk
import os
from src.ui import ConfigTab, ProcessingTab
from src.utils import ConfigManager, ExcelManager, PDFManager

class FileOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Organizer")
        self.root.geometry("1200x800")
        
        # Initialize managers
        self.config_manager = ConfigManager()
        self.excel_manager = ExcelManager()
        self.pdf_manager = PDFManager()
        
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
        
        self.notebook.add(self.processing_tab, text='Processing')
        self.notebook.add(self.config_tab, text='Configuration')
        
        # Create status bar
        self.status_bar = ttk.Label(root, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Load configuration
        self.config_manager.load_config()
        
        # Load initial data if configuration exists
        config = self.config_manager.get_config()
        if config['excel_file'] and config['excel_sheet']:
            self.processing_tab.load_excel_data()
            
        # Load initial PDF if source folder exists
        if config['source_folder'] and os.path.exists(config['source_folder']):
            self.processing_tab.load_next_pdf()
            
        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.config_tab.save_config())
        self.root.bind('<Control-n>', lambda e: self.processing_tab.load_next_pdf())
        self.root.bind('<Return>', lambda e: self.processing_tab.process_current_file())
        self.root.bind('<Control-plus>', lambda e: self.processing_tab.zoom_in())
        self.root.bind('<Control-minus>', lambda e: self.processing_tab.zoom_out())

def main():
    root = tk.Tk()
    app = FileOrganizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
