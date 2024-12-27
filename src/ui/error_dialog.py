import tkinter as tk
from tkinter import ttk, scrolledtext
import traceback

class ErrorDialog(tk.Toplevel):
    def __init__(self, parent, title, error, show_traceback=True):
        super().__init__(parent)
        self.title(title)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Calculate position
        window_width = 500
        window_height = 300
        screen_width = parent.winfo_screenwidth()
        screen_height = parent.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        
        # Set size and position
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.minsize(window_width, window_height)
        
        # Format error message
        if isinstance(error, Exception) and show_traceback:
            message = f"Error: {str(error)}\n\nTraceback:\n{traceback.format_exc()}"
        elif isinstance(error, Exception):
            message = f"Error: {str(error)}"
        else:
            message = str(error)
        
        # Add message
        message_frame = ttk.Frame(self)
        message_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Add scrolled text widget for error message
        self.text_widget = scrolledtext.ScrolledText(
            message_frame, 
            wrap=tk.WORD, 
            width=50, 
            height=10
        )
        self.text_widget.pack(fill='both', expand=True, padx=5, pady=5)
        self.text_widget.insert('1.0', message)
        self.text_widget.configure(state='disabled')
        
        # Add buttons
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        copy_button = ttk.Button(
            button_frame, 
            text="Copy to Clipboard",
            command=self.copy_to_clipboard
        )
        copy_button.pack(side='left', padx=5)
        
        close_button = ttk.Button(
            button_frame, 
            text="Close",
            command=self.destroy
        )
        close_button.pack(side='right', padx=5)
        
        # Bind escape key to close
        self.bind('<Escape>', lambda e: self.destroy())
        
        # Center the dialog
        self.update_idletasks()
        
        # Make dialog resizable
        self.resizable(True, True)
        
    def copy_to_clipboard(self):
        error_text = self.text_widget.get('1.0', tk.END)
        self.clipboard_clear()
        self.clipboard_append(error_text)
