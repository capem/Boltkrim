import tkinter as tk
from tkinter import ttk

class FuzzySearchFrame(ttk.Frame):
    """A frame that provides fuzzy search functionality with a text entry and listbox.
    
    This widget allows users to search through a list of values using fuzzy matching,
    displaying the best matches in a scrollable listbox.
    """
    
    def __init__(
        self,
        master: tk.Widget,
        values: list[str] | None = None,
        search_threshold: int = 65,
        identifier: str | None = None,
        **kwargs
    ) -> None:
        super().__init__(master, **kwargs)
        
        self.all_values = [str(v) for v in (values or []) if v is not None]
        self.search_threshold = max(0, min(100, search_threshold))  # Clamp between 0 and 100
        self.identifier = identifier or 'unnamed'
        
        # Debouncing variables
        self._prev_value = ''
        self._after_id = None
        self._ignore_next_keyrelease = False
        self._debounce_delay = 0  # Changed from 150 to 0 to remove delay
        self._focus_after_id = None
        
        self._create_widgets()
        self._bind_events()
        self._update_listbox()
        
        # Schedule focus set after widget is fully realized
        self.after(100, self._ensure_focus)
        
    def _create_widgets(self) -> None:
        """Create and configure all child widgets."""
        # Entry widget
        self.entry = ttk.Entry(self)
        self.entry.pack(fill='x', padx=2, pady=2)
        
        # Bind focus-related events
        self.entry.bind('<FocusIn>', self._on_focus_in)
        self.entry.bind('<FocusOut>', self._on_focus_out)
        
        # Listbox frame with scrollbar
        listbox_frame = ttk.Frame(self)
        listbox_frame.pack(fill='both', expand=True, padx=2)
        
        # Listbox
        self.listbox = tk.Listbox(
            listbox_frame,
            height=5,
            exportselection=False,
            selectmode=tk.SINGLE
        )
        self.listbox.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(
            listbox_frame,
            orient='vertical',
            command=self.listbox.yview
        )
        scrollbar.pack(side='right', fill='y')
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
    def _bind_events(self) -> None:
        """Bind widget events to their handlers."""
        self.entry.bind('<KeyRelease>', self._on_keyrelease)
        self.listbox.bind('<<ListboxSelect>>', self._on_select)
        self.listbox.bind('<Button-1>', lambda e: self.after(50, self._on_select, e))
        self.listbox.bind('<Double-Button-1>', self._on_select)
        self.entry.bind('<Return>', lambda e: self._select_top_match())
        self.entry.bind('<Down>', lambda e: self._focus_listbox())
        self.entry.bind('<Tab>', self._handle_tab)
        
    def set_values(self, values: list[str] | None) -> None:
        """Update the list of searchable values."""
        self.all_values = [str(v) for v in (values or []) if v is not None]
        self.entry.delete(0, tk.END)
        self._update_listbox()
        
    def get(self) -> str:
        """Get the current entry text."""
        return self.entry.get()
        
    def set(self, value: str) -> None:
        """Set the entry text."""
        self.entry.delete(0, tk.END)
        self.entry.insert(0, str(value))
        
    def _on_keyrelease(self, event: tk.Event) -> None:
        """Handle key release events in the entry widget."""
        if self._ignore_next_keyrelease:
            self._ignore_next_keyrelease = False
            return
            
        # Update immediately without debouncing
        self._update_listbox()
        
    def _update_listbox(self) -> None:
        """Update the listbox with fuzzy search results."""
        current_value = self.entry.get().strip()
        
        # Clear current listbox
        self.listbox.delete(0, tk.END)
        
        # If empty, show all values
        if not current_value:
            for value in self.all_values:
                self.listbox.insert(tk.END, value)
            return
            
        try:
            # Convert search term to lowercase for case-insensitive matching
            search_lower = current_value.lower()
            
            # 1. Exact matches (case-insensitive)
            exact_matches = [v for v in self.all_values if v.lower() == search_lower]
            
            # 2. Prefix matches (prioritized by length)
            prefix_matches = [
                v for v in self.all_values 
                if v.lower().startswith(search_lower) 
                and v not in exact_matches
            ]
            # Sort prefix matches by length (shorter first) then alphabetically
            prefix_matches.sort(key=lambda x: (len(x), x.lower()))
            
            # 3. Contains matches (words that contain the search term)
            contains_matches = [
                v for v in self.all_values 
                if search_lower in v.lower() 
                and v not in exact_matches 
                and v not in prefix_matches
            ]
            
            # Add matches to listbox in priority order
            # 1. Exact matches
            for match in exact_matches:
                self.listbox.insert(tk.END, match)
                
            # 2. Prefix matches
            for match in prefix_matches:
                self.listbox.insert(tk.END, match)
                
            # 3. Contains matches
            for match in contains_matches:
                self.listbox.insert(tk.END, match)
                
        except Exception as e:
            print(f"Error in fuzzy search ({self.identifier}): {str(e)}")
            # Fall back to simple substring matching
            for value in self.all_values:
                if current_value.lower() in value.lower():
                    self.listbox.insert(tk.END, value)
                    
    def _on_select(self, event: tk.Event) -> None:
        """Handle selection events in the listbox."""
        if not self.listbox.size():  # If listbox is empty, do nothing
            return
            
        selection = self.listbox.curselection()
        if selection:
            value = self.listbox.get(selection[0])
            self._ignore_next_keyrelease = True
            self.set(value)
            # Generate a virtual event that can be bound by parent widgets
            self.event_generate('<<ValueSelected>>')
            self.entry.focus_set()
            
    def _select_top_match(self) -> None:
        """Select the top match in the listbox when Enter is pressed."""
        if self.listbox.size() > 0:
            value = self.listbox.get(0)
            self._ignore_next_keyrelease = True
            self.set(value)
            
    def _focus_listbox(self) -> None:
        """Move focus to the listbox when Down arrow is pressed."""
        if self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.focus_set()
            
    def _ensure_focus(self) -> None:
        """Ensure the entry widget has focus."""
        if not self.entry.focus_get():
            self.entry.focus_force()
            # Schedule another check
            self._focus_after_id = self.after(100, self._ensure_focus)
            
    def _on_focus_in(self, event=None) -> None:
        """Handle focus-in event."""
        if self._focus_after_id:
            self.after_cancel(self._focus_after_id)
            self._focus_after_id = None
            
    def _on_focus_out(self, event=None) -> None:
        """Handle focus-out event."""
        # If we lose focus, try to get it back after a short delay
        # This helps with the initial focus issues
        self._focus_after_id = self.after(100, self._ensure_focus)
        
    def _handle_tab(self, event=None) -> None:
        """Handle Tab key press: select first result and move to next widget."""
        if self.listbox.size() > 0:
            # Select the first item
            value = self.listbox.get(0)
            self._ignore_next_keyrelease = True
            self.set(value)
            # Generate the ValueSelected event
            self.event_generate('<<ValueSelected>>')
            
        # Prevent default Tab behavior
        if event:
            event.widget.tk_focusNext().focus()
            return "break"
