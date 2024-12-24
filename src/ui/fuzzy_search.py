from tkinter import END, SINGLE, Widget, Event, Listbox
from tkinter.ttk import Frame, Entry, Scrollbar
from difflib import SequenceMatcher
from typing import List, Optional, Any

class FuzzySearchFrame(Frame):
    """A frame that provides fuzzy search functionality with a text entry and listbox.
    
    This widget allows users to search through a list of values using fuzzy matching,
    displaying the best matches in a scrollable listbox.
    """
    
    def __init__(
        self,
        master: Widget,
        values: Optional[List[str]] = None,
        search_threshold: int = 65,
        identifier: Optional[str] = None,
        **kwargs: Any
    ) -> None:
        super().__init__(master, **kwargs)
        
        self.all_values = [str(v) for v in (values or []) if v is not None]
        self.search_threshold = max(0, min(100, search_threshold))  # Clamp between 0 and 100
        self.identifier = identifier or 'unnamed'
        
        # Remove debouncing since we want instantaneous results
        self._prev_value = ''
        self._ignore_next_keyrelease = False
        self._focus_after_id: Optional[str] = None
        
        self._create_widgets()
        self._bind_events()
        self._update_listbox()
        
        self.after(100, self._ensure_focus)
        
    def _create_widgets(self) -> None:
        """Create and configure all child widgets."""
        # Entry widget
        self.entry = Entry(self)
        self.entry.pack(fill='x', padx=2, pady=2)
        
        # Bind focus-related events
        self.entry.bind('<FocusIn>', self._on_focus_in)
        self.entry.bind('<FocusOut>', self._on_focus_out)
        
        # Listbox frame with scrollbar
        listbox_frame = Frame(self)
        listbox_frame.pack(fill='both', expand=True, padx=2)
        
        # Listbox
        self.listbox = Listbox(
            listbox_frame,
            height=5,
            exportselection=False,
            selectmode=SINGLE
        )
        self.listbox.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        scrollbar = Scrollbar(
            listbox_frame,
            orient='vertical',
            command=self.listbox.yview
        )
        scrollbar.pack(side='right', fill='y')
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        # Configure mousewheel scrolling
        def _on_mousewheel(event: Event) -> str:
            self.listbox.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        
        def _bind_mousewheel(event: Event) -> None:
            self.listbox.bind_all("<MouseWheel>", _on_mousewheel)
            
        def _unbind_mousewheel(event: Event) -> None:
            self.listbox.unbind_all("<MouseWheel>")
        
        # Bind mousewheel only when mouse is over the listbox area
        self.listbox.bind('<Enter>', _bind_mousewheel)
        self.listbox.bind('<Leave>', _unbind_mousewheel)
        scrollbar.bind('<Enter>', _bind_mousewheel)
        scrollbar.bind('<Leave>', _unbind_mousewheel)
        
    def _bind_events(self) -> None:
        """Bind widget events to their handlers."""
        self.entry.bind('<KeyRelease>', self._on_keyrelease)
        self.listbox.bind('<<ListboxSelect>>', self._on_select)
        self.listbox.bind('<Button-1>', lambda e: self.after(50, self._on_select, e))
        self.listbox.bind('<Double-Button-1>', self._on_select)
        self.entry.bind('<Return>', lambda e: self._select_top_match())
        self.entry.bind('<Down>', lambda e: self._focus_listbox())
        self.entry.bind('<Tab>', self._handle_tab)
        
    def set_values(self, values: Optional[List[str]]) -> None:
        """Update the list of searchable values."""
        self.all_values = [str(v) for v in (values or []) if v is not None]
        self.entry.delete(0, END)
        self._update_listbox()
        
    def get(self) -> str:
        """Get the current entry text."""
        return self.entry.get()
        
    def set(self, value: str) -> None:
        """Set the entry text."""
        self.entry.delete(0, END)
        self.entry.insert(0, str(value))
        
    def _on_keyrelease(self, event: Event) -> None:
        """Handle key release events in the entry widget."""
        if self._ignore_next_keyrelease:
            self._ignore_next_keyrelease = False
            return
            
        # Update immediately without debouncing
        self._update_listbox()
        
    def _update_listbox(self) -> None:
        """Update the listbox with intelligent fuzzy search results."""
        current_value = self.entry.get().strip()
        
        # Clear current listbox
        self.listbox.delete(0, END)
        
        # If empty, show all values
        if not current_value:
            for value in self.all_values:
                self.listbox.insert(END, value)
            return
            
        try:
            search_lower = current_value.lower()
            
            # Calculate scores for all values
            scored_matches: List[tuple[int, str]] = []
            for value in self.all_values:
                value_lower = value.lower()
                score = 0
                
                # Exact match gets highest priority
                if value_lower == search_lower:
                    score = 100
                
                # Prefix match gets high priority
                elif value_lower.startswith(search_lower):
                    score = 90 - len(value)  # Shorter matches rank higher
                
                # Word boundary match
                elif any(word.startswith(search_lower) for word in value_lower.split()):
                    score = 80 - len(value)
                
                # Contains match
                elif search_lower in value_lower:
                    score = 70 - value_lower.index(search_lower)  # Earlier matches rank higher
                
                # Fuzzy match using sequence matcher
                else:
                    ratio = SequenceMatcher(None, search_lower, value_lower).ratio()
                    if ratio > 0.5:  # Only include if somewhat similar
                        score = int(ratio * 60)  # Max score of 60 for fuzzy matches
                
                if score > 0:
                    scored_matches.append((score, value))
            
            # Sort by score (highest first) and add to listbox
            scored_matches.sort(reverse=True, key=lambda x: (x[0], -len(x[1])))
            
            for _, value in scored_matches:
                self.listbox.insert(END, value)
                
        except Exception as e:
            print(f"Error in fuzzy search ({self.identifier}): {str(e)}")
            # Fall back to simple substring matching
            for value in self.all_values:
                if current_value.lower() in value.lower():
                    self.listbox.insert(END, value)
                    
    def _on_select(self, event: Event) -> None:
        """Handle selection events in the listbox."""
        if not self.listbox.size():  # If listbox is empty, do nothing
            return
            
        selection = self.listbox.curselection()
        if selection:
            value = self.listbox.get(selection[0])
            self._select_value(value)
            self.entry.focus_set()
            
    def _select_value(self, value: str) -> None:
        """Common method to handle value selection and event generation."""
        self._ignore_next_keyrelease = True
        self.set(value)
        self.event_generate('<<ValueSelected>>')
            
    def _select_top_match(self) -> None:
        """Select the top match in the listbox when Enter is pressed."""
        if self.listbox.size() > 0:
            self._select_value(self.listbox.get(0))
            
    def _focus_listbox(self) -> None:
        """Move focus to the listbox when Down arrow is pressed."""
        if self.listbox.size() > 0:
            # Clear any existing selection
            self.listbox.selection_clear(0, END)
            # Set both selection and active item to first item
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            self.listbox.focus_set()
            self.listbox.see(0)  # Ensure the selected item is visible
            
            # Bind keyboard events when listbox gets focus
            self.listbox.bind('<Up>', self._on_listbox_arrow)
            self.listbox.bind('<Down>', self._on_listbox_arrow)
            self.listbox.bind('<Tab>', self._on_listbox_tab)
            
    def _on_listbox_tab(self, event: Event) -> str:
        """Handle Tab key when pressed in the listbox."""
        active = self.listbox.index('active')
        if active >= 0:
            value = self.listbox.get(active)
            self._select_value(value)
            self.entry.focus_set()
            event.widget.tk_focusNext().focus()
        return 'break'
            
    def _on_listbox_arrow(self, event: Event) -> str:
        """Handle up/down arrow keys in listbox to maintain selection."""
        if event.keysym == 'Up' and self.listbox.index('active') > 0:
            new_index = self.listbox.index('active') - 1
            self.listbox.selection_clear(0, END)
            self.listbox.selection_set(new_index)
            self.listbox.activate(new_index)
            self.listbox.see(new_index)
        elif event.keysym == 'Down' and self.listbox.index('active') < self.listbox.size() - 1:
            new_index = self.listbox.index('active') + 1
            self.listbox.selection_clear(0, END)
            self.listbox.selection_set(new_index)
            self.listbox.activate(new_index)
            self.listbox.see(new_index)
        return 'break'
            
    def _ensure_focus(self) -> None:
        """Ensure the entry widget has focus."""
        if not self.entry.focus_get():
            self.entry.focus_force()
            # Schedule another check
            self._focus_after_id = self.after(100, self._ensure_focus)
            
    def _on_focus_in(self, event: Optional[Event] = None) -> None:
        """Handle focus-in event."""
        if self._focus_after_id:
            self.after_cancel(self._focus_after_id)
            self._focus_after_id = None
            
    def _on_focus_out(self, event: Optional[Event] = None) -> None:
        """Handle focus-out event."""
        # If we lose focus, try to get it back after a short delay
        # This helps with the initial focus issues
        self._focus_after_id = self.after(100, self._ensure_focus)
        
    def _handle_tab(self, event: Optional[Event] = None) -> Optional[str]:
        """Handle Tab key press in the entry widget."""
        if self.listbox.winfo_ismapped() and self.listbox.size() > 0:
            # If listbox is visible and has items, select the first one
            self._select_value(self.listbox.get(0))
            
        # Move to next widget
        if event:
            next_widget = event.widget.tk_focusNext()
            if isinstance(next_widget, Entry):
                # If next widget is an entry, focus it directly
                next_widget.focus_set()
            else:
                # Otherwise, follow normal tab order
                event.widget.tk_focusNext().focus()
            return "break"
        return None
