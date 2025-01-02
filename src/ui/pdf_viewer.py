from __future__ import annotations
from tkinter import (
    Event as TkEvent,
    Canvas as TkCanvas,
    Label as TkLabel,
    Widget as TkWidget,
)
from tkinter.ttk import (
    Frame as ttkFrame,
    Scrollbar as ttkScrollbar,
)
from PIL.ImageTk import PhotoImage as PILPhotoImage
from typing import Optional, Any
from .error_dialog import ErrorDialog


class PDFViewer(ttkFrame):
    """A modernized PDF viewer widget with zoom and scroll capabilities."""

    def __init__(self, master: TkWidget, pdf_manager: Any):
        super().__init__(master)
        self.pdf_manager = pdf_manager
        self.current_image: Optional[PILPhotoImage] = None
        self.current_pdf: Optional[str] = None
        self.zoom_level = 1.0

        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_ui()

    def setup_ui(self) -> None:
        """Setup the PDF viewer interface."""
        # Create a container frame with fixed padding for scrollbars
        self.container_frame = ttkFrame(self)
        self.container_frame.grid(row=0, column=0, sticky="nsew")
        self.container_frame.grid_columnconfigure(0, weight=1)
        self.container_frame.grid_rowconfigure(0, weight=1)

        # Create canvas with modern styling
        self.canvas = TkCanvas(
            self.container_frame,
            bg="#f8f9fa",  # Light gray background
            highlightthickness=0,  # Remove border
            width=20,  # Minimum width to prevent collapse
            height=20,  # Minimum height to prevent collapse
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Modern scrollbars - always create them to reserve space
        self.h_scrollbar = ttkScrollbar(
            self.container_frame, orient="horizontal", command=self.canvas.xview
        )
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.v_scrollbar = ttkScrollbar(
            self.container_frame, orient="vertical", command=self.canvas.yview
        )
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure canvas scrolling
        self.canvas.configure(
            xscrollcommand=self._on_x_scroll, yscrollcommand=self._on_y_scroll
        )

        # Initially hide scrollbars but keep their space reserved
        self.h_scrollbar.grid_remove()
        self.v_scrollbar.grid_remove()

        # Create a frame for the loading message that won't affect layout
        self.loading_frame = ttkFrame(self.container_frame)
        self.loading_frame.place(relx=0.5, rely=0.5, anchor="center")

        self._bind_events()

    def _on_x_scroll(self, *args) -> None:
        """Handle horizontal scrolling and scrollbar visibility."""
        self.h_scrollbar.set(*args)
        self._update_scrollbar_visibility()

    def _on_y_scroll(self, *args) -> None:
        """Handle vertical scrolling and scrollbar visibility."""
        self.v_scrollbar.set(*args)
        self._update_scrollbar_visibility()

    def _update_scrollbar_visibility(self) -> None:
        """Update scrollbar visibility based on content size."""
        if not self.current_image:
            self.h_scrollbar.grid_remove()
            self.v_scrollbar.grid_remove()
            return

        # Get the scroll region and canvas size
        x1, y1, x2, y2 = (
            self.canvas.bbox("all") if self.canvas.find_all() else (0, 0, 0, 0)
        )
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Show/hide horizontal scrollbar
        if x2 - x1 > canvas_width:
            self.h_scrollbar.grid()
        else:
            self.h_scrollbar.grid_remove()

        # Show/hide vertical scrollbar
        if y2 - y1 > canvas_height:
            self.v_scrollbar.grid()
        else:
            self.v_scrollbar.grid_remove()

    def _bind_events(self) -> None:
        """Bind mouse and keyboard events."""

        def _on_mousewheel(event: TkEvent) -> None:
            if event.state & 4:  # Ctrl key
                if event.delta > 0:
                    self.zoom_in()
                else:
                    self.zoom_out()
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event: TkEvent) -> None:
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event: TkEvent) -> None:
            self.canvas.unbind_all("<MouseWheel>")

        # Bind mousewheel only when mouse is over the PDF viewer area
        self.canvas.bind("<Enter>", _bind_mousewheel)
        self.canvas.bind("<Leave>", _unbind_mousewheel)
        self.v_scrollbar.bind("<Enter>", _bind_mousewheel)
        self.v_scrollbar.bind("<Leave>", _unbind_mousewheel)

        # Pan functionality
        self.canvas.bind("<Button-1>", self._start_drag)
        self.canvas.bind("<B1-Motion>", self._do_drag)
        self.canvas.bind("<ButtonRelease-1>", self._stop_drag)

        # Window resize handling
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Key>", self._on_key)

    def _start_drag(self, event: TkEvent) -> None:
        """Start panning the view."""
        self.canvas.scan_mark(event.x, event.y)
        self.canvas.configure(cursor="fleur")

    def _do_drag(self, event: TkEvent) -> None:
        """Continue panning the view."""
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _stop_drag(self, event: TkEvent) -> None:
        """Stop panning the view."""
        self.canvas.configure(cursor="")

    def _on_key(self, event: TkEvent) -> None:
        """Handle keyboard navigation."""
        key = event.keysym
        shift_pressed = event.state & 0x1

        if key == "Up":
            self.canvas.yview_scroll(-1 * (5 if shift_pressed else 1), "units")
        elif key == "Down":
            self.canvas.yview_scroll(1 * (5 if shift_pressed else 1), "units")
        elif key == "Left":
            self.canvas.xview_scroll(-1 * (5 if shift_pressed else 1), "units")
        elif key == "Right":
            self.canvas.xview_scroll(1 * (5 if shift_pressed else 1), "units")
        elif key == "Prior":  # Page Up
            self.canvas.yview_scroll(-1, "pages")
        elif key == "Next":  # Page Down
            self.canvas.yview_scroll(1, "pages")
        elif key == "Home":
            self.canvas.yview_moveto(0)
        elif key == "End":
            self.canvas.yview_moveto(1)

    def _on_resize(self, event: TkEvent) -> None:
        """Handle window resize events."""
        if event.widget == self.canvas:
            self._center_image()
            self._update_scrollbar_visibility()

    def _center_image(self) -> None:
        """Center the PDF image in the canvas."""
        if not self.current_image:
            return

        # Get dimensions
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        image_width = self.current_image.width()
        image_height = self.current_image.height()

        # Calculate centering offsets
        x = max(0, (canvas_width - image_width) // 2)
        y = max(0, (canvas_height - image_height) // 2)

        # Set scroll region to image bounds plus padding
        scroll_width = max(canvas_width, image_width + x * 2)
        scroll_height = max(canvas_height, image_height + y * 2)

        self.canvas.configure(scrollregion=(0, 0, scroll_width, scroll_height))

        # Clear and redraw image
        self.canvas.delete("all")
        image_x = (scroll_width - image_width) // 2
        image_y = (scroll_height - image_height) // 2
        self.canvas.create_image(
            image_x, image_y, anchor="nw", image=self.current_image
        )

        # Update scrollbar visibility
        self._update_scrollbar_visibility()

    def display_pdf(
        self, pdf_path: str, zoom: float = 1.0, show_loading: bool = True
    ) -> None:
        """Display a PDF file with the specified zoom level."""
        try:
            self.current_pdf = pdf_path
            self.zoom_level = zoom

            # Show loading message using place geometry manager
            loading_label = None
            if show_loading:
                loading_label = TkLabel(
                    self.loading_frame, text="Loading PDF...", font=("Segoe UI", 10)
                )
                loading_label.pack(pady=20)
                self.loading_frame.lift()  # Bring loading message to front
                self.update()

            # Render PDF
            image = self.pdf_manager.render_pdf_page(pdf_path, zoom=zoom)

            if loading_label:
                loading_label.destroy()
                self.loading_frame.place_forget()  # Hide the loading frame

            self.current_image = PILPhotoImage(image)
            self._center_image()
            self.canvas.focus_set()

        except Exception as e:
            if loading_label:
                loading_label.destroy()
                self.loading_frame.place_forget()  # Hide the loading frame in case of error
            ErrorDialog(self, "Error", f"Error displaying PDF: {str(e)}")

    def zoom_in(self, step: float = 0.2) -> None:
        """Zoom in the PDF view."""
        if self.current_pdf:
            self.zoom_level = min(3.0, self.zoom_level + step)
            self.display_pdf(self.current_pdf, self.zoom_level, show_loading=False)

    def zoom_out(self, step: float = 0.2) -> None:
        """Zoom out the PDF view."""
        if self.current_pdf:
            self.zoom_level = max(0.2, self.zoom_level - step)
            self.display_pdf(self.current_pdf, self.zoom_level, show_loading=False)
