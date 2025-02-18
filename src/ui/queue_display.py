from __future__ import annotations
from tkinter import (
    Event as TkEvent,
    Widget as TkWidget,
    messagebox as TkMessagebox,
    Menu as TkMenu,
    Toplevel as TkTopLevel,
)
from tkinter.ttk import (
    Frame as ttkFrame,
    Scrollbar as ttkScrollbar,
    Label as ttkLabel,
    Style as ttkStyle,
    Treeview as ttkTreeview,
)

from typing import Dict
from datetime import datetime
from os import path
from ..utils import PDFTask

class QueueDisplay(ttkFrame):
    def __init__(self, master: TkWidget):
        super().__init__(master)
        self.setup_ui()

    def setup_ui(self) -> None:
        # Configure grid weights for responsive layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)  # For buttons

        print("[DEBUG] Setting up QueueDisplay UI")

        # Create a frame for the table and scrollbar with a light background
        table_frame = ttkFrame(self, style="Card.TFrame")
        table_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Setup table with more informative columns
        columns = ("task_id", "filename", "values", "status", "time")
        print("[DEBUG] Creating Treeview with columns:", columns)
        self.table = ttkTreeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",  # Single selection mode
            style="Queue.Treeview",
        )

        # Create context menu
        print("[DEBUG] Creating context menu")
        self.context_menu = TkMenu(self, tearoff=0)
        self.context_menu.add_command(label="Revert Task", command=self._on_revert_task)

        # Bind right-click to show context menu
        print("[DEBUG] Binding context menu events")
        self.table.bind(
            "<Button-3>", self._show_context_menu
        )  # Windows/Linux right-click
        self.table.bind("<Button-2>", self._show_context_menu)  # macOS right-click

        # Configure modern style for the treeview
        style = ttkStyle()
        style.configure(
            "Queue.Treeview",
            background="#ffffff",
            foreground="#333333",
            rowheight=30,
            fieldbackground="#ffffff",
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Queue.Treeview.Heading",
            background="#f0f0f0",
            foreground="#333333",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Queue.Treeview",
            background=[("selected", "#e7f3ff")],
            foreground=[("selected", "#000000")],
        )

        # Configure headings with sort functionality and modern look
        self.table.heading(
            "task_id",
            text="Task ID",
            anchor="w",
            command=lambda: self._sort_column("task_id"),
        )
        self.table.heading(
            "filename",
            text="File",
            anchor="w",
            command=lambda: self._sort_column("filename"),
        )
        self.table.heading(
            "values",
            text="Selected Values",
            anchor="w",
            command=lambda: self._sort_column("values"),
        )
        self.table.heading(
            "status",
            text="Status",
            anchor="w",
            command=lambda: self._sort_column("status"),
        )
        self.table.heading(
            "time", text="Time", anchor="w", command=lambda: self._sort_column("time")
        )

        # Configure column properties
        self.table.column(
            "task_id", width=0, minwidth=0, stretch=False
        )  # Hidden column for internal use
        self.table.column("filename", width=250, minwidth=150, stretch=True)
        self.table.column("values", width=250, minwidth=150, stretch=True)
        self.table.column("status", width=100, minwidth=80, stretch=False)
        self.table.column("time", width=80, minwidth=80, stretch=False)

        # Add modern-looking scrollbars
        style.configure(
            "Queue.Vertical.TScrollbar",
            background="#ffffff",
            troughcolor="#f0f0f0",
            width=10,
        )
        style.configure(
            "Queue.Horizontal.TScrollbar",
            background="#ffffff",
            troughcolor="#f0f0f0",
            width=10,
        )

        v_scrollbar = ttkScrollbar(
            table_frame,
            orient="vertical",
            command=self.table.yview,
            style="Queue.Vertical.TScrollbar",
        )
        self.table.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttkScrollbar(
            table_frame,
            orient="horizontal",
            command=self.table.xview,
            style="Queue.Horizontal.TScrollbar",
        )
        self.table.configure(xscrollcommand=h_scrollbar.set)

        # Grid table and scrollbars
        self.table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure status colors
        self.table.tag_configure("pending", foreground="#666666")
        self.table.tag_configure("processing", foreground="#007bff")
        self.table.tag_configure("completed", foreground="#28a745")
        self.table.tag_configure("failed", foreground="#dc3545")
        self.table.tag_configure("reverted", foreground="#6c757d")
        self.table.tag_configure("skipped", foreground="#ffc107")  # Yellow color for skipped files

        # Add status icons
        self.status_icons = {
            "pending": "⋯",  # Three dots
            "processing": "↻",  # Rotating arrow
            "completed": "✓",  # Checkmark
            "failed": "✗",  # X mark
            "reverted": "↺",  # Curved arrow for reverted status
            "skipped": "⤳",  # Arrow pointing right for skipped files
        }

        # Bind events for interactivity
        self.table.bind("<Double-1>", self._show_task_details)
        self.table.bind("<Return>", self._show_task_details)

    def _sort_column(self, column: str) -> None:
        """Sort treeview column."""
        # Get all items in the treeview
        items = [
            (self.table.set(item, column), item) for item in self.table.get_children("")
        ]

        # Check if we're reversing the sort
        reverse = False
        if hasattr(self, "_last_sort") and self._last_sort == (column, False):
            reverse = True

        # Store the sort state
        self._last_sort = (column, reverse)

        # Sort the items
        items.sort(reverse=reverse)

        # Rearrange items in sorted positions
        for index, (_, item) in enumerate(items):
            self.table.move(item, "", index)

        # Update the heading to show sort direction
        for col in self.table["columns"]:
            if col == column:
                self.table.heading(
                    col,
                    text=f"{self.table.heading(col)['text']} {'↓' if reverse else '↑'}",
                )
            else:
                # Remove sort indicator from other columns
                self.table.heading(
                    col, text=self.table.heading(col)["text"].split(" ")[0]
                )

    def _format_values_display(self, values_str: str) -> str:
        """Format the values string for better readability."""
        if not values_str or " | " not in values_str:
            return values_str

        parts = values_str.split(" | ")
        return " → ".join(parts)  # Using arrow for better visual flow

    def update_display(self, tasks: Dict[str, PDFTask]) -> None:
        """Update the queue display with the current tasks."""
        # Clear existing items
        for item in self.table.get_children():
            self.table.delete(item)

        # Add tasks to the table
        for task in tasks.values():
            # Format values for display
            values_display = self._format_values_display(" | ".join(task.filter_values))
            
            # Format time
            time_display = ""
            if task.end_time:
                duration = task.end_time - task.start_time
                time_display = f"{duration.seconds}s"
            elif task.start_time:
                duration = datetime.now() - task.start_time
                time_display = f"{duration.seconds}s"

            # Insert task into table
            self.table.insert(
                "",
                "end",
                values=(
                    task.task_id,
                    path.basename(task.pdf_path),
                    values_display,
                    f"{self.status_icons.get(task.status, '')} {task.status}",
                    time_display
                ),
                tags=(task.status,)
            )

    def _get_processing_tab(self):
        """Get the parent ProcessingTab instance by looping through parent widgets."""
        current_widget = self
        while current_widget:
            if str(current_widget.__class__.__name__) == "ProcessingTab":
                return current_widget
            current_widget = current_widget.master
        return None

    def _revert_task(self, task_id: str) -> None:
        """Handle reverting a task. Returns True if successful, False otherwise."""
        processing_tab = self._get_processing_tab()
        if not processing_tab:
            TkMessagebox.showerror("Error", "Internal error: Could not find processing tab.")
            return

        task = processing_tab.pdf_queue.get_task_by_id(task_id)
        if not task:
            TkMessagebox.showwarning("Cannot Revert", "Task not found.")
            return

        if task.status != "completed":
            TkMessagebox.showwarning("Cannot Revert", "Only completed tasks can be reverted.")
            return

        confirm = TkMessagebox.askyesno(
            "Confirm Revert",
            f"Are you sure you want to revert the task for '{path.basename(task.pdf_path)}'?"
        )
        if not confirm:
            return

        try:
            # Revert Excel hyperlink
            processing_tab.excel_manager.revert_pdf_link(
                excel_file=processing_tab.config_manager.get_config()["excel_file"],
                sheet_name=processing_tab.config_manager.get_config()["excel_sheet"],
                row_idx=task.row_idx,
                filter2_col=processing_tab.config_manager.get_config()["filter2_column"],
                original_hyperlink=task.original_excel_hyperlink,
                original_value=task.value2,
            )

            # Revert PDF location
            processing_tab.pdf_manager.revert_pdf_location(task=task)

            # Update task status
            processing_tab.pdf_queue.update_task_status(task_id, "reverted")

            TkMessagebox.showinfo(
                "Revert Successful",
                f"Task for '{path.basename(task.pdf_path)}' has been reverted successfully."
            )

        except Exception as e:
            print(f"[DEBUG] Revert failed: {str(e)}")
            import traceback
            print(traceback.format_exc())
            TkMessagebox.showerror("Revert Failed", f"Failed to revert the task: {str(e)}")

    def _show_task_details(self, event: TkEvent) -> None:
        """Show task details in a dialog when double-clicking a row."""
        item = self.table.identify("item", event.x, event.y)
        if not item:
            return

        # Get the task values
        values = self.table.item(item)["values"]
        if not values or len(values) < 4:  # Make sure we have all required values
            return

        # Create a dialog window
        dialog = TkTopLevel(self)
        dialog.title("Task Details")
        dialog.transient(self)  # Make dialog modal
        dialog.grab_set()  # Make dialog modal

        # Calculate position to center the dialog
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (400 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")

        # Create a frame with padding
        frame = ttkFrame(dialog, padding="20")
        frame.pack(fill="both", expand=True)

        # Add details with proper formatting and spacing
        details = [
            ("File:", path.basename(str(values[1]))),
            ("Full Path:", str(values[1])),
            ("Selected Values:", str(values[2])),
            ("Status:", str(values[3])),
            ("Time:", str(values[4]) if len(values) > 4 else ""),
        ]

        # Add each detail with proper styling
        for i, (label, value) in enumerate(details):
            ttkLabel(frame, text=label, font=("Segoe UI", 10, "bold")).grid(
                row=i, column=0, sticky="w", pady=(0, 10)
            )
            ttkLabel(frame, text=value, wraplength=250).grid(
                row=i, column=1, sticky="w", padx=(10, 0), pady=(0, 10)
            )

        # Add error message if status is failed
        if values[3] and "failed" in str(values[3]).lower():
            ttkLabel(frame, text="Error Message:", font=("Segoe UI", 10, "bold")).grid(
                row=len(details), column=0, sticky="w", pady=(10, 0)
            )

            # Get error message from the task
            task_id = values[0]  # First column is task_id
            error_msg = "No error message available"
            task = None

            processing_tab = self._get_processing_tab()
            if processing_tab and task_id:
                task = processing_tab.pdf_queue.get_task_by_id(task_id)
                if task and task.error_msg:
                    error_msg = task.error_msg

            ttkLabel(frame, text=error_msg, wraplength=250).grid(
                row=len(details) + 1, column=0, columnspan=2, sticky="w", pady=(0, 10)
            )

    def _show_context_menu(self, event):
        """Display the context menu on right-click."""
        print(f"[DEBUG] Right-click event at y={event.y}")
        selected_item = self.table.identify_row(event.y)
        print(f"[DEBUG] Identified row: {selected_item}")

        if selected_item:
            print(f"[DEBUG] Setting selection to: {selected_item}")
            self.table.selection_set(selected_item)
            print(f"[DEBUG] Current selection: {self.table.selection()}")
            print(f"[DEBUG] Posting menu at x={event.x_root}, y={event.y_root}")
            self.context_menu.post(event.x_root, event.y_root)
        else:
            print("[DEBUG] No row identified at click position")

    def _on_revert_task(self):
        """Handle the Revert Task action from the context menu."""
        selected_items = self.table.selection()
        if not selected_items:
            print("[DEBUG] No items selected")
            return

        selected_item = selected_items[0]
        item_values = self.table.item(selected_item)

        # Get the task ID from the first column
        task_id = item_values.get("values", [None])[0]
        if not task_id:
            print("[DEBUG] No task ID found")
            return

        self._revert_task(task_id)
