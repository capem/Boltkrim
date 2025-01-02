from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
from datetime import datetime

@dataclass
class PDFTask:
    task_id: str  # Unique identifier for the task
    pdf_path: str
    value1: str
    value2: str
    value3: str
    status: str = "pending"  # pending, processing, failed, completed, reverted, skipped
    error_msg: str = ""
    row_idx: int = -1  # Add row index field
    original_excel_hyperlink: Optional[str] = None
    original_pdf_location: Optional[str] = None
    processed_pdf_location: Optional[str] = None
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None

    @staticmethod
    def generate_id() -> str:
        """Generate a unique task ID."""
        from uuid import uuid4
        return str(uuid4())

    def get_elapsed_time(self) -> str:
        """Calculate and format the elapsed time."""
        if not self.start_time:
            return "-"
        
        end = self.end_time if self.end_time else datetime.now()
        elapsed = end - self.start_time
        
        # Convert to total seconds
        total_seconds = int(elapsed.total_seconds())
        
        # Format as MM:SS
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{minutes:02d}:{seconds:02d}"