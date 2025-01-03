from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional, List
from datetime import datetime

@dataclass
class PDFTask:
    task_id: str  # Unique identifier for the task
    pdf_path: str
    filter_values: List[str] = field(default_factory=list)  # Dynamic list of filter values
    status: str = "pending"  # pending, processing, failed, completed, reverted, skipped
    error_msg: str = ""
    row_idx: int = -1  # Add row index field
    original_excel_hyperlink: Optional[str] = None
    original_pdf_location: Optional[str] = None
    processed_pdf_location: Optional[str] = None
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None

    def __post_init__(self):
        """Initialize start time if not provided."""
        if not self.start_time:
            self.start_time = datetime.now()

    @property
    def value1(self) -> str:
        """Backward compatibility for first filter value."""
        return self.filter_values[0] if self.filter_values else ""

    @property
    def value2(self) -> str:
        """Backward compatibility for second filter value."""
        return self.filter_values[1] if len(self.filter_values) > 1 else ""

    @property
    def value3(self) -> str:
        """Backward compatibility for third filter value."""
        return self.filter_values[2] if len(self.filter_values) > 2 else ""

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