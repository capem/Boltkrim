from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

@dataclass
class PDFTask:
    task_id: str  # Unique identifier for the task
    pdf_path: str
    value1: str
    value2: str
    value3: str
    status: str = "pending"  # pending, processing, failed, completed
    error_msg: str = ""
    row_idx: int = -1  # Add row index field
    original_excel_hyperlink: Optional[str] = None
    original_pdf_location: Optional[str] = None
    processed_pdf_location: Optional[str] = None
    @staticmethod
    def generate_id() -> str:
        """Generate a unique task ID."""
        from uuid import uuid4
        return str(uuid4())