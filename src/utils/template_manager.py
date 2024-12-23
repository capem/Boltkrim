from datetime import datetime
import re
from typing import Any, Dict, Optional

class TemplateManager:
    """Manages template parsing and processing with support for various operations."""
    
    def __init__(self):
        self.date_operations = {
            'year': lambda dt: dt.strftime('%Y'),
            'month': lambda dt: dt.strftime('%m'),
            'year_month': lambda dt: dt.strftime('%Y-%m'),
            'format': lambda dt, fmt: dt.strftime(fmt.replace('%', ''))
        }
        
        self.string_operations = {
            'upper': str.upper,
            'lower': str.lower,
            'title': str.title,
            'replace': lambda s, old, new: s.replace(old, new),
            'slice': lambda s, start, end=None: s[int(start):None if end == '' else int(end)]
        }
    
    def _parse_field(self, field: str) -> tuple[str, list[str]]:
        """Parse a field into its name and operations."""
        parts = field.split('|')
        field_name = parts[0].strip()
        operations = parts[1:] if len(parts) > 1 else []
        return field_name, operations
    
    def _apply_date_operation(self, date_value: datetime, operation: str) -> str:
        """Apply a date operation to a datetime value."""
        op_parts = operation.split('.')
        if len(op_parts) != 2:
            raise ValueError(f"Invalid date operation format: {operation}")
        
        op_type = op_parts[1]
        if ':' in op_type:  # Handle format operation
            op_name, format_str = op_type.split(':', 1)
            if op_name not in self.date_operations:
                raise ValueError(f"Unknown date operation: {op_name}")
            return self.date_operations[op_name](date_value, format_str)
        else:
            if op_type not in self.date_operations:
                raise ValueError(f"Unknown date operation: {op_type}")
            return self.date_operations[op_type](date_value)
    
    def _apply_string_operation(self, value: str, operation: str) -> str:
        """Apply a string operation to a value."""
        op_parts = operation.split('.')
        if len(op_parts) != 2:
            raise ValueError(f"Invalid string operation format: {operation}")
        
        op_type = op_parts[1]
        if ':' in op_type:  # Handle operations with parameters
            op_name, *params = op_type.split(':')
            if op_name not in self.string_operations:
                raise ValueError(f"Unknown string operation: {op_name}")
            return self.string_operations[op_name](value, *params)
        else:
            if op_type not in self.string_operations:
                raise ValueError(f"Unknown string operation: {op_type}")
            return self.string_operations[op_type](value)
    
    def _apply_operations(self, value: Any, operations: list[str]) -> str:
        """Apply a sequence of operations to a value."""
        result = value
        for operation in operations:
            if operation.startswith('date.'):
                if not isinstance(value, datetime):
                    raise ValueError(f"Date operations can only be applied to datetime objects: {value}")
                result = self._apply_date_operation(value, operation)
            elif operation.startswith('str.'):
                result = self._apply_string_operation(str(result), operation)
            else:
                raise ValueError(f"Unknown operation type: {operation}")
        return str(result)
    
    def process_template(self, template: str, data: Dict[str, Any]) -> str:
        """
        Process a template string using the provided data.
        
        Args:
            template: The template string containing fields and operations
            data: Dictionary containing the values for the template fields
            
        Returns:
            The processed template with all fields replaced with their processed values
        """
        def replace_field(match: re.Match) -> str:
            field_content = match.group(1)
            field_name, operations = self._parse_field(field_content)
            
            if field_name not in data:
                raise ValueError(f"Field not found in data: {field_name}")
            
            value = data[field_name]
            return self._apply_operations(value, operations)
        
        pattern = r'\{([^}]+)\}'
        return re.sub(pattern, replace_field, template)