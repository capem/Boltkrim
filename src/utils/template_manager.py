from datetime import datetime
from re import sub, Match
from typing import Any, Dict, Optional, Callable, Union

class TemplateManager:
    """Manages template parsing and processing with support for various operations."""
    
    def __init__(self) -> None:
        self.date_operations: Dict[str, Callable] = {
            'year': lambda dt: dt.strftime('%Y'),
            'month': lambda dt: dt.strftime('%m'),
            'year_month': lambda dt: dt.strftime('%Y-%m'),
            'format': lambda dt, fmt: dt.strftime(fmt.replace('%', ''))
        }
        
        def sanitize_path(s: str) -> str:
            """Sanitize a string to be safe for use in file paths.
            Handles various special characters while preserving basic readability."""
            # Characters that are problematic in file paths
            replacements = {
                '/': '_',    # Forward slash
                '\\': '_',   # Backslash
                ':': '-',    # Colon
                '*': '+',    # Asterisk
                '?': '',     # Question mark
                '"': "'",    # Double quote
                '<': '(',    # Less than
                '>': ')',    # Greater than
                '|': '-',    # Pipe
                '\0': '',    # Null character
                '\n': ' ',   # Newline
                '\r': ' ',   # Carriage return
                '\t': ' ',   # Tab
            }
            result = s
            for char, replacement in replacements.items():
                result = result.replace(char, replacement)
            # Remove any leading/trailing whitespace and dots
            result = result.strip('. ')
            # Collapse multiple spaces into one
            result = ' '.join(result.split())
            return result
        
        def get_first_word(s: str) -> str:
            """Get the first word from a string."""
            return s.split()[0] if s else ''
            
        def split_by_no_get_last(s: str) -> str:
            """Split string by N° and get the last element, preserving the N° prefix."""
            if 'N°' in s:
                parts = s.split('N°')
                return f"N°{parts[-1].strip()}"
            return s.strip()
        
        self.string_operations: Dict[str, Callable] = {
            'upper': str.upper,
            'lower': str.lower,
            'title': str.title,
            'replace': lambda s, old, new: s.replace(old, new),
            'slice': lambda s, start, end=None: s[int(start):None if end == '' else int(end)],
            'sanitize': sanitize_path,
            'first_word': get_first_word,
            'split_no_last': split_by_no_get_last
        }
    
    def _parse_field(self, field: str) -> tuple[str, list[str]]:
        """Parse a field into its name and operations.
        
        Args:
            field: The field string to parse
            
        Returns:
            A tuple containing the field name and list of operations
        """
        parts = field.split('|')
        field_name = parts[0].strip()
        operations = parts[1:] if len(parts) > 1 else []
        return field_name, operations
    
    def _apply_date_operation(self, date_value: datetime, operation: str) -> str:
        """Apply a date operation to a datetime value.
        
        Args:
            date_value: The datetime value to operate on
            operation: The operation string to apply
            
        Returns:
            The result of applying the date operation
            
        Raises:
            ValueError: If the operation format is invalid or unknown
        """
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
        """Apply a string operation to a value.
        
        Args:
            value: The string value to operate on
            operation: The operation string to apply
            
        Returns:
            The result of applying the string operation
            
        Raises:
            ValueError: If the operation format is invalid or unknown
        """
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
        """Apply a sequence of operations to a value.
        
        Args:
            value: The value to operate on
            operations: List of operations to apply
            
        Returns:
            The result of applying all operations in sequence
            
        Raises:
            ValueError: If an operation type is unknown or invalid for the value type
        """
        result = value
        for operation in operations:
            if operation.startswith('date.'):
                # If it's already a datetime object, use it directly
                if isinstance(value, datetime):
                    result = self._apply_date_operation(value, operation)
                # Otherwise try to parse it if it's a string
                elif isinstance(value, str):
                    try:
                        # Try to parse the date string in common formats
                        for fmt in ['%d_%m_%Y', '%Y-%m-%d', '%d/%m/%Y']:
                            try:
                                parsed_date = datetime.strptime(value, fmt)
                                result = self._apply_date_operation(parsed_date, operation)
                                break
                            except ValueError:
                                continue
                        else:
                            raise ValueError(f"Could not parse date string: {value}")
                    except Exception as e:
                        raise ValueError(f"Could not convert string to date: {value} - {str(e)}")
                else:
                    raise ValueError(f"Date operations can only be applied to datetime objects or date strings: {value}")
            elif operation.startswith('str.'):
                result = self._apply_string_operation(str(result), operation)
            else:
                raise ValueError(f"Unknown operation type: {operation}")
        return str(result)
    
    def process_template(self, template: str, data: Dict[str, Any]) -> str:
        """Process a template string using the provided data.
        
        Args:
            template: The template string containing fields and operations
            data: Dictionary containing the values for the template fields
            
        Returns:
            The processed template with all fields replaced with their processed values
            
        Raises:
            ValueError: If a field is not found in the data or an operation fails
        """
        def replace_field(match: Match) -> str:
            field_content = match.group(1)
            field_name, operations = self._parse_field(field_content)
            
            if field_name not in data:
                raise ValueError(f"Field not found in data: {field_name}")
            
            value = data[field_name]
            
            if field_name != 'processed_folder' and isinstance(value, str):
                value = self.string_operations['sanitize'](value)
            
            return self._apply_operations(value, operations)
        
        pattern = r'\{([^}]+)\}'
        return sub(pattern, replace_field, template)