# File Organizer

A Python application for organizing PDF files using Excel data.

## Project Structure

```
file_organizer/
├── src/
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── fuzzy_search.py
│   │   ├── error_dialog.py
│   │   ├── config_tab.py
│   │   └── processing_tab.py
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── config_manager.py
│   │   ├── excel_manager.py
│   │   └── pdf_manager.py
│   └── __init__.py
├── main.py
├── config.json
└── requirements.txt
```

## Components

- **UI Components** (`src/ui/`):
  - `fuzzy_search.py`: Implements fuzzy search functionality with text entry and listbox
  - `error_dialog.py`: Handles error message display
  - `config_tab.py`: Configuration interface
  - `processing_tab.py`: Main processing interface

- **Utilities** (`src/utils/`):
  - `config_manager.py`: Handles configuration loading/saving
  - `excel_manager.py`: Excel file operations
  - `pdf_manager.py`: PDF file operations

## Dependencies

- pandas (>=2.0.0): Data manipulation and Excel file handling
- PyMuPDF (>=1.22.0): PDF file handling
- openpyxl (>=3.1.0): Excel file operations
- fuzzywuzzy (>=0.18.0): Fuzzy string matching
- python-Levenshtein (>=0.21.0): Fast string matching

## Usage

1. Run `main.py` to start the application
2. Configure the application in the Configuration tab:
   - Set source and processed folders
   - Select Excel file and sheet
   - Choose columns for filtering
3. Use the Processing tab to:
   - View PDF files
   - Select appropriate filter values
   - Process files with new names

## Keyboard Shortcuts

- `Ctrl+S`: Save configuration
- `Right Arrow`: Load next PDF
- `Enter`: Process current file
- `Ctrl++`: Zoom in
- `Ctrl+-`: Zoom out
