# File Organizer

## Overview

File Organizer is a desktop application designed to automate the process of organizing and processing PDF files. It leverages data from Excel files and user-defined templates to dynamically rename and structure PDF documents, streamlining file management workflows.

## Key Features

- **Configuration Tab:**  Easily configure application settings, including source and processed folders, Excel file and sheet selection, output templates, and data filters.
- **Processing Tab:** Manage a queue of PDF files for processing, apply configured settings, and monitor task statuses.
- **Integrated PDF Viewer:** Review PDF files directly within the application with zoom and rotation controls.
- **Excel Data Integration:** Utilize data from Excel spreadsheets to dynamically name and organize PDF files based on flexible templates.
- **Template Engine:** Define output paths and filenames using a powerful template syntax with support for date and string operations.
- **Dynamic Filters:** Configure multiple filters based on Excel columns to precisely match and process PDF files with relevant data.
- **Processing Queue:** Track the status of PDF processing tasks in real-time, with options to retry failed tasks or clear completed ones.

## Installation

### Prerequisites

- [Python](https://www.python.org) 3.13 or higher
- [uv](https://astral.sh/blog/uv) package installer

### Dependencies

The application relies on the following Python packages, which are specified in `pyproject.toml`:

- openpyxl>=3.1.5
- pandas>=2.2.3
- pillow>=11.1.0
- pyinstaller>=6.12.0
- pymupdf>=1.25.3
- pywin32>=308

### Steps

1. **Install Python:** Ensure you have Python 3.13 or a later version installed on your system. You can download it from the official Python website.
2. **Install uv:** Follow the instructions in the [uv documentation](https://astral.sh/blog/uv) to install the uv package installer.
3. **Clone the Repository:** Clone the project repository to your local machine.
4. **Navigate to the Project Directory:** Open a terminal or command prompt and navigate to the cloned project directory.
5. **Create a virtual environment:** Run the following command to create a virtual environment using uv:

   ```bash
   uv venv
   ```

6. **Activate the virtual environment:** Activate the virtual environment based on your operating system (e.g., `source .venv/bin/activate` on Linux/macOS, `.venv\Scripts\activate` on Windows).
7. **Install Dependencies:** Run the following command to synchronize your environment with the project's lock file using uv:

   ```bash
   uv sync
   ```
   This command ensures that the dependencies are installed according to the `uv.lock` file, providing reproducible builds. **Note:** If the `uv.lock` file is not present in the project, `uv sync` will resolve the dependencies based on `pyproject.toml` and create a new `uv.lock` file to ensure consistent builds in the future.
## Usage

1. **Configuration:**
   - Open the "Configuration" tab within the application.
   - **Folder Settings:**
     - **Source:** Specify the folder containing the PDF files you want to organize.
     - **Processed:**  Define the destination folder where processed PDFs will be saved.
   - **Excel Configuration:**
     - **File:** Select the Excel file that contains the data for processing.
     - **Sheet:** Choose the specific sheet within the Excel file to use.
   - **Column Configuration:** Configure up to three filters to match columns in your Excel sheet with data in your PDF files. These filters are used to identify the correct rows in your Excel data for each PDF.
   - **Output Template:** Define a template for naming and organizing your processed PDF files. Use placeholders like `{field_name}` to insert data from your Excel file.

     **Template Syntax Guide:**

     - **Basic fields:** `{field_name}` -  Inserts the value of the specified Excel column.
     - **Multiple operations:** `{field|operation1|operation2}` - Applies a series of operations to the field value.

     **Date Operations:**

     - **Extract year:** `{DATE FACTURE|date.year}` - Extracts the year from a date field.
     - **Extract month:** `{DATE FACTURE|date.month}` - Extracts the month from a date field.
     - **Year-Month:** `{DATE FACTURE|date.year_month}` - Formats date as YYYY-MM.
     - **Custom format:** `{DATE FACTURE|date.format:%Y/%m}` -  Formats date using custom specifiers (e.g., `%Y/%m` for YYYY/MM).

     **String Operations:**

     - **Uppercase:** `{field|str.upper}` - Converts the field value to uppercase.
     - **Lowercase:** `{field|str.lower}` - Converts the field value to lowercase.
     - **Title Case:** `{field|str.title}` - Converts the field value to title case.
     - **Replace:** `{field|str.replace:old:new}` - Replaces occurrences of "old" with "new" in the field value.
     - **Slice:** `{field|str.slice:0:4}` - Extracts a substring from the field value (e.g., first 4 characters).

   - Click "Save Configuration" to store your settings. You can also save and load configurations as Presets for easy reuse.

2. **Processing:**
   - Navigate to the "Processing" tab.
   - The application will automatically attempt to load the first PDF file from the configured "Source Folder."
   - **File Information:** The left panel displays information about the currently loaded PDF file. Click the file name to manually select a PDF file from the source folder.
   - **PDF Viewer:** The center panel shows the center panel shows the currently loaded PDF. Use the zoom controls (+/-) and rotation controls (↶/↷) to adjust the view as needed.
   - **Filters:** The right panel contains filter input fields based on your "Column Configuration." Select values in these filters to match rows in your Excel data. The filters are dynamically populated based on the Excel data and are used to find the correct data row for processing the current PDF.
   - **Process File:** Once you have selected values for all filters, click the "Process File" button. This will:
     - Process the current PDF file using the configured settings and selected filter data.
     - Rename the PDF file based on the "Output Template."
     - Move the processed PDF to the "Processed Folder."
     - Update the Excel file with a hyperlink to the processed PDF (if configured).
     - Automatically load the next PDF file from the "Source Folder."
     - **Next File (Ctrl+N):** If you wish to skip processing the current PDF, click the "Next File (Ctrl+N)" button or use the keyboard shortcut `Ctrl+N`. This will move the current PDF to a "skipped" folder (if configured) and load the next PDF file.
     - **Processing Queue:** The "Processing Queue" in the left panel displays a list of PDF files that are pending, processing, completed, or failed. You can monitor the status of your tasks in this queue.
     - **Clear Completed:** Click "Clear Completed" to remove successfully processed tasks from the queue display.
     - **Retry Failed:** Click "Retry Failed" to reset the status of failed tasks to "pending" and re-attempt processing them.

## Building Executable

To build an executable version of the File Organizer application, you can use PyInstaller. PyInstaller is already included as a dependency in `pyproject.toml`, so it will be installed when you run `uv sync`.

Follow these steps to build the executable:

1. **Navigate to the Project Directory:** Open a terminal or command prompt and navigate to the root directory of the project, where the `FileOrganizer.spec` file is located.
2. **Run PyInstaller:** Execute the following command to build the executable:

   ```bash
   pyinstaller FileOrganizer.spec --clean
   ```

   This command does the following:
   - `pyinstaller`: Invokes the PyInstaller tool.
   - `FileOrganizer.spec`: Specifies the PyInstaller spec file, which contains the configuration for building the executable.
   - `--clean`: Instructs PyInstaller to clean the PyInstaller cache and temporary files before building, ensuring a fresh build.

3. **Executable Location:** After the build process is complete, the executable file will be located in the `dist` directory within the project root.

## Architecture

The File Organizer application is structured into UI and utility modules, following a clear separation of concerns:

- **`main.py`**: Serves as the application's entry point, initializing the main Tkinter window and orchestrating the UI and backend managers.
- **`src/ui`**: Houses all user interface components, built using Tkinter. Key components include:
    - `ConfigTab`: Implements the "Configuration" tab, allowing users to manage application settings and presets.
    - `ProcessingTab`: Implements the "Processing" tab, providing the main workflow for PDF processing, queue management, and PDF viewing.
    - `FuzzySearchFrame`: A custom widget for fuzzy-searchable dropdown/combobox inputs, used for filter selection.
    - `PDFViewer`: A component for displaying PDF documents with zoom and rotation capabilities.
    - `QueueDisplay`:  A table-based display for the PDF processing queue, showing task statuses and details.
    - `ErrorDialog`: A reusable dialog for displaying error messages and detailed tracebacks.
- **`src/utils`**: Contains utility modules that handle the core logic of the application:
    - `ConfigManager`: Manages application configuration, including loading, saving, resetting, and handling presets.
    - `ExcelManager`:  Manages interactions with Excel files, including loading data, caching hyperlinks, updating cell values, and adding new rows.
    - `PDFManager`: Handles PDF processing tasks such as generating output paths based on templates, processing PDFs (renaming and moving), and managing PDF display functionalities (rotation, rendering).
    - `TemplateManager`:  Provides the template processing engine, parsing templates and applying data and operations to generate output strings (e.g., for filenames and paths).
     - `models.py`: Defines data models used throughout the application, such as `PDFTask` for representing processing tasks in the queue.

**Design Patterns:**

The application employs several design patterns to enhance its structure and maintainability:

- **Manager Pattern:** Utility modules like `ConfigManager`, `ExcelManager`, `PDFManager`, and `TemplateManager` encapsulate specific functionalities, promoting modularity and separation of concerns.
- **Singleton Pattern:** The `ProcessingTab` class uses a singleton pattern to ensure only one instance of the processing tab exists, which is necessary for managing the processing queue and UI state consistently.
- **Observer Pattern:** The `ConfigManager` utilizes an observer pattern to notify UI components (like `ProcessingTab` and `ConfigTab`) about configuration changes, allowing them to update their displays and states accordingly.
- **Threading:** The `ProcessingQueue` uses threading to manage PDF processing tasks in the background, preventing the UI from freezing during long-running operations and improving responsiveness.

This architecture ensures a robust, modular, and maintainable application, making it easier to extend and adapt to future requirements.