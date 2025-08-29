# Analizzatore Turni

This tool extracts and analyzes work shifts from Italian Word documents (.docx) and creates comprehensive Excel reports.

## Desktop (macOS) App

A macOS standalone app can be built so non-technical users can just double‑click and use the interface without opening a browser.

### Run as Desktop App (dev)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python webview_app.py
```text

This launches the Flask backend and opens a native window via `pywebview`.

### Build a .app Bundle (py2app)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install py2app
python setup.py py2app -A   # alias mode (fast, for testing)
python setup.py py2app      # create dist/Analizzatore Turni.app
```

Distribute the produced `dist/Analizzatore Turni.app` to users. On first run macOS Gatekeeper may warn; right‑click → Open to approve.

Uploaded files & generated Excel outputs are stored in `~/ShiftsAnalyzerUploads` when packaged (avoids read‑only bundle paths). You can safely delete that folder to reset the app cache.

If you add or change templates, rebuild the app so the new HTML files are included.

---

## Project Structure

```
analizzatore-turni/
├── extract_employee_shifts.py  # Main script to extract shifts
├── employee_shifts.xlsx        # Generated Excel report with 3 sheets
├── turni/                      # Folder containing .docx shift documents
└── tests/                      # Testing and utility scripts
    ├── view_all_sheets.py      # Display all Excel sheets content
    ├── view_results.py         # Convert Excel to CSV and display
    ├── debug_filtering.py      # Show which files are processed vs filtered
    └── create_csv_output.py    # Convert Excel to CSV using pandas
```

## Main Script: `extract_employee_shifts.py`

### Features:
- **Date parsing**: Converts filenames like "55. 11:11 - 15:11.docx" to actual dates (Nov 11-15, 2024)
- **File filtering**: Automatically excludes temporary files (`~$...`) and underscore patterns (`XX_....docx`)
- **Weekend extension**: Adds Saturday/Sunday when "Guardia" is assigned on Friday
- **Multi-sheet output**: Creates 3 different views of the data
- **Command-line employee input**: Specify employee name as argument

### Usage:
```bash
# Activate the conda environment
mamba activate turni

# Run the main script with employee name
python extract_employee_shifts.py "employee_name"

# Example:
python extract_employee_shifts.py "John Doe"
```

### Output Excel File Structure:

1. **"Tutti i Turni"** - Complete list of all shifts with dates
2. **"Riepilogo per Turno"** - Summary count by shift type  
3. **"Date per Turno"** - Horizontal view with dates grouped by shift type
```

## Testing Scripts

Located in the `tests/` folder:

- **`view_all_sheets.py`** - View all Excel sheets in terminal
- **`debug_filtering.py`** - See which files are processed vs filtered out
- **`view_results.py`** - Convert Excel to CSV for easy viewing
- **`create_csv_output.py`** - Alternative CSV conversion using pandas

### Run tests from main directory:
```bash
python tests/view_all_sheets.py
python tests/debug_filtering.py
```

## File Naming Convention

The script expects .docx files with this naming pattern:
- `XX. DD:MM - DD:MM.docx` (e.g., `55. 11:11 - 15:11.docx`)
- Where `XX` is a sequence number and `DD:MM` represents day:month

### Filtered Files:
- ❌ Temporary Word files: `~$...`
- ❌ Underscore patterns: `XX_...`

## Dependencies

- `python-docx` - Read Word documents
- `openpyxl` - Create Excel files
- `datetime` - Date manipulation
- `re` - Regular expressions for filename parsing

## Year Logic

- **November/December dates** → 2024
- **January onwards** → 2025
