# Analizzatore Turni

This tool extracts and analyzes work shifts from Italian Word documents (.docx) and creates comprehensive Excel reports.

## Containerized Deployment

The application can run inside Docker (production-ready using Gunicorn) with an attached volume for persistent uploads.

### Quick Start (Docker)

```bash
# Build image
docker build -t shifts-turni .

# Run container (port 5000 exposed)
docker run -p 5000:5000 -e SECRET_KEY="replace-me" shifts-turni
```

### Using docker-compose

```bash
docker compose up --build
```

Then open: <http://localhost:5000>

Uploads persist in the local `uploads/` directory mounted as a volume.

### Development Mode (hot reload)

Uncomment the `command` and related `environment` lines in `docker-compose.yml` to use Flask’s reloader (not for production).

### Environment Variables

- `SECRET_KEY` – Required for session security (set a strong random value in prod)
- `PORT` – Listening port inside container (default 5000)
- `FLASK_DEBUG` – Set to `1` for debug mode (only for local development)

---

## Run Locally (mamba / conda + python)

If you prefer not to use Docker, you can run the app directly in a Python environment.

### 1. Create & activate environment (mamba recommended)

```bash
mamba create -n shifts-turni python=3.13 -y
mamba activate shifts-turni
```

Or with conda:

```bash
conda create -n shifts-turni python=3.13 -y
conda activate shifts-turni
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set environment variables (at minimum SECRET_KEY)

macOS / Linux (zsh/bash):
```bash
export SECRET_KEY="replace-with-strong-value"
export FLASK_DEBUG=1   # optional for auto-reload
export PORT=5000       # optional override
```

Windows PowerShell:
```powershell
$Env:SECRET_KEY = "replace-with-strong-value"
$Env:FLASK_DEBUG = "1"
```

### 4. Run the application

Direct (uses `app.py`):
```bash
python app.py
```

Or with Flask CLI (enable debug reloader):
```bash
FLASK_DEBUG=1 flask run --host=0.0.0.0 --port=5000
```

Then open <http://localhost:5000>

### 5. (Optional) Run extraction script directly

```bash
python extract_employee_shifts.py "Employee Name"
```

---

## Project Structure

```text
analizzatore-turni/
├── extract_employee_shifts.py  # Main script to extract shifts
├── app.py                      # Flask web application
├── wsgi.py                     # Gunicorn entrypoint
├── uploads/                    # Uploaded session data (mounted volume)
├── templates/                  # HTML templates
├── Dockerfile                  # Container build
├── docker-compose.yml          # Multi-service / dev orchestration
└── tests/                      # Testing and utility scripts
```

## Main Script: `extract_employee_shifts.py`

### Features

- **Date parsing**: Converts filenames like "55. 11:11 - 15:11.docx" to actual dates (Nov 11-15, 2024)
- **File filtering**: Automatically excludes temporary files (`~$...`) and underscore patterns (`XX_....docx`)
- **Weekend extension**: Adds Saturday/Sunday when "Guardia" is assigned on Friday
- **Multi-sheet output**: Creates 3 different views of the data
- **Command-line employee input**: Specify employee name as argument

### Usage

```bash
python extract_employee_shifts.py "employee_name"
```

### Output Excel File Structure

1. **"Tutti i Turni"** - Complete list of all shifts with dates
2. **"Riepilogo per Turno"** - Summary count by shift type  
3. **"Date per Turno"** - Horizontal view with dates grouped by shift type

## Testing Scripts

Located in the `tests/` folder:

- `view_all_sheets.py` - View all Excel sheets in terminal
- `debug_filtering.py` - See which files are processed vs filtered out
- `view_results.py` - Convert Excel to CSV for easy viewing
- `create_csv_output.py` - Alternative CSV conversion using pandas

### Run tests from main directory

```bash
python tests/view_all_sheets.py
python tests/debug_filtering.py
```

## File Naming Convention

The script expects .docx files with this naming pattern:

- `XX. DD:MM - DD:MM.docx` (e.g., `55. 11:11 - 15:11.docx`)
- Where `XX` is a sequence number and `DD:MM` represents day:month

### Filtered Files

- ❌ Temporary Word files: `~$...`
- ❌ Underscore patterns: `XX_...`

## Dependencies

- `python-docx` - Read Word documents
- `openpyxl` - Create Excel files
- `flask` - Web framework
- `gunicorn` - Production WSGI server

## Year Logic

- **November/December dates** → 2024
- **January onwards** → 2025
