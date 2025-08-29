from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
from werkzeug.utils import secure_filename
import extract_employee_shifts
import shutil
from docx import Document
import re

# ---------------------------------------------------------------------------
# Application factory / configuration
# ---------------------------------------------------------------------------
# Notes for deployment inside containers:
# - SECRET_KEY should be provided via environment variable in production.
# - DEBUG is toggled via FLASK_DEBUG=1 (docker-compose can set it for dev).
# - PORT can be overridden with PORT env (default 5000) for platforms like Heroku / Render.
# ---------------------------------------------------------------------------

app = Flask(__name__)
# Use environment variable (fallback ONLY for local/dev). Replace default before real prod.
app.secret_key = os.environ.get('SECRET_KEY', 'dev-insecure-change-me')

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/test')
def test_page():
    return render_template('test.html')

@app.route('/test_route', methods=['GET', 'POST'])
def test_route():
    if request.method == 'POST':
        return jsonify({'status': 'POST received', 'data': dict(request.form)})
    return jsonify({'status': 'GET received'})

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_files():
    """Analyze uploaded files and return summary data for heatmap"""
    try:
        if 'files' not in request.files:
            return jsonify({'error': 'No files selected'}), 400
        
        files = request.files.getlist('files')
        
        if len(files) == 0:
            return jsonify({'error': 'No files selected'}), 400
        
        # Create a temporary directory for this analysis
        session_dir = tempfile.mkdtemp(dir=UPLOAD_FOLDER)
        
        # Save uploaded files with filename mapping
        uploaded_files = []
        filename_mapping = {}
        for file in files:
            if file and file.filename and allowed_file(file.filename):
                original_filename = file.filename
                secure_name = secure_filename(file.filename)
                file_path = os.path.join(session_dir, secure_name)
                file.save(file_path)
                
                filename_mapping[secure_name] = original_filename
                uploaded_files.append(secure_name)
        
        if not uploaded_files:
            shutil.rmtree(session_dir)
            return jsonify({'error': 'No valid .docx files uploaded'}), 400
        
        # Save filename mapping for later use
        mapping_file = os.path.join(session_dir, 'filename_mapping.txt')
        with open(mapping_file, 'w', encoding='utf-8') as f:
            for secure_name, original_name in filename_mapping.items():
                f.write(f'{secure_name}:{original_name}\n')
        
        # Analyze all employees and shifts
        summary_data = analyze_all_employees(session_dir, filename_mapping)
        
        return jsonify({
            'success': True,
            'session_dir': os.path.basename(session_dir),
            'summary': summary_data
        })
        
    except Exception as e:
        return jsonify({'error': f'Analysis failed: {str(e)}'}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    """Generate individual employee report"""
    try:
        employee_name = request.form.get('employee_name')
        session_id = request.form.get('session_id')
        
        if not employee_name or not session_id:
            return jsonify({'error': 'Missing employee name or session ID'}), 400
        
        session_dir = os.path.join(UPLOAD_FOLDER, session_id)
        
        if not os.path.exists(session_dir):
            return jsonify({'error': 'Session not found'}), 404
        
        # Load filename mapping
        mapping_file = os.path.join(session_dir, 'filename_mapping.txt')
        filename_mapping = {}
        
        if os.path.exists(mapping_file):
            with open(mapping_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if ':' in line:
                        temp_name, original_name = line.strip().split(':', 1)
                        filename_mapping[temp_name] = original_name
        else:
            # If no mapping file, create one from current files
            for secure_fname in os.listdir(session_dir):
                if secure_fname.endswith('.docx'):
                    filename_mapping[secure_fname] = secure_fname
        
        # Extract shifts for the specific employee
        shifts = extract_with_mapping(employee_name, session_dir, filename_mapping)
        
        if not shifts:
            return jsonify({'error': f'No shifts found for employee: {employee_name}'}), 404
        
        # Generate Excel file
        output_file = os.path.join(session_dir, f'{employee_name}_shifts.xlsx')
        extract_employee_shifts.write_to_xlsx(shifts, output_file)
        
        return jsonify({
            'success': True,
            'message': f'Found {len(shifts)} shifts for {employee_name}',
            'download_url': f'/download/{session_id}/{employee_name}_shifts.xlsx',
            'session_dir': session_id
        })
        
    except Exception as e:
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

def analyze_all_employees(session_dir, filename_mapping):
    """Analyze all employees and return summary data for heatmap"""
    employee_shifts = {}
    all_employee_names = set()
    
    # First pass: Extract all possible employee names to build a comprehensive list
    for secure_fname in os.listdir(session_dir):
        if not secure_fname.endswith('.docx'):
            continue
            
        filepath = os.path.join(session_dir, secure_fname)
        original_filename = filename_mapping.get(secure_fname, secure_fname)
        
        try:
            doc = Document(filepath)
            
            # Find the table and extract all possible names
            for table in doc.tables:
                for row in table.rows[1:]:  # Skip header
                    shift_type = row.cells[0].text.strip()
                    
                    # Skip "Assenti" entries
                    if 'assenti' in shift_type.lower():
                        continue
                    
                    for cell in row.cells[1:]:
                        cell_text = cell.text.strip()
                        if cell_text and not cell_text.isspace():
                            # Extract all possible names (both comma-separated and from parentheses)
                            # First, get names separated by commas/newlines
                            base_names = [name.strip() for name in cell_text.replace('\n', ',').split(',') if name.strip()]
                            
                            for base_name in base_names:
                                # Remove non-name parentheses content (like shifts, numbers, etc.)
                                # but preserve potential employee names
                                clean_base = re.sub(r'\([^)]*(?:turno|shift|ore|h|:|\d+)[^)]*\)', '', base_name, flags=re.IGNORECASE)
                                
                                # Extract main name (before parentheses)
                                main_name = re.sub(r'\([^)]*\)', '', clean_base).strip()
                                if main_name and len(main_name) > 1 and not is_roman_numeral(main_name):
                                    all_employee_names.add(main_name)
                                
                                # Extract potential names from parentheses
                                parentheses_matches = re.findall(r'\(([^)]+)\)', clean_base)
                                for match in parentheses_matches:
                                    potential_name = match.strip()
                                    # Check if it looks like a name (not a number, time, or shift info)
                                    if (potential_name and 
                                        len(potential_name) > 1 and 
                                        not re.match(r'^[\d:.-]+$', potential_name) and
                                        not re.search(r'\b(?:turno|shift|ore|h)\b', potential_name, re.IGNORECASE)):
                                        all_employee_names.add(potential_name)
        
        except Exception as e:
            print(f"Error in first pass processing {original_filename}: {e}")
            continue
    
    print(f"Found {len(all_employee_names)} unique employee names: {sorted(all_employee_names)}")
    
    # Second pass: Extract shifts using the comprehensive name list
    for secure_fname in os.listdir(session_dir):
        if not secure_fname.endswith('.docx'):
            continue
            
        filepath = os.path.join(session_dir, secure_fname)
        original_filename = filename_mapping.get(secure_fname, secure_fname)
        
        try:
            doc = Document(filepath)
            
            # Extract date range from ORIGINAL filename
            start_day, start_month, end_day, end_month, start_year, end_year = extract_employee_shifts.extract_date_range_from_filename(original_filename)
            if not all([start_day, start_month, end_day, end_month]):
                continue
            
            # Find the table and extract all employees
            for table in doc.tables:
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                days = headers[1:]
                
                for row in table.rows[1:]:
                    shift_type = row.cells[0].text.strip()
                    
                    # Skip "Assenti" entries
                    if 'assenti' in shift_type.lower():
                        continue
                    
                    for i, cell in enumerate(row.cells[1:]):
                        cell_text = cell.text.strip()
                        if cell_text and not cell_text.isspace():
                            # Extract employee names using the comprehensive approach
                            employee_names = extract_employee_names_from_cell(cell_text, all_employee_names)
                            
                            for employee_name in employee_names:
                                if employee_name:
                                    if employee_name not in employee_shifts:
                                        employee_shifts[employee_name] = {}
                                    
                                    if shift_type not in employee_shifts[employee_name]:
                                        employee_shifts[employee_name][shift_type] = 0
                                    
                                    employee_shifts[employee_name][shift_type] += 1
                                    
                                    # Add weekend shifts for "Guardia" on Friday
                                    day_name = days[i] if i < len(days) else ''
                                    if 'guardia' in shift_type.lower() and day_name.lower() == 'venerdì':
                                        employee_shifts[employee_name][shift_type] += 2  # Saturday + Sunday
        
        except Exception as e:
            print(f"Error processing {original_filename}: {e}")
            continue
    
    return employee_shifts

def is_roman_numeral(text):
    """Check if text is a Roman numeral in parentheses that should be ignored"""
    if not text:
        return False
    
    # Remove parentheses if present
    clean_text = text.strip('()')
    
    # Check if it's a Roman numeral (I, II, III, IV, V, VI, VII, VIII, IX, X, etc.)
    roman_pattern = r'^(I{1,3}|IV|V|VI{0,3}|IX|X{1,3}|XL|L|LX{0,3}|XC|C{1,3}|CD|D|DC{0,3}|CM|M{1,3})$'
    return bool(re.match(roman_pattern, clean_text.upper()))

def extract_employee_names_from_cell(cell_text, known_names):
    """Extract employee names from a cell, considering both comma-separated and parentheses formats"""
    employee_names = []
    
    # Split by commas and newlines first
    base_parts = [part.strip() for part in cell_text.replace('\n', ',').split(',') if part.strip()]
    
    for part in base_parts:
        # Remove non-name parentheses content (shifts, times, etc.)
        cleaned_part = re.sub(r'\([^)]*(?:turno|shift|ore|h|:|\d+)[^)]*\)', '', part, flags=re.IGNORECASE)
        
        # Extract main name (before any parentheses)
        main_name = re.sub(r'\([^)]*\)', '', cleaned_part).strip()
        if main_name and main_name in known_names:
            employee_names.append(main_name)
        
        # Check parentheses for additional employee names
        parentheses_matches = re.findall(r'\(([^)]+)\)', cleaned_part)
        for match in parentheses_matches:
            potential_name = match.strip()
            
            # Skip Roman numerals
            if is_roman_numeral(potential_name):
                continue
                
            if potential_name in known_names:
                employee_names.append(potential_name)
            else:
                # Check if it's a partial match (like "Di Bella" matching "Di Bella")
                for known_name in known_names:
                    if (potential_name.lower() in known_name.lower() or 
                        known_name.lower() in potential_name.lower()) and len(potential_name) > 2:
                        employee_names.append(known_name)
                        break
    
    return list(set(employee_names))  # Remove duplicates

def extract_with_mapping(employee_name, session_dir, filename_mapping):
    """Extract shifts for specific employee using filename mapping"""
    results = []
    
    # First pass: Extract all employee names from all files
    all_employee_names = set()
    for secure_fname in os.listdir(session_dir):
        if not secure_fname.endswith('.docx'):
            continue
            
        filepath = os.path.join(session_dir, secure_fname)
        
        try:
            doc = Document(filepath)
            
            # Find the table and extract all possible names
            for table in doc.tables:
                for row in table.rows[1:]:  # Skip header
                    shift_type = row.cells[0].text.strip()
                    
                    # Skip "Assenti" entries
                    if 'assenti' in shift_type.lower():
                        continue
                    
                    for cell in row.cells[1:]:
                        cell_text = cell.text.strip()
                        if cell_text and not cell_text.isspace():
                            # Simple extraction for building the name list
                            base_names = [name.strip() for name in cell_text.replace('\n', ',').split(',') if name.strip()]
                            
                            for base_name in base_names:
                                # Remove non-name parentheses content
                                clean_base = re.sub(r'\([^)]*(?:turno|shift|ore|h|:|\d+)[^)]*\)', '', base_name, flags=re.IGNORECASE)
                                
                                # Extract main name
                                main_name = re.sub(r'\([^)]*\)', '', clean_base).strip()
                                if main_name and len(main_name) > 1 and not is_roman_numeral(main_name):
                                    all_employee_names.add(main_name)
                                
                                # Extract potential names from parentheses
                                parentheses_matches = re.findall(r'\(([^)]+)\)', clean_base)
                                for match in parentheses_matches:
                                    potential_name = match.strip()
                                    
                                    # Skip Roman numerals
                                    if is_roman_numeral(potential_name):
                                        continue
                                        
                                    if (potential_name and 
                                        len(potential_name) > 1 and 
                                        not re.match(r'^[\d:.-]+$', potential_name) and
                                        not re.search(r'\b(?:turno|shift|ore|h)\b', potential_name, re.IGNORECASE)):
                                        all_employee_names.add(potential_name)
        
        except Exception as e:
            print(f"Error in name extraction for {secure_fname}: {e}")
            continue
    
    # Second pass: Extract shifts for the specific employee
    for secure_fname in os.listdir(session_dir):
        if not secure_fname.endswith('.docx'):
            continue
            
        filepath = os.path.join(session_dir, secure_fname)
        original_filename = filename_mapping.get(secure_fname, secure_fname)
        
        try:
            doc = Document(filepath)
            
            # Extract date range from ORIGINAL filename
            start_day, start_month, end_day, end_month, start_year, end_year = extract_employee_shifts.extract_date_range_from_filename(original_filename)
            if not all([start_day, start_month, end_day, end_month]):
                continue
                
            # Get actual dates for this week
            week_dates = extract_employee_shifts.get_week_dates_from_range(start_day, start_month, end_day, end_month, start_year, end_year)
            
            # Find the table
            for table in doc.tables:
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                days = headers[1:]
                
                for row in table.rows[1:]:
                    shift_type = row.cells[0].text.strip()
                    
                    for i, cell in enumerate(row.cells[1:]):
                        cell_text = cell.text.strip()
                        if cell_text and not cell_text.isspace():
                            # Extract all possible names using the enhanced method
                            employee_names = extract_employee_names_from_cell(cell_text, all_employee_names)
                            
                            # Check if our target employee is in this cell
                            if employee_name in employee_names:
                                if 'assenti' in shift_type.lower():
                                    continue
                                    
                                day_name = days[i] if i < len(days) else f'Day{i+1}'
                                date_str = week_dates[i] if i < len(week_dates) else ''
                                
                                results.append({
                                    'File': original_filename,
                                    'Data': date_str,
                                    'Giorno': day_name,
                                    'Turno': shift_type,
                                    'Dipendente': employee_name
                                })
                                
                                # Add weekend shifts for "Guardia" on Friday
                                if 'guardia' in shift_type.lower() and day_name.lower() == 'venerdì':
                                    # Add Saturday
                                    saturday_date = extract_employee_shifts.add_days_to_date(date_str, 1) if date_str else ''
                                    results.append({
                                        'File': original_filename,
                                        'Data': saturday_date,
                                        'Giorno': 'Sabato',
                                        'Turno': shift_type,
                                        'Dipendente': employee_name
                                    })
                                    
                                    # Add Sunday
                                    sunday_date = extract_employee_shifts.add_days_to_date(date_str, 2) if date_str else ''
                                    results.append({
                                        'File': original_filename,
                                        'Data': sunday_date,
                                        'Giorno': 'Domenica',
                                        'Turno': shift_type,
                                        'Dipendente': employee_name
                                    })
        
        except Exception as e:
            print(f"Error processing {original_filename}: {e}")
            continue
    
    return results

@app.route('/download/<session_id>/<filename>')
def download_file(session_id, filename):
    try:
        session_dir = os.path.join(UPLOAD_FOLDER, session_id)
        file_path = os.path.join(session_dir, filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(file_path, as_attachment=True, download_name=filename)
    
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/cleanup/<session_id>', methods=['POST'])
def cleanup_session(session_id):
    try:
        session_dir = os.path.join(UPLOAD_FOLDER, session_id)
        if os.path.exists(session_dir):
            shutil.rmtree(session_dir)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': f'Cleanup failed: {str(e)}'}), 500

if __name__ == '__main__':
    # Local/dev execution path (container uses gunicorn via wsgi:app)
    debug = os.environ.get('FLASK_DEBUG', '0') == '1'
    port = int(os.environ.get('PORT', '5000'))
    host = os.environ.get('HOST', '0.0.0.0')
    app.run(debug=debug, host=host, port=port)
