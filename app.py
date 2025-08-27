from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
from werkzeug.utils import secure_filename
import extract_employee_shifts
import shutil
from docx import Document
import re

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this in production

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
        
        # Analyze all employees and shifts
        summary_data = analyze_all_employees(session_dir, filename_mapping)
        
        return jsonify({
            'success': True,
            'session_dir': os.path.basename(session_dir),
            'summary': summary_data
        })
        
    except Exception as e:
        return jsonify({'error': f'Analysis failed: {str(e)}'}), 500

def analyze_all_employees(session_dir, filename_mapping):
    """Analyze all employees and return summary data for heatmap"""
    employee_shifts = {}
    
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
                            # Extract employee names from cell (could be multiple, separated by commas or newlines)
                            employee_names = [name.strip() for name in cell_text.replace('\n', ',').split(',') if name.strip()]
                            
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

def extract_with_mapping(employee_name, session_dir, filename_mapping):
    """Extract shifts for specific employee using filename mapping"""
    results = []
    
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
                        cell_text = cell.text.strip().lower()
                        if employee_name.lower() in cell_text:
                            if 'assenti' in shift_type.lower():
                                continue
                                
                            day_name = days[i] if i < len(days) else f'Day{i+1}'
                            date_str = week_dates[i] if i < len(week_dates) else ''
                            
                            results.append({
                                'File': original_filename,
                                'Data': date_str,
                                'Giorno': day_name,
                                'Turno': shift_type
                            })
                            
                            # Weekend guardia logic
                            if 'guardia' in shift_type.lower() and day_name.lower() == 'venerdì':
                                if date_str:
                                    from datetime import datetime, timedelta
                                    saturday_date = datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=1)
                                    results.append({
                                        'File': original_filename,
                                        'Data': saturday_date.strftime('%Y-%m-%d'),
                                        'Giorno': 'Sabato',
                                        'Turno': shift_type
                                    })
                                    
                                    sunday_date = datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=2)
                                    results.append({
                                        'File': original_filename,
                                        'Data': sunday_date.strftime('%Y-%m-%d'),
                                        'Giorno': 'Domenica',
                                        'Turno': shift_type
                                    })
        
        except Exception as e:
            print(f"Error processing {original_filename}: {e}")
            continue
    
    return results

@app.route('/upload', methods=['POST'])
def upload_files():
    """Generate report for specific employee from previously analyzed data"""
    try:
        employee_name = request.form.get('employee_name', '').strip()
        session_id = request.form.get('session_id', '').strip()
        
        if not employee_name:
            return jsonify({'error': 'Employee name is required'}), 400
        
        if not session_id:
            return jsonify({'error': 'Session ID is required'}), 400
        
        session_dir = os.path.join(UPLOAD_FOLDER, session_id)
        
        if not os.path.exists(session_dir):
            return jsonify({'error': 'Session expired or invalid'}), 400
        
        # Reconstruct filename mapping
        filename_mapping = {}
        for fname in os.listdir(session_dir):
            if fname.endswith('.docx'):
                # Try to reconstruct original filename (this is a simplified approach)
                # In a production system, you'd want to store the mapping
                original = fname.replace('_', ' ').replace('.docx', '.docx')
                # Add colons back for time format
                original = re.sub(r'(\d+) (\d+) - (\d+) (\d+)', r'\1:\2 - \3:\4', original)
                filename_mapping[fname] = original
        
        # Extract shifts using the existing logic
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
    app.run(debug=True, port=5001)
