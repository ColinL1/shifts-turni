from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
from werkzeug.utils import secure_filename
import extract_employee_shifts
import shutil
from docx import Document

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

@app.route('/upload', methods=['POST'])
def upload_files():
    print("Upload route called")  # Debug
    try:
        employee_name = request.form.get('employee_name', '').strip()
        print(f"Employee name: {employee_name}")  # Debug
        
        if not employee_name:
            return jsonify({'error': 'Employee name is required'}), 400
        
        if 'files' not in request.files:
            print("No files in request")  # Debug
            return jsonify({'error': 'No files selected'}), 400
        
        files = request.files.getlist('files')
        print(f"Number of files: {len(files)}")  # Debug
        
        if len(files) == 0:
            return jsonify({'error': 'No files selected'}), 400
        
        # Create a temporary directory for this session
        session_dir = tempfile.mkdtemp(dir=UPLOAD_FOLDER)
        
        # Save uploaded files with original names preserved for processing
        uploaded_files = []
        filename_mapping = {}  # Maps secure filename to original filename
        for file in files:
            if file and file.filename and allowed_file(file.filename):
                original_filename = file.filename
                secure_name = secure_filename(file.filename)
                file_path = os.path.join(session_dir, secure_name)
                file.save(file_path)
                
                # Store mapping for later processing
                filename_mapping[secure_name] = original_filename
                uploaded_files.append(secure_name)
                print(f"Saved file: {secure_name} (original: {original_filename})")  # Debug
        
        print(f"Total uploaded files: {len(uploaded_files)}")  # Debug
        
        if not uploaded_files:
            shutil.rmtree(session_dir)
            return jsonify({'error': 'No valid .docx files uploaded'}), 400
        
        # Process the files with filename mapping
        try:
            # Create a custom extraction function that uses original filenames
            def extract_with_mapping(employee_name, session_dir, filename_mapping):
                results = []
                
                for secure_filename in os.listdir(session_dir):
                    if not secure_filename.endswith('.docx'):
                        continue
                        
                    filepath = os.path.join(session_dir, secure_filename)
                    original_filename = filename_mapping.get(secure_filename, secure_filename)
                    
                    print(f"Processing {secure_filename} as {original_filename}")  # Debug
                    
                    doc = Document(filepath)
                    
                    # Extract date range from ORIGINAL filename
                    start_day, start_month, end_day, end_month = extract_employee_shifts.extract_date_range_from_filename(original_filename)
                    if not all([start_day, start_month, end_day, end_month]):
                        print(f"Could not parse date from {original_filename}")  # Debug
                        continue
                        
                    # Get actual dates for this week
                    week_dates = extract_employee_shifts.get_week_dates_from_range(start_day, start_month, end_day, end_month)
                    
                    # Find the table (same logic as original)
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
                
                return results
            
            # Extract shifts using the custom function
            shifts = extract_with_mapping(employee_name, session_dir, filename_mapping)
            print(f"Found {len(shifts)} shifts")  # Debug
            
            if not shifts:
                shutil.rmtree(session_dir)
                return jsonify({'error': f'No shifts found for employee: {employee_name}'}), 404
            
            # Generate Excel file
            output_file = os.path.join(session_dir, f'{employee_name}_shifts.xlsx')
            extract_employee_shifts.write_to_xlsx(shifts, output_file)
            
            # Return success with file info
            return jsonify({
                'success': True,
                'message': f'Found {len(shifts)} shifts for {employee_name}',
                'download_url': f'/download/{os.path.basename(session_dir)}/{employee_name}_shifts.xlsx',
                'session_dir': os.path.basename(session_dir)
            })
            
        except Exception as e:
            shutil.rmtree(session_dir)
            return jsonify({'error': f'Error processing files: {str(e)}'}), 500
    
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
