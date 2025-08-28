import os
import re
import sys
from docx import Document
from openpyxl import Workbook
from datetime import datetime, timedelta

# Folder containing the .docx files
TURNI_FOLDER = 'turni'
OUTPUT_XLSX = 'employee_shifts.xlsx'

# Helper to extract date range from filename (e.g., '57. 25/11/24 - 29/11/24.docx' or '59. 09:12:24 - 13:12:24.docx' or '55. 11:11 - 15:11.docx')
def extract_date_range_from_filename(filename):
    # First try the new format with forward slashes: "25/11/24 - 29/11/24"
    match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})\s*-\s*(\d{1,2})/(\d{1,2})/(\d{2,4})', filename)
    if match:
        start_day, start_month, start_year, end_day, end_month, end_year = match.groups()
        # Convert 2-digit years to 4-digit (24 -> 2024)
        start_year = int(start_year)
        end_year = int(end_year)
        if start_year < 100:
            start_year += 2000
        if end_year < 100:
            end_year += 2000
        return int(start_day), int(start_month), int(end_day), int(end_month), start_year, end_year
    
    # Try format with colons and year: "09:12:24 - 13:12:24" (day:month:year)
    match = re.search(r'(\d{1,2}):(\d{1,2}):(\d{2,4})\s*-\s*(\d{1,2}):(\d{1,2}):(\d{2,4})', filename)
    if match:
        start_day, start_month, start_year, end_day, end_month, end_year = match.groups()
        # Convert 2-digit years to 4-digit (24 -> 2024, 25 -> 2025)
        start_year = int(start_year)
        end_year = int(end_year)
        if start_year < 100:
            start_year += 2000
        if end_year < 100:
            end_year += 2000
        return int(start_day), int(start_month), int(end_day), int(end_month), start_year, end_year
    
    # Fall back to old format: "11:11 - 15:11" (day:month without year)
    match = re.search(r'(\d{1,2}):(\d{1,2})\s*-\s*(\d{1,2}):(\d{1,2})', filename)
    if match:
        start_day, start_month, end_day, end_month = match.groups()
        return int(start_day), int(start_month), int(end_day), int(end_month), None, None
    
    return None, None, None, None, None, None

# Helper to get the year based on the month (November 2024 to 2025)
def get_year_for_month(month):
    # Assuming the schedule starts in November 2024 and continues into 2025
    # November and December are 2024, January onwards are 2025
    if month >= 11:  # November, December
        return 2024
    else:  # January onwards
        return 2025

# Helper to get week dates from the date range
def get_week_dates_from_range(start_day, start_month, end_day, end_month, start_year=None, end_year=None):
    # Use provided years if available, otherwise use the old logic
    if start_year is None:
        start_year = get_year_for_month(start_month)
    if end_year is None:
        end_year = get_year_for_month(end_month)
    
    start_date = datetime(start_year, start_month, start_day)
    end_date = datetime(end_year, end_month, end_day)
    
    # Generate dates for the work week (Monday to Friday)
    week_dates = []
    current_date = start_date
    
    while current_date <= end_date:
        # Only include weekdays (Monday=0 to Friday=4)
        if current_date.weekday() < 5:
            week_dates.append(current_date.strftime('%Y-%m-%d'))
        current_date += timedelta(days=1)
    
    return week_dates

def add_days_to_date(date_str, days):
    """Add days to a date string and return the new date string"""
    if not date_str:
        return ''
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        new_date = date_obj + timedelta(days=days)
        return new_date.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        return ''

# Helper to get all docx files in the folder
def get_docx_files(folder):
    files = []
    for f in os.listdir(folder):
        if f.endswith('.docx') and not f.startswith('~$'):  # Skip temporary Word files
            # Skip files that start with pattern like "90_." (number + underscore + dot)
            if re.match(r'^\d+_\.', f):
                continue
            files.append(os.path.join(folder, f))
    return files

# Main extraction logic
def extract_employee_shifts(employee_name):
    results = []
    
    for filepath in get_docx_files(TURNI_FOLDER):
        doc = Document(filepath)
        filename = os.path.basename(filepath)
        
        # Extract date range from filename
        start_day, start_month, end_day, end_month, start_year, end_year = extract_date_range_from_filename(filename)
        if not all([start_day, start_month, end_day, end_month]):
            continue
            
        # Get actual dates for this week
        week_dates = get_week_dates_from_range(start_day, start_month, end_day, end_month, start_year, end_year)
        
        # Find the table
        for table in doc.tables:
            # Assume first column is shift type, first row is days
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            days = headers[1:]  # skip first column (type)
            
            for row in table.rows[1:]:
                shift_type = row.cells[0].text.strip()
                
                for i, cell in enumerate(row.cells[1:]):
                    cell_text = cell.text.strip().lower()
                    if employee_name.lower() in cell_text:
                        # Skip "Assenti" entries
                        if 'assenti' in shift_type.lower():
                            continue
                            
                        # Get the day name and corresponding date
                        day_name = days[i] if i < len(days) else f'Day{i+1}'
                        date_str = week_dates[i] if i < len(week_dates) else ''
                        
                        results.append({
                            'File': filename,
                            'Data': date_str,
                            'Giorno': day_name,
                            'Turno': shift_type
                        })
                        
                        # If it's "Guardia" on Friday, add Saturday and Sunday too
                        if 'guardia' in shift_type.lower() and day_name.lower() == 'venerdì':
                            # Add Saturday
                            if date_str:
                                saturday_date = datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=1)
                                results.append({
                                    'File': filename,
                                    'Data': saturday_date.strftime('%Y-%m-%d'),
                                    'Giorno': 'Sabato',
                                    'Turno': shift_type
                                })
                                
                                # Add Sunday
                                sunday_date = datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=2)
                                results.append({
                                    'File': filename,
                                    'Data': sunday_date.strftime('%Y-%m-%d'),
                                    'Giorno': 'Domenica',
                                    'Turno': shift_type
                                })
    
    return results

def write_to_xlsx(data, output_path):
    from collections import Counter
    
    # Sort data by date (oldest to newest)
    sorted_data = sorted(data, key=lambda x: x['Data'])
    
    wb = Workbook()
    
    # Sheet 1: All shifts (sorted by date)
    ws1 = wb.active
    ws1.title = 'Tutti i Turni'
    headers = ['File', 'Data', 'Giorno', 'Turno']
    ws1.append(headers)
    for row in sorted_data:
        ws1.append([row[h] for h in headers])
    
    # Sheet 2: Summary count by shift type
    ws2 = wb.create_sheet(title='Riepilogo per Turno')
    shift_counts = Counter(row['Turno'] for row in sorted_data)
    ws2.append(['Tipo di Turno', 'Numero di Volte'])
    for shift_type, count in sorted(shift_counts.items()):
        ws2.append([shift_type, count])
    
    # Sheet 3: Dates grouped by shift type (horizontal layout)
    ws3 = wb.create_sheet(title='Date per Turno')
    
    # Group data by shift type
    shifts_by_type = {}
    for row in sorted_data:
        shift_type = row['Turno']
        if shift_type not in shifts_by_type:
            shifts_by_type[shift_type] = []
        shifts_by_type[shift_type].append(row['Data'])
    
    # Sort shift types and get unique sorted dates for each
    sorted_shift_types = sorted(shifts_by_type.keys())
    for shift_type in sorted_shift_types:
        shifts_by_type[shift_type] = sorted(list(set(shifts_by_type[shift_type])))
    
    # Write headers (shift types)
    for col, shift_type in enumerate(sorted_shift_types, 1):
        ws3.cell(row=1, column=col, value=shift_type)
    
    # Write dates under each shift type column
    for col, shift_type in enumerate(sorted_shift_types, 1):
        dates = shifts_by_type[shift_type]
        for row, date in enumerate(dates, 2):  # Start from row 2
            ws3.cell(row=row, column=col, value=date)
    
    wb.save(output_path)

if __name__ == '__main__':
    # Check if employee name is provided as command line argument
    if len(sys.argv) != 2:
        print("Usage: python extract_employee_shifts.py <employee_name>")
        print("Example: python extract_employee_shifts.py 'John Doe'")
        sys.exit(1)
    
    employee_name = sys.argv[1]
    print(f"Extracting shifts for: {employee_name}")
    
    shifts = extract_employee_shifts(employee_name)
    
    if not shifts:
        print(f"No shifts found for employee: {employee_name}")
        print("Please check the employee name spelling and try again.")
        sys.exit(1)
    
    write_to_xlsx(shifts, OUTPUT_XLSX)
    print(f'Saved {len(shifts)} shifts for {employee_name} to {OUTPUT_XLSX}')
