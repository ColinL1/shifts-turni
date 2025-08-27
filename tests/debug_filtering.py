import os
import re

TURNI_FOLDER = '../turni'

def show_file_filtering():
    print("Files in turni folder:")
    print("=" * 50)
    
    all_files = [f for f in os.listdir(TURNI_FOLDER) if f.endswith('.docx')]
    processed_files = []
    filtered_files = []
    
    for f in all_files:
        if f.startswith('~$'):
            filtered_files.append(f + " (temporary file)")
        elif re.match(r'^\d+_\.', f):
            filtered_files.append(f + " (underscore pattern)")
        else:
            processed_files.append(f)
    
    print(f"Files that WILL BE PROCESSED ({len(processed_files)}):")
    for f in sorted(processed_files):
        print(f"  ✓ {f}")
    
    print(f"\nFiles that will be FILTERED OUT ({len(filtered_files)}):")
    for f in sorted(filtered_files):
        print(f"  ✗ {f}")

if __name__ == '__main__':
    show_file_filtering()
