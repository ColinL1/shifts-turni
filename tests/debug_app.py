from flask import Flask
import extract_employee_shifts

app = Flask(__name__)
app.secret_key = 'test-key'

# Import all routes from the main app
from app import *

if __name__ == '__main__':
    print("Registered routes:")
    for rule in app.url_map.iter_rules():
        print(f"  {rule.rule} -> {rule.endpoint} [{', '.join(rule.methods)}]")
    
    print("\nTesting extract_employee_shifts import...")
    try:
        shifts = extract_employee_shifts.extract_employee_shifts("test")
        print(f"Extract function works, found {len(shifts)} shifts")
    except Exception as e:
        print(f"Extract function error: {e}")
