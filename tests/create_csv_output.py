import pandas as pd

# Read the Excel file and convert to CSV
df = pd.read_excel('../employee_shifts.xlsx')
df.to_csv('../employee_shifts_improved.csv', index=False)
print("Created CSV version: ../employee_shifts_improved.csv")
print("\nPreview of the data:")
print(df.to_string())
