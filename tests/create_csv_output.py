import pandas as pd

# Read the Excel file and convert to CSV
df = pd.read_excel('../ostardo_turni.xlsx')
df.to_csv('../ostardo_turni_improved.csv', index=False)
print("Created CSV version: ../ostardo_turni_improved.csv")
print("\nPreview of the data:")
print(df.to_string())
