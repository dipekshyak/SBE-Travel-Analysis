import pandas as pd

file_path = "SBE Travel Spreadsheet working.xlsx"
all_sheets = pd.read_excel(file_path, sheet_name=None)

print(f"Found {len(all_sheets)} tab(s) in the file.\n")

for name, df in all_sheets.items():
    print(f"=== Tab: {name} ===")
    print(f"Rows: {len(df)}")
    print(f"Columns: {list(df.columns)}")
    print()