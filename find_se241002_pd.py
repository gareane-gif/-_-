
import pandas as pd
import json

path = r'xls\قسم المساحة25-26.xls'
sheets = pd.read_excel(path, sheet_name=None, engine='xlrd')

found = False
for name, df in sheets.items():
    print(f"\nSearching sheet: {name}")
    # Search for SE241002
    for r in range(len(df)):
        row = df.iloc[r]
        if any('SE241002' in str(val) for val in row):
            print(f"Found at row {r}")
            # Get surrounding rows
            data = df.iloc[max(0, r-2):min(len(df), r+8)].astype(str).values.tolist()
            print(json.dumps(data, ensure_ascii=False, indent=2))
            found = True
            break
    if found: break

if not found:
    print("Not found")
