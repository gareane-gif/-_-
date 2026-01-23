#!/usr/bin/env python
import os, sys, re
try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

file_path = os.path.join(os.path.dirname(__file__), 'xls', 'قسم الحاسوب25-26.xls')
if not os.path.exists(file_path):
    print('File not found:', file_path)
    sys.exit(1)

ext = os.path.splitext(file_path)[1].lower()
try:
    if ext == '.xls':
        sheets = pd.read_excel(file_path, sheet_name=None, engine='xlrd', dtype=str)
    else:
        sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl', dtype=str)
except Exception as e:
    print('Failed to read workbook:', e)
    sys.exit(1)

print('Sheets:', list(sheets.keys()))

sheet = None
selected_name = None
for name, df in sheets.items():
    if 'الثالث' in name:
        sheet = df
        selected_name = name
        print('\nUsing sheet:', name)
        break
if sheet is None:
    # pick first
    selected_name, sheet = next(iter(sheets.items()))
    print('\nNo sheet with "الثالث" found; using first sheet:', selected_name)

df = sheet.fillna('')
nrows = len(df)
ncols = len(df.columns)
print('Rows:', nrows, 'Cols:', ncols)

start = 1
end = min(40, nrows+1)
for r in range(start-1, end):
    row = df.iloc[r].tolist()
    display = []
    for c, v in enumerate(row):
        t = type(v).__name__
        display.append(f'[{c+1}]({t}):{v}')
    print(f'Row {r+1}:', ' | '.join(display))

# find cell values equal to 232029
print('\nSearch for 232029 (raw substrings):')
for r in range(nrows):
    row = df.iloc[r].tolist()
    for c, v in enumerate(row):
        if '232029' in str(v):
            print('Found at r', r+1, 'c', c+1, 'value:', v, 'type:', type(v).__name__)
