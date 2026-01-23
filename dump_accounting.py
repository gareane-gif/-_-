#!/usr/bin/env python
import os, sys, re
try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

file_path = os.path.join(os.path.dirname(__file__), 'xls', 'قسم المحاسبة25-26.xls')
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

name = None
for n in sheets.keys():
    if 'الاول' in n or 'الأول' in n:
        name = n
        break
if name is None:
    name = list(sheets.keys())[0]
print('\nUsing sheet:', name)

df = sheets[name].fillna('')
nrows = len(df)
ncols = len(df.columns)
print('Rows:', nrows, 'Cols:', ncols)

for r in range(0, min(20, nrows)):
    row = df.iloc[r].tolist()
    parts = []
    for c, v in enumerate(row):
        parts.append(f'[{c+1}]:{v}')
    print(f'Row {r+1}:', ' | '.join(parts))

# find student AC252002
target = 'AC252002'
norm = re.compile(r'[^0-9]')
found = False
for r in range(nrows):
    row = df.iloc[r].tolist()
    for c, v in enumerate(row):
        s = str(v)
        if norm.sub('', s).lstrip('0') == norm.sub('', target).lstrip('0'):
            print('\nFound target at row', r+1, 'col', c+1, 'value:', v)
            # print header candidates (first 6 rows)
            print('\nHeader rows (1-8):')
            for hr in range(0, min(8, nrows)):
                print(f' H{hr+1}:', df.iloc[hr].tolist())
            print('\nMatched row and context:')
            start = max(0, r-2)
            end = min(nrows, r+6)
            for rr in range(start, end):
                print(f' R{rr+1}:', df.iloc[rr].tolist())
            found = True
            break
    if found:
        break

if not found:
    print('\nTarget not found in this sheet')
