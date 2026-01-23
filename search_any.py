#!/usr/bin/env python
import os
import re
import sys

try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

if len(sys.argv) < 2:
    print('Usage: python search_any.py <text>')
    sys.exit(1)

term = sys.argv[1]

xls_dir = os.path.join(os.path.dirname(__file__), 'xls')
files = [f for f in os.listdir(xls_dir) if f.lower().endswith(('.xls', '.xlsx'))]
if not files:
    print('No Excel files found in', xls_dir)
    sys.exit(0)

found = False
for file in files:
    path = os.path.join(xls_dir, file)
    ext = os.path.splitext(file)[1].lower()
    try:
        if ext == '.xls':
            sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', dtype=str)
        else:
            sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl', dtype=str)
    except Exception as e:
        print('Could not open', file, '->', e)
        continue

    for sheet_name, df in sheets.items():
        df = df.fillna('')
        data = df.astype(str).values.tolist()
        nrows = len(data)
        for r in range(nrows):
            row = data[r]
            for c in range(len(row)):
                cell = str(row[c])
                if term in cell:
                    print('\nMATCH (substring)')
                    print('File:', file)
                    print('Sheet:', sheet_name)
                    print('Cell (row,col):', r+1, c+1)
                    print('Cell value:', cell)
                    found = True

if not found:
    print('No substring matches found for', term)
else:
    print('\nSubstring search completed.')
