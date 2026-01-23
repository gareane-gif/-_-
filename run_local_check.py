#!/usr/bin/env python
"""Quick local checker to verify Python Excel reading works with available files.

Usage: python run_local_check.py
It will list files in `xls/`, attempt to read the first sheet of each workbook and print row/col counts and a small sample.
"""
import os
import sys

try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

HERE = os.path.dirname(__file__)
xls_dir = os.path.join(HERE, 'xls')
if not os.path.isdir(xls_dir):
    print('xls directory not found:', xls_dir)
    sys.exit(1)

files = [f for f in os.listdir(xls_dir) if f.lower().endswith(('.xls', '.xlsx'))]
if not files:
    print('No Excel files found in', xls_dir)
    sys.exit(0)

for fn in files:
    path = os.path.join(xls_dir, fn)
    ext = os.path.splitext(fn)[1].lower()
    print('\n---', fn)
    try:
        if ext == '.xls':
            sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', dtype=str)
        else:
            sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl', dtype=str)
    except Exception as e:
        print('  Could not open ->', e)
        continue

    print('  Sheets:', list(sheets.keys()))
    for name, df in sheets.items():
        r = len(df)
        c = len(df.columns)
        print(f'   sheet: {name!r}  rows={r} cols={c}')
        if r:
            print('    sample row 1:', df.fillna('').astype(str).iloc[0].tolist()[:10])
        break

print('\nLocal check complete.')
