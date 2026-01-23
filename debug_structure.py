import pandas as pd
import os
import sys

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
files = [
    'قسم المحاسبة25-26.xls',
    'قسم المساحة25-26.xls'
]

student_ids = ['AC242041', 'SE241002']

for f in files:
    path = os.path.join(xls_dir, f)
    print(f"\n======= File: {f} =======")
    try:
        # Using engine='xlrd' for .xls files
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd')
    except Exception as e:
        print(f"Error opening {f}: {e}")
        continue
    
    for name, df in sheets.items():
        df = df.fillna('')
        data = df.values.tolist()
        for r_idx, row in enumerate(data):
            row_str = " ".join([str(c) for c in row])
            if any(sid in row_str for sid in student_ids):
                # We found the ID row (assume it's the start of the block)
                print(f"[{name}] Found Student Row {r_idx+2}")
                # Analyze next 10 rows for labels
                for i in range(1, 10):
                    if r_idx + i < len(data):
                        label_cell = str(data[r_idx+i][4]) if len(data[r_idx+i]) > 4 else ''
                        print(f"  Row {r_idx+i+2}, Col 4 (index 4): '{label_cell}' | hex: {label_cell.encode('utf-8').hex()}")
                print("-" * 30)
