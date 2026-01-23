import pandas as pd
import os
import re

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
files = [
    'قسم المحاسبة25-26.xls',
    'قسم المساحة25-26.xls'
]

for f in files:
    path = os.path.join(xls_dir, f)
    print(f"\n======= File: {f} =======")
    try:
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd')
    except Exception as e:
        print(f"Error: {e}")
        continue
    
    for name, df in sheets.items():
        print(f"--- Sheet: {name} ---")
        df = df.fillna('')
        data = df.astype(str).values.tolist()
        # Dump Rows 1 to 10
        for i in range(10):
            if i < len(data):
                print(f"Row {i+2}: {' | '.join(data[i])}")
        break
