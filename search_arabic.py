import pandas as pd
import os
import re

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
f = 'قسم المساحة25-26.xls'
path = os.path.join(xls_dir, f)

def is_arabic(s):
    return any('\u0600' <= c <= '\u06FF' for c in str(s))

try:
    sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', header=None)
except Exception as e:
    print(f"Error: {e}")
    exit(1)

for name, df in sheets.items():
    print(f"--- Sheet: {name} ---")
    data = df.values.tolist()
    for r_idx, row in enumerate(data):
        for c_idx, cell in enumerate(row):
            val = str(cell).strip()
            if is_arabic(val) and len(val) > 4:
                # Potential subject name or other text
                print(f"Row {r_idx+1}, Col {c_idx+1}: {val}")
    break
