import pandas as pd
import os

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
f = 'قسم المساحة25-26.xls'
path = os.path.join(xls_dir, f)

try:
    sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', header=None)
except Exception as e:
    print(f"Error: {e}")
    exit(1)

for name, df in sheets.items():
    print(f"--- Sheet: {name} ---")
    data = df.values.tolist()
    row5 = data[4] # Row 5
    for c_idx, cell in enumerate(row5):
        val = str(cell)
        print(f"Col {c_idx}: '{val}' | hex: {val.encode('utf-8').hex()}")
    break
