import pandas as pd
import os

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
f = 'قسم المساحة25-26.xls'
path = os.path.join(xls_dir, f)

print(f"======= File: {f} =======")
try:
    sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', header=None)
except Exception as e:
    print(f"Error: {e}")
    exit(1)

for name, df in sheets.items():
    print(f"--- Sheet: {name} ---")
    data = df.values.tolist()
    for i in range(min(10, len(data))):
        row = [str(c) for c in data[i]]
        print(f"Row {i+1}: {' | '.join(row)}")
    break
