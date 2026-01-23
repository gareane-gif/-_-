import pandas as pd
import os

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
files = [
    'قسم المحاسبة25-26.xls',
    'قسم المساحة25-26.xls'
]

for f in files:
    path = os.path.join(xls_dir, f)
    print(f"\n======= File: {f} =======")
    try:
        # Use header=None to see everything
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', header=None)
    except Exception as e:
        print(f"Error: {e}")
        continue
    
    for name, df in sheets.items():
        print(f"--- Sheet: {name} ---")
        df = df.fillna('')
        data = df.astype(str).values.tolist()
        # Dump Rows 1 to 15
        for i in range(min(15, len(data))):
            print(f"Row {i+1}: {' | '.join(data[i])}")
        break
