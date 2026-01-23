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
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd')
    except Exception as e:
        print(f"Error: {e}")
        continue
    
    for name, df in sheets.items():
        print(f"--- Sheet: {name} ---")
        df = df.fillna('')
        data = df.values.tolist()
        # Print first 10 rows to see headers
        for i in range(min(15, len(data))):
            print(f"Row {i+2}: {' | '.join([str(c) for c in data[i]])}")
        break # Only first sheet for now
