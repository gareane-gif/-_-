import pandas as pd
import os

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
files = [
    'قسم الكهرباء25-26.xls',
    'قسم الميكانيكا 25-26.xls'
]

for f in files:
    path = os.path.join(xls_dir, f)
    print(f"\n======= File: {f} =======")
    try:
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', header=None)
    except Exception as e:
        print(f"Error: {e}")
        continue
    
    for name, df in sheets.items():
        print(f"--- Sheet: {name} ---")
        data = df.values.tolist()
        # Find some student or keywords
        found = False
        for i in range(len(data)):
            row_str = " ".join([str(c) for c in data[i]])
            if any(k in row_str for k in ['24', '23']): # Likely an ID part
                print(f"Example Student Row {i+1}: {row_str}")
                for next_r in range(1, 6):
                    if i + next_r < len(data):
                        print(f"  +{next_r}: {' | '.join([str(c) for c in data[i+next_r]])}")
                found = True
                break
        if found: break
