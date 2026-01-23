import os
import pandas as pd

HERE = os.path.dirname(__file__)
xls_dir = os.path.join(HERE, 'xls')

matches = [
    ('قسم الحاسوب25-26.xls', 'الثالث حاسوب', 6, 'result_حاسو_الثالث_row6.csv'),
    ('قسم المحاسبة25-26.xls', 'الرابع محاسبة', 36, 'result_محاسبة_الرابع_row36.csv'),
]

outs = []
for fname, sheet_name, row_no, outname in matches:
    path = os.path.join(xls_dir, fname)
    if not os.path.exists(path):
        print('Missing file', path)
        continue
    ext = os.path.splitext(path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine=engine, dtype=str)
    except Exception as e:
        print('Failed to read', fname, sheet_name, '->', e)
        continue
    df = df.fillna('')
    idx = row_no - 1
    if idx < 0 or idx >= len(df):
        print('Row out of range for', fname, sheet_name, row_no)
        continue
    row = df.iloc[[idx]]
    outpath = os.path.join(HERE, outname)
    row.to_csv(outpath, index=False, encoding='utf-8-sig')
    print('Wrote', outpath)
    outs.append(outpath)

if outs:
    print('\nGenerated CSVs:')
    for p in outs:
        print(' -', p)
else:
    print('No CSVs generated.')
