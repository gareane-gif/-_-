
import xlrd
import os

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]

for file in files:
    path = os.path.join(xls_dir, file)
    print(f"\n{file}:")
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        for sheet in wb.sheets():
            print(f"  - {sheet.name}")
    except Exception as e:
        print(f"  Error: {e}")
