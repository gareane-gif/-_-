
import xlrd
import os
import json

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]
results = {}

for file in files:
    path = os.path.join(xls_dir, file)
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0)
        rows = []
        for r in range(min(5, sheet.nrows)):
            row = [str(sheet.cell_value(r, c)).strip() for c in range(min(15, sheet.ncols))]
            rows.append(row)
        results[file] = rows
    except Exception as e:
        results[file] = f"Error: {e}"

with open('first_5_rows.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
