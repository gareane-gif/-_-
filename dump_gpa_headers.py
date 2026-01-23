
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
        found = False
        for r in range(min(20, sheet.nrows)):
            row = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
            for i, val in enumerate(row):
                if any(kw in val for kw in ['معدل', 'GPA', 'فصلي']):
                    results[file] = val
                    found = True
                    break
            if found: break
    except Exception as e:
        results[file] = f"Error: {e}"

with open('gpa_headers.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
