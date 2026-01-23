
import xlrd
import os
import json

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]
all_headers = {}

for file in files:
    path = os.path.join(xls_dir, file)
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0)
        file_headers = []
        for r in range(min(15, sheet.nrows)):
            row = [str(sheet.cell_value(r, i)).strip() for i in range(sheet.ncols)]
            file_headers.append(row)
        all_headers[file] = file_headers
    except Exception as e:
        all_headers[file] = f"Error: {e}"

with open('all_headers.json', 'w', encoding='utf-8') as f:
    json.dump(all_headers, f, ensure_ascii=False, indent=2)
