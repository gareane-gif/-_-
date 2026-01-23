
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
        results[file] = [sheet.name for sheet in wb.sheets()]
    except Exception as e:
        results[file] = f"Error: {e}"

with open('sheet_names.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
