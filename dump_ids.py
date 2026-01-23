
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
        ids = []
        for r in range(sheet.nrows):
            for c in range(min(sheet.ncols, 10)):
                val = str(sheet.cell_value(r, c)).strip()
                # Use a very broad heuristic to find candidate IDs
                if any(char.isdigit() for char in val) and len(val) >= 4:
                    if any(kw in val for kw in ['AC', 'CO', 'EL', 'PO', 'SU', 'ME', '2', '1']):
                        ids.append(val)
                        if len(ids) > 10: break
            if len(ids) > 10: break
        results[file] = ids
    except Exception as e:
        results[file] = f"Error: {e}"

with open('id_samples.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
