
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
        
        found_block = False
        for r in range(sheet.nrows):
            for c in range(min(sheet.ncols, 10)):
                val = str(sheet.cell_value(r, c)).strip()
                if val.startswith(('AC', 'CO', 'EL', 'PO', 'SU', 'ME')) and len(val) >= 6:
                    # Found a student row. Now look for labels in subsequent rows.
                    labels = {}
                    for row_off in range(1, 10):
                        if r + row_off < sheet.nrows:
                            for c_search in range(min(sheet.ncols, 15)):
                                v = str(sheet.cell_value(r + row_off, c_search)).strip()
                                if 'أعمال' in v or 'اعمال' in v: labels['Work'] = row_off
                                if 'نهائي' in v: labels['Final'] = row_off
                                if 'مجموع' in v: labels['Total'] = row_off
                                if 'معدل' in v or 'GPA' in v: labels['GPA'] = row_off
                    results[file] = labels
                    found_block = True
                    break
            if found_block: break
    except Exception as e:
        results[file] = f"Error: {e}"

with open('label_offsets.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
