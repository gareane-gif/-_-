
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
                    block_labels = []
                    for row_off in range(6):
                        if r + row_off < sheet.nrows:
                            # look for a label in the first few columns
                            label = ""
                            for c_search in range(min(sheet.ncols, 15)):
                                v = str(sheet.cell_value(r + row_off, c_search)).strip()
                                if any(kw in v for kw in ['أعمال', 'نهائي', 'مجموع', 'معدل', 'نتيجة', 'حالة', 'تقدير']):
                                    label = v
                                    break
                            block_labels.append(label)
                    results[file] = block_labels
                    found_block = True
                    break
            if found_block: break
    except Exception as e:
        results[file] = f"Error: {e}"

with open('structure_results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("Results written to structure_results.json")
