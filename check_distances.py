
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
        
        student_rows = []
        for r in range(sheet.nrows):
            for c in range(min(sheet.ncols, 10)):
                val = str(sheet.cell_value(r, c)).strip()
                if any(kw in val for kw in ['AC', 'CO', 'EL', 'PO', 'SU', 'ME']) and len(val) >= 6:
                    student_rows.append(r)
                    break
        
        if len(student_rows) > 1:
            diffs = [student_rows[i+1] - student_rows[i] for i in range(len(student_rows)-1)]
            results[file] = diffs
        else:
            results[file] = "Not enough students"
    except Exception as e:
        results[file] = f"Error: {e}"

with open('student_distances.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
