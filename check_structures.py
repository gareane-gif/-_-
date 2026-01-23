
import xlrd
import os

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]

for file in files:
    path = os.path.join(xls_dir, file)
    print(f"\nAnalyzing: {file}")
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0) # check first sheet
        
        student_rows = []
        for r in range(sheet.nrows):
            for c in range(min(sheet.ncols, 10)):
                val = str(sheet.cell_value(r, c)).strip()
                # Simple heuristic for student ID
                if val.startswith(('AC', 'CO', 'EL', 'PO', 'SU', 'ME')) and len(val) >= 6:
                    student_rows.append(r)
                    break
                    
        if len(student_rows) > 1:
            diffs = [student_rows[i+1] - student_rows[i] for i in range(len(student_rows)-1)]
            from collections import Counter
            counts = Counter(diffs)
            print(f"  Rows between students: {counts}")
        else:
            print("  Not enough students found to determine structure.")
    except Exception as e:
        print(f"  Error: {e}")
