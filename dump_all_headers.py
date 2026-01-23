
import xlrd
import os

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]

for file in files:
    path = os.path.join(xls_dir, file)
    print(f"\n{'='*20} {file} {'='*20}")
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0)
        
        # Look for a row containing typical header keywords
        for r in range(min(50, sheet.nrows)):
            row_vals = [str(sheet.cell_value(r, c)).strip() for c in range(min(sheet.ncols, 30))]
            if any(kw in "".join(row_vals) for kw in ['اسم', 'الطالب', 'القيد', 'الرقم']):
                print(f"Header at row {r+1}: {' | '.join([v for v in row_vals if v])}")
                break
    except Exception as e:
        print(f"  Error: {e}")
