
import xlrd
import os

xls_dir = 'xls'
files = [f for f in os.listdir(xls_dir) if f.endswith('.xls')]

for file in files:
    path = os.path.join(xls_dir, file)
    print(f"\n{file}:")
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0)
        for r in range(min(20, sheet.nrows)):
            row = [str(sheet.cell_value(r, i)).strip() for i in range(sheet.ncols)]
            for i, val in enumerate(row):
                if any(kw in val for kw in ['معدل', 'GPA', 'فصلي', 'تراكمي', 'عام']):
                    print(f"  Row {r+1}, Col {i}: {val}")
    except Exception as e:
        print(f"  Error: {e}")
