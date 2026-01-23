
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
        for r in range(min(5, sheet.nrows)):
            row = [str(sheet.cell_value(r, c)).strip() for c in range(min(15, sheet.ncols))]
            row_str = " ".join(row)
            if any(kw in row_str for kw in ['قيد', 'الرقم', 'الطالب', 'اسم']):
                print(f"  Row {r+1}: {row_str}")
    except Exception as e:
        print(f"  Error: {e}")
