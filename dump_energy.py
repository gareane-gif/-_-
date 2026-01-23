
import xlrd

path = r'xls\قسم الطاقة25-26.xls'
wb = xlrd.open_workbook(path, encoding_override='cp1256')
sheet = wb.sheet_by_index(0)
for r in range(min(100, sheet.nrows)):
    val = str(sheet.cell_value(r, 3)).strip()
    if val:
        print(f"Row {r+1}: {val}")
