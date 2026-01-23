
import xlrd

path = r'xls\قسم الطاقة25-26.xls'
wb = xlrd.open_workbook(path, encoding_override='cp1256')
sheet = wb.sheet_by_index(0)
for r in range(5, 15):
    row = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
    print(f"Row {r+1}: {' | '.join(row)}")
