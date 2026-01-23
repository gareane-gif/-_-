
import xlrd
import os

path = os.path.join('xls', 'قسم الكهرباء25-26.xls')
wb = xlrd.open_workbook(path, encoding_override='cp1256')

for sheet in wb.sheets():
    print(f"\n--- Sheet: {sheet.name} ---")
    for r in range(min(20, sheet.nrows)):
        vals = [str(sheet.cell_value(r, c)).strip() for c in range(min(sheet.ncols, 20))]
        print(f"Row {r+1}: {vals}")
