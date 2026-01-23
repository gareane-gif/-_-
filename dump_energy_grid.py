
import xlrd
import json

path = r'xls\قسم الطاقة25-26.xls'
wb = xlrd.open_workbook(path, encoding_override='cp1256')
sheet = wb.sheet_by_index(0)
rows = []
for r in range(min(50, sheet.nrows)):
    row = [str(sheet.cell_value(r, c)).strip() for c in range(min(15, sheet.ncols))]
    rows.append(row)

with open('energy_grid.json', 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)
