# -*- coding: utf-8 -*-
import xlrd

wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')

print("Sheet information:")
for i, sheet in enumerate(wb.sheets()):
    print(f"\nSheet {i+1}:")
    print(f"  Name: {sheet.name}")
    print(f"  Rows: {sheet.nrows}")
    print(f"  Cols: {sheet.ncols}")
    
    # Search for CO231001 in this sheet
    found = False
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            val = str(sheet.cell(r, c).value)
            if 'CO231001' in val:
                print(f"  >>> CONTAINS CO231001 at row {r+1}, col {c+1}")
                found = True
                break
        if found:
            break
    
    if not found:
        print(f"  (CO231001 not in this sheet)")
