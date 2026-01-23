# -*- coding: utf-8 -*-
import xlrd
import json

def dump_sheet_top():
    wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')
    sheet = wb.sheet_by_name("الخامس حاسوب")
    
    rows = []
    for r in range(min(sheet.nrows, 15)):
        rows.append([str(sheet.cell(r, c).value).strip() for c in range(sheet.ncols)])
    
    with open("sheet_top.json", "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    dump_sheet_top()
