# -*- coding: utf-8 -*-
import xlrd

def list_all_sheets():
    wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')
    print("Available Sheets:")
    for i, sheet in enumerate(wb.sheets()):
        print(f"{i}: {sheet.name} ({sheet.nrows} rows, {sheet.ncols} cols)")

if __name__ == "__main__":
    list_all_sheets()
