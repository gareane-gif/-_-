# -*- coding: utf-8 -*-
import xlrd

def dump_headers():
    wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')
    for sheet in wb.sheets():
        print(f"\nSheet: {sheet.name}")
        # Look for the header row
        header_row = -1
        for r in range(min(sheet.nrows, 30)):
            row = [str(sheet.cell(r, c).value).strip() for c in range(sheet.ncols)]
            if any(kw in "".join(row) for kw in ['اسم', 'الطالب', 'قيد']):
                header_row = r
                break
        
        if header_row != -1:
            headers = [str(sheet.cell(header_row, c).value).strip() for c in range(sheet.ncols)]
            print(f"Header at Row {header_row + 1}:")
            for i, h in enumerate(headers):
                if h:
                    print(f"  Col {i}: {h}")
        else:
            print("  No header found in first 30 rows.")

if __name__ == "__main__":
    dump_headers()
