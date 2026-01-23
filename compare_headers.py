
import xlrd
import json

def get_headers(file_path):
    wb = xlrd.open_workbook(file_path, encoding_override='cp1256')
    sheet = wb.sheet_by_index(0)
    headers = []
    for r in range(10): # look at first 10 rows
        row = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        headers.append(row)
    return headers

data = {
    "Accounting": get_headers('xls/قسم المحاسبة25-26.xls'),
    "Computer": get_headers('xls/قسم الحاسوب25-26.xls')
}

with open('header_comparison.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
