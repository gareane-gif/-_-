import xlrd

def inspect_sheet():
    file_path = 'xls/قسم المساحة25-26.xls'
    rb = xlrd.open_workbook(file_path)
    sheet = rb.sheet_by_name('الاول مساحة')
    print(f"Sheet: {sheet.name}")
    for r in range(10):
        print(f"Row {r+1}: {sheet.row_values(r)}")

inspect_sheet()
