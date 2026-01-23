import xlrd
import os

def check_surveying():
    file_path = 'xls/قسم المساحة25-26.xls'
    rb = xlrd.open_workbook(file_path)
    sheet = rb.sheet_by_name('الاول مساحة')
    for r in range(sheet.nrows):
        row = sheet.row_values(r)
        if 'SE252003' in str(row):
            print(f"Row {r+1}: {row}")

check_surveying()
