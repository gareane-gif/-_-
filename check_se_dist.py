
import xlrd
import os

path = os.path.join('xls', 'قسم المساحة25-26.xls')
wb = xlrd.open_workbook(path, encoding_override='cp1256')
sheet = wb.sheet_by_name('الرابع مساحة')

student_id = 'SE241002'
found_r = -1
for r in range(sheet.nrows):
    for c in range(sheet.ncols):
        if student_id == str(sheet.cell_value(r, c)).strip():
            found_r = r
            break
    if found_r != -1: break

if found_r != -1:
    print(f"Student found at row {found_r + 1}")
    # Header row (row 5 in 1-based, index 4)
    header_r = 4
    for c in range(15, 21):
        header = sheet.cell_value(header_r, c)
        val = sheet.cell_value(found_r, c)
        work = sheet.cell_value(found_r + 1, c)
        final = sheet.cell_value(found_r + 2, c)
        total = sheet.cell_value(found_r + 3, c)
        print(f"Col {c}: Header='{header}' | Units='{val}' | Work='{work}' | Final='{final}' | Total='{total}'")
else:
    print("Student not found")
