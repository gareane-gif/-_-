
import xlrd
import os

def dump_student(filename, student_id):
    path = os.path.join('xls', filename)
    print(f"\n--- {filename} ({student_id}) ---")
    wb = xlrd.open_workbook(path, encoding_override='cp1256')
    sheet = wb.sheet_by_index(0)
    for r in range(sheet.nrows):
        for c in range(min(sheet.ncols, 10)):
            if student_id in str(sheet.cell_value(r, c)):
                for row_off in range(5):
                    row_idx = r + row_off
                    vals = [str(sheet.cell_value(row_idx, col)).strip() for col in range(min(sheet.ncols, 20))]
                    print(f"Row {row_idx+1}: {' | '.join(vals)}")
                return
    print("Student not found")

dump_student('قسم المحاسبة25-26.xls', 'AC') # find any AC student
dump_student('قسم الحاسوب25-26.xls', 'CO') # find any CO student
