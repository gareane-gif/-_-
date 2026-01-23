
import xlrd
import os

def dump_student_detailed(filename, student_id):
    path = os.path.join('xls', filename)
    print(f"\n--- {filename} ({student_id}) ---")
    wb = xlrd.open_workbook(path, encoding_override='cp1256')
    sheet = wb.sheet_by_index(0)
    for r in range(sheet.nrows):
        for c in range(min(sheet.ncols, 10)):
            if student_id in str(sheet.cell_value(r, c)):
                print(f"Found {student_id} at Row {r+1}")
                # Print header for context
                header_row = 4 # based on prev output
                h_vals = [str(sheet.cell_value(header_row-1, col)).strip() for col in range(min(sheet.ncols, 20))]
                print(f"Headers: {' | '.join(h_vals)}")
                
                for row_off in range(-1, 6):
                    row_idx = r + row_off
                    if 0 <= row_idx < sheet.nrows:
                        vals = [str(sheet.cell_value(row_idx, col)).strip() for col in range(min(sheet.ncols, 20))]
                        print(f"Row {row_idx+1}: {' | '.join(vals)}")
                return
    print("Student not found")

dump_student_detailed('قسم المساحة25-26.xls', 'SE241002')
