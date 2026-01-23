import xlrd

def find_gpa_global(file_path, student_id):
    try:
        wb = xlrd.open_workbook(file_path, encoding_override='cp1256')
        print(f"File: {file_path}")
        
        # 1. Search for GPA keywords globally
        for sheet in wb.sheets():
            print(f"Sheet: {sheet.name}")
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    val = str(sheet.cell_value(r, c))
                    if "المعدل" in val or "GPA" in val.upper() or "فصلي" in val:
                        # Print some context around the found GPA label
                        context = [sheet.cell_value(r, max(0, c+i)) for i in range(5)]
                        print(f"  FOUND label at R{r+1} C{c+1}: '{val}', Context: {context}")
            
            # 2. Search for the student and look at the end of their row
            found_r = -1
            for r in range(sheet.nrows):
                if student_id in str(sheet.cell_value(r, 3)) or student_id in str(sheet.cell_value(r, 1)):
                    found_r = r
                    break
            
            if found_r != -1:
                print(f"  Found student R{found_r+1}. Checking last few columns...")
                # Usually GPA is in one of the last columns
                for c in range(max(0, sheet.ncols - 10), sheet.ncols):
                    print(f"    Col {c+1}: {sheet.cell_value(found_r, c)}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    find_gpa_global(r'xls\قسم الحاسوب25-26.xls', '231017')
