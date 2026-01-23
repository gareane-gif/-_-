import xlrd

def find_gpa_col(file_path):
    try:
        wb = xlrd.open_workbook(file_path, encoding_override='cp1256')
        sheet = wb.sheets()[0]
        print(f"Sheet: {sheet.name}")
        
        # Look for "المعدل" in rows 1-20
        gpa_col = -1
        for r in range(min(30, sheet.nrows)):
            for c in range(sheet.ncols):
                val = str(sheet.cell_value(r, c))
                # "المعدل" in CP1256 or similar
                if "المعدل" in val or "GPA" in val.upper() or "معدل" in val:
                    print(f"FOUND 'المعدل' at R{r+1} C{c+1}: '{val}'")
                    gpa_col = c
        
        if gpa_col != -1:
            print(f"Checking GPA column (Col {gpa_col+1}) for some students...")
            for r in range(sheet.nrows):
                if "CO" in str(sheet.cell_value(r, 3)):
                    # GPA should be at r+4
                    if r + 4 < sheet.nrows:
                        print(f"  Student {sheet.cell_value(r, 3)} at R{r+1}: GPA val at R{r+5} = {sheet.cell_value(r+4, gpa_col)}")
        else:
            print("Could not find 'المعدل' label. Scanning last columns for numeric patterns...")
            for c in range(sheet.ncols - 1, max(-1, sheet.ncols - 10), -1):
                nums = []
                for r in range(min(100, sheet.nrows)):
                    v = sheet.cell_value(r, c)
                    if isinstance(v, (int, float)) and 0 < v <= 100:
                        nums.append(v)
                if len(nums) > 5:
                    print(f"  Col {c+1} looks like a candidate (many numbers): {nums[:5]}...")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    find_gpa_col(r'xls\قسم الحاسوب25-26.xls')
