import xlrd
import os

xls_files = [
    'قسم المحاسبة25-26.xls',
    'قسم الحاسوب25-26.xls',
    'قسم الطاقة25-26.xls',
    'قسم الكهرباء25-26.xls',
    'قسم المساحة25-26.xls',
    'قسم الميكانيكا 25-26.xls'
]

xls_dir = 'xls'

for file in xls_files:
    path = os.path.join(xls_dir, file)
    if not os.path.exists(path):
        print(f"File not found: {path}")
        continue
    
    print(f"\n{'='*50}")
    print(f"FILE: {file}")
    print(f"{'='*50}")
    
    try:
        wb = xlrd.open_workbook(path, encoding_override='cp1256')
        sheet = wb.sheet_by_index(0)
        print(f"Sheet: {sheet.name}")
        print(f"Rows: {sheet.nrows}, Cols: {sheet.ncols}")
        
        # Look for a student-like ID
        found_row = -1
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                val = str(sheet.cell(r, c).value)
                if any(x in val for x in ['231', '252', '241', '221']):
                    found_row = r
                    break
            if found_row != -1:
                break
        
        if found_row != -1:
            print(f"Found student-like ID at row {found_row + 1}")
            start = max(0, found_row - 2)
            end = min(sheet.nrows, found_row + 8)
            for r in range(start, end):
                row_vals = []
                for c in range(min(15, sheet.ncols)):
                    val = str(sheet.cell(r, c).value).strip()
                    row_vals.append(val[:20])
                print(f"R{r+1:02}: {' | '.join(row_vals)}")
        else:
            print("No student-like ID found in first 100 rows.")
            
    except Exception as e:
        print(f"Error opening {file}: {e}")
