import xlrd
import os

def check_structure(file_path):
    print(f"\n--- Checking: {os.path.basename(file_path)} ---")
    try:
        wb = xlrd.open_workbook(file_path, encoding_override='cp1256')
        for sheet in wb.sheets():
            print(f"Sheet: {sheet.name}")
            # Find a row that looks like a student ID
            found_id = False
            for r in range(sheet.nrows):
                row_vals = [str(sheet.cell_value(r, c)).strip() for c in range(min(20, sheet.ncols))]
                # look for something like CO231004 or AC252002
                if any(v.startswith(('CO', 'AC', 'EL', 'EN', 'SU', 'ME')) for v in row_vals):
                    print(f"Found potential student ID at row {r+1}: {row_vals}")
                    # Print following 6 rows to see structure
                    for i in range(1, 7):
                        if r + i < sheet.nrows:
                            next_row = [str(sheet.cell_value(r + i, c)).strip() for c in range(min(20, sheet.ncols))]
                            print(f"  R{r+i+1}: {next_row}")
                    found_id = True
                    break
            if not found_id:
                print("No student ID pattern found in first 100 rows.")
    except Exception as e:
        print(f"Error checking {file_path}: {e}")

if __name__ == "__main__":
    xls_dir = r'xls'
    for f in os.listdir(xls_dir):
        if f.endswith('.xls'):
            check_structure(os.path.join(xls_dir, f))
