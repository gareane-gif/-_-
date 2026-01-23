import xlrd
import sys

# Load the Excel file (.xls format)
wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls')

# Find the sheet with "حاسوب" in the name
target_sheet = None
for sheet in wb.sheets():
    if 'حاسوب' in sheet.name and 'ثالث' in sheet.name:
        target_sheet = sheet
        print(f"Found sheet: {sheet.name}")
        break

if not target_sheet:
    print("Sheet not found! Available sheets:")
    for sheet in wb.sheets():
        print(f"  - {sheet.name}")
    sys.exit(1)

# Deep Search found CO231001 at Row 27, Column 4 (D27)
# Note: xlrd uses 0-based indexing, so Row 27 = index 26, Column 4 (D) = index 3
print("\n" + "="*80)
print("EXAMINING AREA AROUND ROW 27, COLUMN 4 (where CO231001 was found)")
print("="*80)

# Check rows 20-35, columns A-J (indices 19-34, 0-9)
for row_idx in range(19, 35):
    row_data = []
    for col_idx in range(0, 10):  # A-J
        try:
            cell = target_sheet.cell(row_idx, col_idx)
            value = cell.value
            if value not in (None, ''):
                col_letter = chr(65 + col_idx)  # A=65
                row_num = row_idx + 1
                row_data.append(f"[{col_letter}{row_num}:{value}]")
        except:
            pass
    
    if row_data:
        print(f"Row {row_idx + 1}: {' '.join(row_data)}")

print("\n" + "="*80)
print("ANALYZING HEADER DETECTION")
print("="*80)

# Look for header row (should contain keywords like اسم, القيد, etc.)
for row_idx in range(0, 30):
    row_data = []
    for col_idx in range(0, 10):
        try:
            cell = target_sheet.cell(row_idx, col_idx)
            value = str(cell.value) if cell.value else ""
            if any(keyword in value for keyword in ['اسم', 'الطالب', 'القيد', 'الرقم', 'تسلسل']):
                col_letter = chr(65 + col_idx)
                row_data.append(f"[{col_letter}:{value}]")
        except:
            pass
    
    if row_data:
        print(f"Row {row_idx + 1} (potential header): {' '.join(row_data)}")

print("\n" + "="*80)
print("STUDENT DATA EXTRACTION (Row 27)")
print("="*80)

# Extract the full row 27 data (index 26)
print("Full Row 27 data:")
for col_idx in range(0, 20):
    try:
        cell = target_sheet.cell(26, col_idx)
        col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
        print(f"  {col_letter}27: {cell.value}")
    except:
        pass

# Check rows around 27 for grade components (work, final, total)
print("\nRows 27-31 (student + grade rows):")
for row_idx in range(26, 31):  # 0-based: 26-30
    print(f"\nRow {row_idx + 1}:")
    for col_idx in range(0, 10):
        try:
            cell = target_sheet.cell(row_idx, col_idx)
            col_letter = chr(65 + col_idx)
            if cell.value not in (None, ''):
                print(f"  {col_letter}{row_idx + 1}: {cell.value}")
        except:
            pass

