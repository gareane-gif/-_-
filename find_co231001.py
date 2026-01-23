import xlrd
import sys

# Load the Excel file (.xls format)
wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')

# Find the sheet with "حاسوب" in the name
target_sheet = None
for sheet in wb.sheets():
    sheet_name = sheet.name
    print(f"Checking sheet: {sheet_name}")
    if 'حاسوب' in sheet_name or 'ثالث' in sheet_name or len(wb.sheets()) == 1:
        target_sheet = sheet
        print(f"Using sheet: {sheet_name}")
        break

if not target_sheet:
    print("Sheet not found! Available sheets:")
    for sheet in wb.sheets():
        print(f"  - {sheet.name}")
    # Just use first sheet
    target_sheet = wb.sheet_by_index(0)
    print(f"Using first sheet: {target_sheet.name}")

print("\n" + "="*80)
print("SEARCHING FOR CO231001")
print("="*80)

# Search for CO231001
found_locations = []
for row_idx in range(target_sheet.nrows):
    for col_idx in range(target_sheet.ncols):
        cell = target_sheet.cell(row_idx, col_idx)
        value = str(cell.value).strip()
        if 'CO231001' in value or '231001' in value:
            col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
            found_locations.append((row_idx + 1, col_idx, col_letter, value))
            print(f"FOUND at {col_letter}{row_idx + 1}: '{value}'")

if not found_locations:
    print("CO231001 NOT FOUND in this sheet!")
    print("\nSearching for any ID starting with CO23...")
    for row_idx in range(target_sheet.nrows):
        for col_idx in range(target_sheet.ncols):
            cell = target_sheet.cell(row_idx, col_idx)
            value = str(cell.value).strip()
            if value.startswith('CO23'):
                col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                print(f"  {col_letter}{row_idx + 1}: '{value}'")
else:
    # Analyze the first found location
    row_num, col_idx, col_letter, value = found_locations[0]
    row_idx = row_num - 1
    
    print("\n" + "="*80)
    print(f"ANALYZING STUDENT AT ROW {row_num}")
    print("="*80)
    
    # Show full row
    print(f"\nFull Row {row_num} data:")
    for c in range(min(20, target_sheet.ncols)):
        cell = target_sheet.cell(row_idx, c)
        col_l = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
        if cell.value not in (None, ''):
            print(f"  {col_l}{row_num}: {cell.value}")
    
    # Show surrounding rows
    print(f"\nRows {row_num-2} to {row_num+4}:")
    for r in range(max(0, row_idx - 2), min(target_sheet.nrows, row_idx + 5)):
        print(f"\nRow {r + 1}:")
        for c in range(min(15, target_sheet.ncols)):
            cell = target_sheet.cell(r, c)
            col_l = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
            if cell.value not in (None, ''):
                print(f"  {col_l}{r + 1}: {cell.value}")
    
    # Find header row
    print("\n" + "="*80)
    print("SEARCHING FOR HEADER ROW")
    print("="*80)
    
    for r in range(max(0, row_idx - 20), row_idx):
        row_text = []
        for c in range(min(10, target_sheet.ncols)):
            cell = target_sheet.cell(r, c)
            value = str(cell.value) if cell.value else ""
            if any(kw in value for kw in ['اسم', 'الطالب', 'القيد', 'الرقم', 'تسلسل', 'Name', 'ID']):
                row_text.append(f"{chr(65+c)}:{value}")
        
        if row_text:
            print(f"Row {r + 1} (potential header): {' | '.join(row_text)}")
