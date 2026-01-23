# -*- coding: utf-8 -*-
import xlrd
import sys

# Load the Excel file (.xls format)
wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')

print("="*80)
print("SEARCHING ALL SHEETS FOR CO231001")
print("="*80)

found_sheet = None
found_row = None
found_col = None

for sheet_idx, sheet in enumerate(wb.sheets()):
    print(f"\nSheet {sheet_idx + 1}: {sheet.name} ({sheet.nrows} rows x {sheet.ncols} cols)")
    
    # Search for CO231001
    for row_idx in range(sheet.nrows):
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            value = str(cell.value).strip()
            if 'CO231001' in value:
                col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                print(f"  >>> FOUND CO231001 at {col_letter}{row_idx + 1}: '{value}'")
                if not found_sheet:
                    found_sheet = sheet
                    found_row = row_idx
                    found_col = col_idx

if not found_sheet:
    print("\n" + "="*80)
    print("CO231001 NOT FOUND IN ANY SHEET!")
    print("="*80)
    print("\nListing all student IDs starting with CO23 in all sheets:")
    for sheet in wb.sheets():
        print(f"\nSheet: {sheet.name}")
        ids_found = []
        for row_idx in range(sheet.nrows):
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                value = str(cell.value).strip()
                if value.startswith('CO23') and len(value) >= 8:
                    ids_found.append(value)
        
        if ids_found:
            unique_ids = sorted(set(ids_found))
            print(f"  Found {len(unique_ids)} unique IDs: {', '.join(unique_ids[:10])}")
            if len(unique_ids) > 10:
                print(f"  ... and {len(unique_ids) - 10} more")
else:
    print("\n" + "="*80)
    print(f"ANALYZING STUDENT DATA IN SHEET: {found_sheet.name}")
    print(f"Found at Row {found_row + 1}, Column {chr(65 + found_col)}")
    print("="*80)
    
    # Show full row
    print(f"\nFull Row {found_row + 1} data:")
    for c in range(min(20, found_sheet.ncols)):
        cell = found_sheet.cell(found_row, c)
        col_l = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
        if cell.value not in (None, ''):
            print(f"  {col_l}{found_row + 1}: {cell.value}")
    
    # Show surrounding rows (student row + grade rows)
    print(f"\nRows {found_row + 1} to {found_row + 5} (student + grade components):")
    for r in range(found_row, min(found_sheet.nrows, found_row + 5)):
        print(f"\nRow {r + 1}:")
        for c in range(min(15, found_sheet.ncols)):
            cell = found_sheet.cell(r, c)
            col_l = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
            if cell.value not in (None, ''):
                print(f"  {col_l}{r + 1}: {cell.value}")
    
    # Find header row
    print("\n" + "="*80)
    print("SEARCHING FOR HEADER ROW (looking backwards from student row)")
    print("="*80)
    
    for r in range(max(0, found_row - 25), found_row):
        row_text = []
        for c in range(min(10, found_sheet.ncols)):
            cell = found_sheet.cell(r, c)
            value = str(cell.value) if cell.value else ""
            if any(kw in value for kw in ['اسم', 'الطالب', 'القيد', 'الرقم', 'تسلسل', 'Name', 'ID']):
                row_text.append(f"{chr(65+c)}:{value}")
        
        if row_text:
            print(f"Row {r + 1} (potential header): {' | '.join(row_text)}")
