# -*- coding: utf-8 -*-
import xlrd
import json

def analyze_student():
    wb = xlrd.open_workbook(r'xls\قسم الحاسوب25-26.xls', encoding_override='cp1256')
    student_id = "CO231001"
    
    results = []
    
    for sheet in wb.sheets():
        found = False
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                val = str(sheet.cell(r, c).value).strip()
                if student_id in val:
                    found = True
                    # Extract headers (heuristic: find row with 'اسم')
                    header_row = -1
                    for hr in range(max(0, r-30), r):
                        hdata = [str(sheet.cell(hr, col).value).strip() for col in range(sheet.ncols)]
                        if any(kw in "".join(hdata) for kw in ['اسم', 'الطالب', 'قيد']):
                            header_row = hr
                            break
                    
                    headers = []
                    if header_row != -1:
                        headers = [str(sheet.cell(header_row, col).value).strip() for col in range(sheet.ncols)]
                    
                    # Extract student block (r to r+5)
                    student_data = []
                    for sr in range(r, min(sheet.nrows, r+6)):
                        srdata = [str(sheet.cell(sr, col).value).strip() for col in range(sheet.ncols)]
                        student_data.append(srdata)
                    
                    results.append({
                        "sheet": sheet.name,
                        "header_row": header_row,
                        "headers": headers,
                        "student_row": r,
                        "student_data": student_data
                    })
                    break
            if found: break

    with open("student_analysis.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    analyze_student()
