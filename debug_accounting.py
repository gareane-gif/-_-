import pandas as pd
import os
import re

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
file_path = os.path.join(xls_dir, 'قسم المحاسبة25-26.xls')
target_id = 'AC242041'

def normalize_id(s):
    if s is None: return ''
    s = str(s).strip().upper()
    if any(c.isalpha() for c in s):
        s = re.sub(r'[^A-Z0-9]', '', s)
    else:
        s = re.sub(r'[^0-9]', '', s)
    return s.lstrip('0')

try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine='xlrd')
except Exception as e:
    print(f"Error: {e}")
    exit(1)

for name, df in sheets.items():
    df = df.fillna('')
    data = df.astype(str).values.tolist()
    header_row = None
    for i in range(min(40, len(data))):
        if any(re.search(r'قيد|الرقم|الطالب|اسم|وحدات|ساعات|مجموع', cell) for cell in data[i]):
            header_row = i
            break
    if header_row is None: header_row = 0
    headers = [x.strip() for x in data[header_row]]
    
    for r in range(header_row + 1, len(data)):
        row = data[r]
        if any(normalize_id(cell) == normalize_id(target_id) for cell in row):
            work_row = final_row = total_row = None
            for rr in range(r, min(len(data), r+8)):
                lbl = data[rr][4] if len(data[rr])>4 else ''
                if not work_row and re.search(r'اعمال|أعمال', lbl): work_row = rr
                if not final_row and re.search(r'امتحان|الأمتحان|اختبار', lbl): final_row = rr
                if not total_row and re.search(r'المجموع|مجموع', lbl): total_row = rr
            
            for col, h in enumerate(headers):
                w = data[work_row][col] if work_row else '?'
                f = data[final_row][col] if final_row else '?'
                t = data[total_row][col] if total_row else '?'
                print(f"Col {col}: Header='{h}' | Work='{w}' | Final='{f}' | Total='{t}'")
            break
