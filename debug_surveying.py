import pandas as pd
import os
import re

xls_dir = r'c:\Users\Arkan\Desktop\HISTJ\نتائج_الطلاب\xls'
file_path = os.path.join(xls_dir, 'قسم المساحة25-26.xls')
target_id = 'SE241002'

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
    nrows = len(data)
    
    # 1. Detect header row (logic from find_student.py)
    header_row = None
    for i in range(min(40, nrows)):
        row = data[i]
        if any(re.search(r'قيد|الرقم|الطالب|اسم|وحدات|ساعات|مجموع', cell) for cell in row):
            header_row = i
            break
    if header_row is None: header_row = 0
    
    headers = [x.strip() for x in data[header_row]]
    print(f"Detected Header Row: {header_row + 2} (1-based)")
    
    # 2. Find student
    for r in range(header_row + 1, nrows):
        row = data[r]
        if any(normalize_id(cell) == normalize_id(target_id) for cell in row):
            print(f"Found Student {target_id} at Row {r+2}")
            
            # 3. Simulate extract_block_results
            label_re_work = re.compile(r'اعمال|أعمال', re.I)
            label_re_final = re.compile(r'امتحان|الأمتحان|اختبار', re.I)
            label_re_total = re.compile(r'المجموع|مجموع', re.I)
            
            work_row = final_row = total_row = None
            for rr in range(r, min(nrows, r+8)):
                lbl = data[rr][4] if len(data[rr])>4 else ''
                if work_row is None and label_re_work.search(lbl): work_row = rr
                if final_row is None and label_re_final.search(lbl): final_row = rr
                if total_row is None and label_re_total.search(lbl): total_row = rr
            
            print(f"Rows found - Work: {work_row+2 if work_row else 'None'}, Final: {final_row+2 if final_row else 'None'}, Total: {total_row+2 if total_row else 'None'}")
            
            # 4. Show values for ALL headers
            for col, h in enumerate(headers):
                w = data[work_row][col] if work_row else '?'
                f = data[final_row][col] if final_row else '?'
                t = data[total_row][col] if total_row else '?'
                skip = re.search(r'رقم تسلسل|اســــــم الطال|رقم القيد|التقييم|معدل|التقدي', h)
                print(f"Col {col}: Header='{h}' | Work='{w}' | Final='{f}' | Total='{t}' | Filtered={bool(skip)}")
            break
