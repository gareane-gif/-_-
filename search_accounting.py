#!/usr/bin/env python
import os
import re
import sys

try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

def normalize_id(s):
    if s is None:
        return ''
    s = str(s).strip()
    s = re.sub(r'[^0-9]', '', s)
    s = re.sub(r'^0+', '', s)
    return s

if len(sys.argv) < 2:
    print('Usage: python search_accounting.py <studentId>')
    sys.exit(1)

id_to_find = sys.argv[1]
target = normalize_id(id_to_find)

base = os.path.dirname(__file__)
path = os.path.join(base, 'xls', 'قسم المحاسبة25-26.xls')
if not os.path.exists(path):
    print('Accounting file not found:', path)
    sys.exit(1)

ext = os.path.splitext(path)[1].lower()
try:
    if ext == '.xls':
        sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', dtype=str)
    else:
        sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl', dtype=str)
except Exception as e:
    print('Failed to open file:', e)
    sys.exit(1)

found = False
for sheet_name, df in sheets.items():
    df = df.fillna('')
    data = df.astype(str).values.tolist()
    nrows = len(data)

    # search all cells for the normalized id
    for r in range(nrows):
        row = data[r]
        for c in range(len(row)):
            if normalize_id(row[c]) == target and target != '':
                print('\n--- MATCH ---')
                print('Sheet:', sheet_name)
                print('Raw ID cell:', row[c])
                # detect header row (look in first 10 rows for Arabic header keywords)
                header_row = None
                for hi in range(min(10, nrows)):
                    crow = data[hi]
                    if any(re.search(r'رقم تسلسل|رقم القيد|اســــــم الطال|مادة|مبادئ|مجموع|المجموع|اسم', str(cell)) for cell in crow if cell is not None):
                        header_row = hi
                        break
                if header_row is None:
                    header_row = 0

                headers = [str(x).strip() for x in data[header_row]] if nrows>0 else []

                # For this workbook each student block spans multiple nearby rows: 'عدد الوحدات', then 'أعمال السنة', 'الأمتحان النهائي', 'المجموع'.
                base_row = r
                # locate rows with those labels within the next 6 rows
                work_row = None
                final_row = None
                total_row = None
                label_re_work = re.compile(r'اعمال|أعمال', re.I)
                label_re_final = re.compile(r'امتحان|الأمتحان|اختبار', re.I)
                label_re_total = re.compile(r'المجموع|مجموع', re.I)
                for rr in range(base_row, min(nrows, base_row+7)):
                    crow = data[rr]
                    # check column 5 (index 4) which typically contains the row label
                    lbl = ''
                    if len(crow) > 4:
                        lbl = str(crow[4])
                    if label_re_work.search(lbl) and work_row is None:
                        work_row = rr
                    if label_re_final.search(lbl) and final_row is None:
                        final_row = rr
                    if label_re_total.search(lbl) and total_row is None:
                        total_row = rr

                # build per-subject output using header columns
                print('النتائج:')
                for col in range(len(headers)):
                    subj = headers[col]
                    if subj is None:
                        continue
                    subj = subj.strip()
                    if not subj:
                        continue
                    # read values from located rows
                    work_val = data[work_row][col] if work_row is not None and col < len(data[work_row]) else ''
                    final_val = data[final_row][col] if final_row is not None and col < len(data[final_row]) else ''
                    total_val = data[total_row][col] if total_row is not None and col < len(data[total_row]) else ''
                    # skip non-subject columns like 'رقم تسلسل' 'اسم الطالب' etc
                    if re.search(r'رقم تسلسل|اســــــم الطال|رقم القيد|التقييم|عدد الوحدات|معدل|التقدي', subj):
                        continue
                    print(f'  {subj}: أعمال الفصل: {work_val or "-"} | الاختبار النهائي: {final_val or "-"} | الدرجة الكلية: {total_val or "-"}')
                found = True

if not found:
    print('No matches found in accounting file for', id_to_find)
else:
    print('\nDone.')
