#!/usr/bin/env python
import os
import re
import sys
import argparse
import json

try:
    import pandas as pd
except Exception:
    print('Missing dependency pandas. Install with: pip install -r requirements.txt')
    raise

parser = argparse.ArgumentParser(description='Find student results')
parser.add_argument('studentId')
parser.add_argument('--file', '-f', help='Optional specific Excel file (relative to xls/ or absolute)')
parser.add_argument('--format', choices=['plain', 'json', 'csv'], default='plain', help='Output format')
parser.add_argument('--out', help='Optional output path to save results (JSON or CSV based on --format)')
args = parser.parse_args()

id_to_find = args.studentId
out_format = args.format
file_arg = args.file
out_path = args.out

def normalize_id(s):
    if s is None:
        return ''
    s = str(s).strip().upper()
    # If it contains letters, keep them
    if any(c.isalpha() for c in s):
        s = re.sub(r'[^A-Z0-9]', '', s)
    else:
        s = re.sub(r'[^0-9]', '', s)
    s = re.sub(r'^0+', '', s)
    return s


def header_group_results(headers, row):
    """Group columns by header words (fallback when block extraction fails)."""
    work_re = re.compile(r'اعمال|عمل|الاعمال', re.I)
    final_re = re.compile(r'اختبار|اختبارات|نهائي|نهائية|Final', re.I)
    total_re = re.compile(r'النتيجة|الدرجة\s*النهائية|الدرجة|مجموع|المجموع', re.I)
    matches = []
    for i in range(max(len(headers), len(row))):
        h = headers[i] if i < len(headers) else ''
        v = row[i] if i < len(row) else ''
        typ = None
        if work_re.search(h or ''):
            typ = 'work'
        elif final_re.search(h or ''):
            typ = 'final'
        elif total_re.search(h or ''):
            typ = 'total'
        if typ:
            subj = re.sub(r'[-:()\[\]/\\]', ' ', h or '')
            subj = work_re.sub('', subj)
            subj = final_re.sub('', subj)
            subj = total_re.sub('', subj)
            subj = re.sub(r'\s+', ' ', subj).strip()
            if not subj:
                subj = '[unknown subject]'
            matches.append((subj, typ, h, v))
    if not matches:
        return {}
    grouped = {}
    for subj, typ, h, v in matches:
        grouped.setdefault(subj, {})[typ] = v
    # normalize to dict of {subj: {work,final,total}}
    out = {}
    for subj, parts in grouped.items():
        out[subj] = {'work': parts.get('work', ''), 'final': parts.get('final', ''), 'total': parts.get('total', '')}
    return out

target = normalize_id(id_to_find)

xls_dir = os.path.join(os.path.dirname(__file__), 'xls')
if not os.path.isdir(xls_dir):
    print('xls directory not found:', xls_dir)
    sys.exit(1)

if file_arg:
    # allow either absolute path or path relative to workspace/xls
    if os.path.isabs(file_arg):
        files = [file_arg]
    else:
        files = [os.path.join(xls_dir, file_arg)]
else:
    files = [f for f in os.listdir(xls_dir) if f.lower().endswith(('.xls', '.xlsx'))]
if not files:
    print('No Excel files found in', xls_dir)
    sys.exit(0)

found = False
results_aggregate = []
for file in files:
    path = file if os.path.isabs(file) else os.path.join(xls_dir, file)
    ext = os.path.splitext(file)[1].lower()
    try:
        if ext == '.xls':
            sheets = pd.read_excel(path, sheet_name=None, engine='xlrd', dtype=str, header=None)
        else:
            sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl', dtype=str, header=None)
    except Exception as e:
        print('Could not open', file, '->', e)
        continue

    for sheet_name, df in sheets.items():
        # convert DataFrame to list-of-rows (strings), preserving empty cells as ''
        try:
            df = df.fillna('')
            data = df.astype(str).values.tolist()
        except Exception:
            # fallback: empty sheet
            data = []
        nrows = len(data)
        ncols = max((len(r) for r in data), default=0)

        header_row = None
        for i in range(min(40, nrows)):
            row = data[i]
            if any(re.search(r'قيد|الرقم|الطالب|اسم|وحدات|ساعات|مجموع', str(cell)) for cell in row if cell is not None):
                header_row = i
                break
        if header_row is None:
            header_row = 0

        headers_raw = [str(x).strip() for x in data[header_row]]
        # Forward-fill headers to handle merged cells
        headers = []
        last_h = ""
        for h in headers_raw:
            if h == "" or h.lower() == "nan":
                headers.append(last_h)
            else:
                headers.append(h)
                last_h = h

        # QUICK SCAN: search any cell for the normalized target (robust fallback)
        id_col = None
        quick_scan_hit = False
        if target != '':
            for r_scan in range(header_row + 1, nrows):
                row_scan = data[r_scan]
                for c_scan in range(len(row_scan)):
                    if normalize_id(row_scan[c_scan]) == target:
                        id_col = c_scan
                        r = r_scan
                        quick_scan_hit = True
                        break
                if quick_scan_hit:
                    break


        # find id column by header (legacy heuristic)
        if id_col is None:
            for idx, h in enumerate(headers):
                if re.search(r'قيد|القيد|الرقم|الطالب|تسلسل', h or ''):
                    id_col = idx
                    break

        # if not found, try to detect column by scanning for matching value
        if id_col is None:
            for r in range(header_row + 1, min(nrows, header_row + 300)):
                row = data[r]
                for c in range(min(len(row), 50)):
                    if normalize_id(row[c]) == target and target != '':
                        id_col = c
                        break
                if id_col is not None:
                    break

        if id_col is None:
            continue

        for r in range(header_row + 1, nrows):
            row = data[r]
            val = normalize_id(row[id_col] if id_col < len(row) else '')
            if val == target and target != '':
                # prepare structured result using block-aware extraction for accounting-like sheets
                def extract_block_results(data, r, headers):
                    nrows = len(data)
                    # look for label rows near r: 'أعمال السنة' / 'الأمتحان النهائي' / 'المجموع'
                    label_re_work = re.compile(r'اعمال|أعمال', re.I)
                    label_re_final = re.compile(r'امتحان|الأمتحان|اختبار', re.I)
                    label_re_total = re.compile(r'المجموع|مجموع', re.I)
                    label_re_units = re.compile(r'عدد\s*الوحدات', re.I)
                    work_row = final_row = total_row = units_row = None
                    # search downward from r for work/final/total
                    # (avoid looking up into previous student's block)
                    for rr in range(r, min(nrows, r+8)):
                        lbl = data[rr][4] if len(data[rr])>4 else ''
                        if units_row is None and label_re_units.search(str(lbl)):
                            units_row = rr
                        if work_row is None and label_re_work.search(str(lbl)):
                            work_row = rr
                        if final_row is None and label_re_final.search(str(lbl)):
                            final_row = rr
                        if total_row is None and label_re_total.search(str(lbl)):
                            total_row = rr
                    
                    # Fallback: if labels are missing/garbled, use fixed offsets
                    # Most departments use: Work: +1, Final: +2, Total: +3
                    if work_row is None and final_row is None and total_row is None:
                        # only apply if we have enough rows below
                        if r + 3 < nrows:
                            work_row = r + 1
                            final_row = r + 2
                            total_row = r + 3
                            if units_row is None:
                                units_row = r
                    # Ensure units_row is at least r (ID row) if still None
                    if units_row is None:
                        units_row = r
                    results = {}
                    for col, h in enumerate(headers):
                        subj = str(h).strip()
                        if not subj or re.search(r'رقم تسلسل|اســــــم الطال|رقم القيد|التقييم|معدل|التقدي', subj):
                            continue
                        work_val = data[work_row][col] if work_row is not None and col < len(data[work_row]) else ''
                        final_val = data[final_row][col] if final_row is not None and col < len(data[final_row]) else ''
                        total_val = data[total_row][col] if total_row is not None and col < len(data[total_row]) else ''
                        units_val = data[units_row][col] if units_row is not None and col < len(data[units_row]) else ''
                        
                        # Handle duplicate header names (merged groups)
                        final_subj = subj
                        if final_subj in results:
                            final_subj = f"{subj} (عمود {col+1})"
                        results[final_subj] = {'work': work_val or '', 'final': final_val or '', 'total': total_val or '', 'units': str(units_val).strip()}
                    return results

                results = extract_block_results(data, r, headers)
                # if block extraction returned empty or only empty values, fallback to header grouping
                if not results or all((not v['work'] and not v['final'] and not v['total']) for v in results.values()):
                    results = header_group_results(headers, row)

                # prune subjects with all-empty fields
                results = {s:vals for s,vals in results.items() if any(str(vals.get(k,'')).strip() for k in ('work','final','total'))}
                # remove subjects whose units are present and equal to zero
                def is_zero_units(v):
                    u = str(v.get('units','')).strip()
                    if u == '':
                        return False
                    try:
                        return float(re.sub(r'[^0-9.-]', '', u) or '0') == 0
                    except Exception:
                        return False
                results = {s:vals for s,vals in results.items() if not is_zero_units(vals)}

                out = {
                    'file': file,
                    'sheet': sheet_name,
                    'header_row_1based': header_row+1,
                    'data_row_1based': r+1,
                    'id_col_1based': id_col+1,
                    'id_value': row[id_col] if id_col < len(row) else '' ,
                    'subjects': results
                }
                # produce output string/structure
                if out_format == 'json':
                    out_str = json.dumps(out, ensure_ascii=False, indent=2)
                elif out_format == 'csv':
                    lines = ['subject,work,final,total']
                    for subj, vals in out['subjects'].items():
                        safe = lambda s: str(s).replace('"', '""')
                        lines.append(f'"{safe(subj)}","{safe(vals.get("work",""))}","{safe(vals.get("final",""))}","{safe(vals.get("total",""))}"')
                    out_str = '\n'.join(lines)
                else:
                    parts = []
                    parts.append('\n--- MATCH FOUND ---')
                    parts.append(f"File: {file}")
                    parts.append(f"Sheet: {sheet_name}")
                    parts.append(f"Header row (1-based): {header_row + 1}")
                    parts.append(f"Data row (1-based): {r + 1}")
                    parts.append(f"ID column (1-based): {id_col + 1}")
                    parts.append(f"ID value: {row[id_col] if id_col < len(row) else ''}")
                    parts.append('Results:')
                    for subj, vals in results.items():
                        parts.append(f"  {subj}: أعمال الفصل: {vals.get('work','-')} | الاختبار النهائي: {vals.get('final','-')} | الدرجة الكلية: {vals.get('total','-')}")
                    out_str = '\n'.join(parts)

                # print to stdout for immediate feedback
                print(out_str)

                # collect result for aggregated output at end
                results_aggregate.append(out)
                found = True

if not found:
    print('No matches found for', id_to_find)
else:
    print('\nSearch completed.')

# write aggregated output once if requested
if out_path and results_aggregate:
    try:
        if out_format == 'json':
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(json.dumps(results_aggregate, ensure_ascii=False, indent=2))
        elif out_format == 'csv':
            # CSV columns: file,sheet,header_row,data_row,id_col,id_value,subject,work,final,total
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write('file,sheet,header_row_1based,data_row_1based,id_col_1based,id_value,subject,work,final,total\n')
                for out in results_aggregate:
                    base = [out.get('file',''), out.get('sheet',''), str(out.get('header_row_1based','')), str(out.get('data_row_1based','')), str(out.get('id_col_1based','')), str(out.get('id_value',''))]
                    for subj, vals in out.get('subjects', {}).items():
                        def esc(s):
                            return '"' + str(s).replace('"','""') + '"'
                        row = base + [subj, vals.get('work',''), vals.get('final',''), vals.get('total','')]
                        row = [esc(x) for x in row]
                        f.write(','.join(row) + '\n')
        else:
            # plain format: write concatenated plain outputs
            with open(out_path, 'w', encoding='utf-8') as f:
                for out in results_aggregate:
                    f.write(f"--- {out.get('file','')} | {out.get('sheet','')} | id={out.get('id_value','')} ---\n")
                    for subj, vals in out.get('subjects', {}).items():
                        f.write(f"{subj}: أعمال={vals.get('work','-')} | اختبار={vals.get('final','-')} | كلية={vals.get('total','-')}\n")
                    f.write('\n')
        print(f"Saved aggregated output to {out_path}")
    except Exception as e:
        print('Failed to write aggregated output file:', e)
