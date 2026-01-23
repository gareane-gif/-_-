import pandas as pd
import os
import xlrd

def find_name(name_to_find):
    xls_dir = 'xls'
    files = [f for f in os.listdir(xls_dir) if f.endswith('.xls') or f.endswith('.xlsx')]
    
    for file in files:
        file_path = os.path.join(xls_dir, file)
        try:
            rb = xlrd.open_workbook(file_path, formatting_info=False)
            for sheet in rb.sheets():
                for r in range(sheet.nrows):
                    row = sheet.row_values(r)
                    for val in row:
                        if name_to_find in str(val):
                            print(f"MATCH FOUND in {file}, Sheet: {sheet.name}, Row: {r+1}")
                            print(f"Row data: {row}")
        except Exception as e:
            print(f"Error reading {file}: {e}")

find_name("امال احمد")
