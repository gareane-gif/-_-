
import pandas as pd
import json

path = r'xls\قسم المساحة25-26.xls'
df = pd.read_excel(path, sheet_name=3, header=None)
rows = df.iloc[:15].astype(str).values.tolist()
print(json.dumps(rows, ensure_ascii=False, indent=2))
