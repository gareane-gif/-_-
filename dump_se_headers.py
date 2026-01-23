
import pandas as pd
path = r'xls\قسم المساحة25-26.xls'
df = pd.read_excel(path, sheet_name=3, header=None) # Sheet 3 corresponds to Row 15 finding
print(df.iloc[3].values.tolist())
