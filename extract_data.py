import pandas as pd
import json

file_path = "c:/Users/User/Projects/Profitability Index/수익성분석지수_품목분류2_24년~26년_20260410.xlsx"
try:
    xl = pd.ExcelFile(file_path)
    output = {"sheets": xl.sheet_names, "data": {}}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, nrows=10)
        output["data"][sheet] = df.fillna("").values.tolist()
    
    with open("c:/Users/User/Projects/Profitability Index/excel_preview.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("Success")
except Exception as e:
    print("Error:", e)
