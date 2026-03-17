import pandas as pd
import os

filepath = "VehicleMonitoring.xlsx"
if os.path.exists(filepath):
    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
        print(f"Sheets: {xl.sheet_names}")
        for sh in xl.sheet_names[:1]: # just check first sheet
            print(f"\n--- {sh} ---")
            df = pd.read_excel(xl, sheet_name=sh, nrows=5)
            print("Columns:")
            for c in df.columns:
                print(f"  {c} ({df[c].dtype})")
            print("\nHead:")
            print(df.head())
    except Exception as e:
        print(f"Error: {e}")
else:
    print("File not found")
