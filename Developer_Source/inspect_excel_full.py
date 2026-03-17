import pandas as pd
import os
import io

def find_header_row(excel_file_obj, sheet_name):
    try:
        df_test = pd.read_excel(excel_file_obj, nrows=15, header=None, sheet_name=sheet_name)
        for i, row in df_test.iterrows():
            if any(isinstance(v, str) and "PLATE" in v.upper() for v in row.values):
                return i
    except:
        pass
    return 3

filepath = "VehicleMonitoring.xlsx"
if os.path.exists(filepath):
    with open(filepath, "rb") as f:
        file_buffer = io.BytesIO(f.read())
    with pd.ExcelFile(file_buffer, engine="openpyxl") as xl:
        for sh in xl.sheet_names:
            h_row = find_header_row(xl, sh)
            df = pd.read_excel(xl, header=h_row, sheet_name=sh)
            df.columns = df.columns.astype(str).str.strip().str.replace("\n", " ")
            
            exp_cols = [c for c in df.columns if "REMINDER" in str(c).upper() or "EXPIRATION" in str(c).upper()]
            plate_cols = [c for c in df.columns if "PLATE" in str(c).upper()]
            
            print(f"\n--- Sheet: {sh} (Header Row: {h_row}) ---")
            if plate_cols:
                print(f"Plate Col: {plate_cols[0]}")
            if exp_cols:
                col = exp_cols[0]
                print(f"Exp Col: {col} (Type: {df[col].dtype})")
                # Print non-datetime values or unique values to inspect
                non_dt = df[col][df[col].notna()]
                print(f"Sample Values: {non_dt.unique()[:5]}")
            else:
                print("No Expiration Column found with 'REMINDER' or 'EXPIRATION'")
else:
    print("File not found")
