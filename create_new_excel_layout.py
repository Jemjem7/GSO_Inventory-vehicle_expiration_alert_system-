import pandas as pd
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

EXCEL_FILE = "VehicleMonitoring.xlsx"
NEW_EXCEL_FILE = "VehicleMonitoring_NewLayout.xlsx"

def find_header_row(xl, sheet_name):
    try:
        df_test = xl.parse(sheet_name, nrows=10)
        for i in range(len(df_test)):
            row_vals = [str(x).upper() for x in df_test.iloc[i].values]
            if any('PLATE' in v for v in row_vals) and any(('NAME' in v or 'OWNER' in v or 'PERSON' in v or 'ACCOUNTABLE' in v) for v in row_vals):
                return i + 1
        return 0
    except:
        return 0

def get_expiration_status(exp_date, status_override=None):
    if status_override:
        return status_override
    
    if pd.isna(exp_date):
        return 'No Expiration Date'

    try:
        if isinstance(exp_date, str):
            exp_date = pd.to_datetime(exp_date).date()
        else:
            exp_date = exp_date.date()
            
        today = datetime.now().date()
        days_until = (exp_date - today).days

        if days_until < 0:
            return 'EXPIRED (RED)'
        elif days_until <= 7:
            return '1 WEEK BEFORE EXPIRY (RED)'
        elif days_until <= 30:
            return '1 MONTH BEFORE EXPIRY (ORANGE)'
        elif days_until <= 60:
            return '2 MONTHS BEFORE EXPIRY (YELLOW)'
        else:
            return 'SUFFICIENT TIME (GREEN)'
    except:
        return 'Invalid Date'

def main():
    print(f"Reading {EXCEL_FILE}...")
    try:
        xl = pd.ExcelFile(EXCEL_FILE)
    except Exception as e:
        print(f"Error opening {EXCEL_FILE}: {e}")
        return

    writer = pd.ExcelWriter(NEW_EXCEL_FILE, engine='xlsxwriter')
    
    for sh in xl.sheet_names:
        print(f"Processing sheet: {sh}")
        h_row = find_header_row(xl, sh)
        df_sheet = pd.read_excel(xl, header=h_row, sheet_name=sh)
        
        if df_sheet.empty:
            continue
            
        df_sheet.columns = df_sheet.columns.astype(str).str.strip().str.replace('\n', ' ')

        # Find the columns
        plate_col = next((c for c in df_sheet.columns if 'PLATE' in str(c).upper()), 'PLATE #')
        owner_col = next((c for c in df_sheet.columns if 'NAME' in str(c).upper() or 'OWNER' in str(c).upper() or 'CUSTOMER' in str(c).upper() or 'ACCOUNTABLE' in str(c).upper() or 'PERSON' in str(c).upper()), None)
        exp_col = next((c for c in df_sheet.columns if 'REMINDER' in str(c).upper() or 'EXPIRATION' in str(c).upper() or 'EXPIRY' in str(c).upper() or ('DATE' in str(c).upper() and 'ACQUISITION' not in str(c).upper())), 'REMINDER')
        status_col = next((c for c in df_sheet.columns if 'REGISTERED' in str(c).upper()), None)
        phys_status_col = next((c for c in df_sheet.columns if 'STATUS' in str(c).upper() and 'NOT' not in str(c).upper()), None)
        alert_col = next((c for c in df_sheet.columns if 'ALERT' in str(c).upper() and 'SYSTEM' not in str(c).upper()), None)
        office_col = next((c for c in df_sheet.columns if 'OFFICE' in str(c).upper()), None)
        engine_col = next((c for c in df_sheet.columns if 'ENGINE' in str(c).upper()), None)
        chassis_col = next((c for c in df_sheet.columns if 'CHASSIS' in str(c).upper()), None)
        brand_col = next((c for c in df_sheet.columns if 'BRAND' in str(c).upper() or 'BODY TYPE' in str(c).upper()), None)
        year_col = next((c for c in df_sheet.columns if 'YEAR' in str(c).upper()), None)
        acq_date_col = next((c for c in df_sheet.columns if 'ACQUISITION DATE' in str(c).upper()), None)
        driver_col = next((c for c in df_sheet.columns if 'DRIVER' in str(c).upper()), None)

        if plate_col not in df_sheet.columns:
            continue

        new_data = []

        for index, row in df_sheet.iterrows():
            plate = row[plate_col]
            if pd.isna(plate) or str(plate).strip() == '' or str(plate).upper() == 'CRITERIA':
                if str(plate).upper() == 'CRITERIA': break
                continue
                
            plate = str(plate).strip()
            owner = str(row[owner_col]).strip() if owner_col and pd.notna(row[owner_col]) else ""
            val_driver = str(row[driver_col]).strip() if driver_col and pd.notna(row[driver_col]) else ""
            val_office = str(row[office_col]).strip() if office_col and pd.notna(row[office_col]) else ""
            val_engine = str(row[engine_col]).strip() if engine_col and pd.notna(row[engine_col]) else ""
            val_chassis = str(row[chassis_col]).strip() if chassis_col and pd.notna(row[chassis_col]) else ""
            val_brand = str(row[brand_col]).strip() if brand_col and pd.notna(row[brand_col]) else ""
            val_year = str(row[year_col]).strip() if year_col and pd.notna(row[year_col]) else ""
            if val_year and val_year.endswith(".0"): val_year = val_year[:-2]
            
            acq_d = row[acq_date_col] if acq_date_col and pd.notna(row[acq_date_col]) else ""
            val_acq_date = ""
            if acq_d != "":
                try:
                    if hasattr(acq_d, 'strftime'): val_acq_date = acq_d.strftime('%Y-%m-%d')
                    else: val_acq_date = str(acq_d).split(" ")[0]
                except:
                    val_acq_date = str(acq_d)
                    
            val_phys_status = str(row[phys_status_col]).strip() if phys_status_col and pd.notna(row[phys_status_col]) else ""
            exp_date = row[exp_col] if exp_col in df_sheet.columns and pd.notna(row[exp_col]) else ""
            
            if exp_date != "" and hasattr(exp_date, 'strftime'):
                exp_date_str = exp_date.strftime('%Y-%m-%d')
            else:
                exp_date_str = str(exp_date)

            status = None
            if alert_col and pd.notna(row[alert_col]) and str(row[alert_col]).strip() != '':
                val = str(row[alert_col]).strip().upper()
                if 'EXPIRED' in val or 'LESS THAN' in val: status = 'EXPIRED (RED)'
                elif '1 WEEK' in val or '1-WEEK' in val or ('WEEK' in val and '1' in val) or '1 TO 7' in val or '1-7' in val: status = '1 WEEK BEFORE EXPIRY (RED)'
                elif '1 MONTH' in val or '1-MONTH' in val or 'WEEK' in val or '8 TO 30' in val or '8-30' in val or '30 DAYS' in val: status = '1 MONTH BEFORE EXPIRY (ORANGE)'
                elif '2 MONTH' in val or '2-MONTH' in val or '60 DAYS' in val or '31 TO 60' in val or '31-60' in val: status = '2 MONTHS BEFORE EXPIRY (YELLOW)'
                elif 'SUFFICIENT' in val or 'MORE' in val: status = 'SUFFICIENT TIME (GREEN)'
                elif 'INPUT' in val: status = 'PLEASE INPUT LAST REG (GRAY)'
                elif 'REGISTERED' in val or 'YES' in val: status = 'REGISTERED (BLUE)'

            if not status:
                status_override = None
                if status_col and pd.notna(row[status_col]):
                    val = str(row[status_col]).strip().upper()
                    if val in ['YES', 'REGISTERED']: status_override = 'REGISTERED'
                status = get_expiration_status(exp_date if exp_date != "" else None, status_override)

            new_data.append({
                "OFFICE": val_office,
                "PLATE NUMBER": plate,
                "ENGINE NUMBER": val_engine,
                "CHASSIS NO.": val_chassis,
                "BRAND/ BODY TYPE": val_brand,
                "YEAR MODEL": val_year,
                "EXPIRATION DATE": exp_date_str,
                "ACQUISITION DATE": val_acq_date,
                "ACCOUNTABLE PERSON": owner,
                "DRIVER": val_driver,
                "STATUS": val_phys_status,
                "ALERT": status
            })

        if new_data:
            df_new = pd.DataFrame(new_data)
            df_new.to_excel(writer, sheet_name=sh, index=False)
            
            # Auto-adjust columns width
            worksheet = writer.sheets[sh]
            for i, col in enumerate(df_new.columns):
                column_len = max(df_new[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)
    
    writer.close()
    print(f"Successfully generated {NEW_EXCEL_FILE}")

if __name__ == "__main__":
    main()
