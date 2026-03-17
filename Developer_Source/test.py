import vehicle_monitor
import io
import pandas as pd

with open('VehicleMonitoring1.xlsx', 'rb') as f:
    fb = io.BytesIO(f.read())

xl = pd.ExcelFile(fb, engine='openpyxl')
hr = vehicle_monitor.find_header_row(xl, 'MARCH 2025')
df = pd.read_excel(xl, header=hr, sheet_name='MARCH 2025')

df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ')
exp_col = [c for c in df.columns if 'REMINDER' in str(c).upper()][0]
status_col = [c for c in df.columns if 'REGISTERED' in str(c).upper()][0]
alert_col = [c for c in df.columns if 'ALERT' in str(c).upper() and 'SYSTEM' not in str(c).upper()][0]
plate_col = [c for c in df.columns if 'PLATE' in str(c).upper()][0]

rows_output = []
for index, row in df.iterrows():
    plate = str(row[plate_col]).strip()
    if pd.isna(plate) or str(plate).strip() == '' or str(plate).upper() == 'CRITERIA':
        if str(plate).upper() == 'CRITERIA':
            break
        continue

    exp_date = row[exp_col] if pd.notna(row[exp_col]) else None
    
    status = None
    if alert_col and pd.notna(row[alert_col]) and str(row[alert_col]).strip() != '':
        val = str(row[alert_col]).strip().upper()
        if 'EXPIRED' in val:
            status = 'EXPIRED (RED)'
        elif 'DAYS BEFORE' in val and 'NOTICE' not in val and '2-WEEK' not in val and '2 WEEK' not in val:
            status = 'DAYS BEFORE EXPIRY (ORANGE)'
        elif '2-WEEK' in val or '2 WEEK' in val or '15 TO' in val:
            status = 'DAYS BEFORE 2 WEEK NOTICE (YELLOW)'
        elif 'SUFFICIENT' in val or '30 DAYS' in val:
            status = 'SUFFICIENT TIME (GREEN)'
        elif 'INPUT' in val:
            status = 'PLEASE INPUT LAST REG (GRAY)'
        elif 'REGISTERED' in val or 'YES' in val:
            status = 'REGISTERED (BLUE)'

    if not status:
        status_override = None
        if pd.notna(row[status_col]):
            val = str(row[status_col]).strip().upper()
            if val in ['YES', 'REGISTERED']:
                status_override = 'REGISTERED'
        status = vehicle_monitor.get_expiration_status(exp_date, status_override)
        
    formatted_plate = vehicle_monitor.format_plate_with_date(plate, exp_date)
    rows_output.append(f"{plate} | {status} | formatted: {formatted_plate}")

for line in rows_output:
    print(line)
