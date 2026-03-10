import os

def build_universal():
    with open('vehicle_monitor_v1.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. format_plate_with_data
    old_format = '''def format_plate_with_data(plate, exp_date, sheet_name="Unknown", owner="Unknown", office="", engine="", chassis="", brand="", year="", cost="", acq_date="", phys_status="", alert=""):
    if pd.isna(exp_date) or str(exp_date).strip() == '':
        dt_str = "N/A"
    else:
        try:
            if not hasattr(exp_date, 'strftime'):
                exp_date_str = str(exp_date).replace('\\\\', '/')
                exp_date = pd.to_datetime(exp_date_str, dayfirst=True)
            dt_str = exp_date.strftime('%Y-%m-%d')
        except:
            dt_str = str(exp_date)
            
    return json.dumps({
        "plate": plate,
        "owner": owner,
        "date": dt_str,
        "sheet": sheet_name,
        "office": office,
        "engine": engine,
        "chassis": chassis,
        "brand": brand,
        "year": year,
        "cost": cost,
        "acq_date": acq_date,
        "status": phys_status,
        "alert": alert
    })'''

    new_format = '''def format_plate_with_data(row_dict):
    # Simply serialize the dictionary, translating dates where needed
    final_dict = {}
    for k, v in row_dict.items():
        if pd.isna(v): v = ""
        elif hasattr(v, 'strftime'): v = v.strftime('%Y-%m-%d')
        final_dict[str(k)] = str(v)
    return json.dumps(final_dict)'''
    
    content = content.replace(old_format, new_format)

    # 2. send_notification
    old_send = '''def send_notification(detailed_alerts, title="⚠ Vehicle Update Detected", is_auto=False):
    if not detailed_alerts:
        return
    gui_queue.put({'type': 'show', 'alerts': detailed_alerts, 'title': title, 'is_auto': is_auto})'''
    
    new_send = '''def send_notification(detailed_alerts, columns_list=[], title="⚠ Update Detected", is_auto=False):
    if not detailed_alerts:
        return
    gui_queue.put({'type': 'show', 'alerts': detailed_alerts, 'columns': columns_list, 'title': title, 'is_auto': is_auto})'''
    
    content = content.replace(old_send, new_send)

    # 3. check_queue 
    old_check = '''self.build_ui(msg['alerts'], msg['title'])'''
    new_check = '''self.build_ui(msg['alerts'], msg.get('columns', []), msg['title'])'''
    content = content.replace(old_check, new_check)

    old_theme_change = '''self.build_ui(self.last_alerts, self.last_title)'''
    new_theme_change = '''self.build_ui(self.last_alerts, getattr(self, "last_columns", []), self.last_title)'''
    content = content.replace(old_theme_change, new_theme_change)

    old_build_def = '''def build_ui(self, detailed_alerts, window_title):
        self.last_alerts = detailed_alerts
        self.last_title = window_title'''
    new_build_def = '''def build_ui(self, detailed_alerts, columns_list, window_title):
        self.last_alerts = detailed_alerts
        self.last_columns = columns_list
        self.last_title = window_title'''
    content = content.replace(old_build_def, new_build_def)

    # 4. find_header_row
    old_header = '''def find_header_row(excel_file_obj, sheet_name):
    """
    Scans the first 15 rows looking for "PLATE". 
    Returns the integer index of the row to use as the header.
    """
    try:
        df_test = pd.read_excel(excel_file_obj, nrows=15, header=None, sheet_name=sheet_name)
        for i, row in df_test.iterrows():
            if any(isinstance(v, str) and 'PLATE' in v.upper() for v in row.values):
                return i
    except:
        pass
    return 3 # fallback default'''

    new_header = '''def find_header_row(excel_file_obj, sheet_name):
    try:
        df_test = pd.read_excel(excel_file_obj, nrows=15, header=None, sheet_name=sheet_name)
        for i, row in df_test.iterrows():
            str_count = sum(1 for v in row.values if isinstance(v, str) and len(v) > 2)
            if str_count >= 3:
                return i
    except:
        pass
    return 0'''
    content = content.replace(old_header, new_header)

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(content)
        
if __name__ == '__main__':
    build_universal()
