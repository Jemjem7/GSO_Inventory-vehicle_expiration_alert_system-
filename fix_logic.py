import os
import re

def execute_repair():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # The AI extraction patch made it so `master_columns` is sometimes empty if the first processed file is empty or fails.
    # It also broke the `format_plate_with_data` JSON serialization by trying to convert all datetime objects.
    # Let's fix process_excel's combined arrays and the JSON serializer.

    old_format = """def format_plate_with_data(row_dict):
    # Simply serialize the dictionary, translating dates where needed
    final_dict = {}
    for k, v in row_dict.items():
        if pd.isna(v): v = ""
        elif hasattr(v, 'strftime'): v = v.strftime('%Y-%m-%d')
        final_dict[str(k)] = str(v)
    return json.dumps(final_dict)"""

    new_format = """def format_plate_with_data(row_dict):
    final_dict = {}
    for k, v in row_dict.items():
        try:
            if pd.isna(v): v = ""
            elif hasattr(v, 'strftime'): v = v.strftime('%Y-%m-%d')
            final_dict[str(k)] = str(v)
        except:
            final_dict[str(k)] = str(v)
    return json.dumps(final_dict)"""

    code = code.replace(old_format, new_format)

    # In process_excel, master_columns gets assigned naively: `master_columns = all_data[0][3] if all_data else []`
    # Let's ensure ALL columns from ALL sheets/files are discovered.
    
    old_master = """    master_columns = all_data[0][3] if all_data else []
    
    for c_state, c_records, s_name, cols in all_data:
        combined_current_state.update(c_state)
        combined_changed_records.extend(c_records)"""

    new_master = """    master_columns = []
    
    for c_state, c_records, s_name, cols in all_data:
        combined_current_state.update(c_state)
        combined_changed_records.extend(c_records)
        for c in cols:
            if c not in master_columns and str(c).upper() != "NAN":
                master_columns.append(c)"""

    code = code.replace(old_master, new_master)

    # In process_excel, the initial alert loop might be firing empty if there's no data.
    # We must ensure `send_notification` ALWAYS fires on manual scan so the UI pops up.
    
    old_initial = """        if initial_alerts: send_notification(initial_alerts, master_columns, title=f"⚠ Initial Scan Results: {os.path.basename(filepath)}", is_auto=True)
        else: send_notification({"SUFFICIENT TIME": ["All Records clear"]}, master_columns, title="⚠ Initial Scan Results", is_auto=True)"""

    new_initial = """        if initial_alerts: send_notification(initial_alerts, master_columns, title=f"⚠ Initial Scan Results", is_auto=True)
        else: send_notification({"SUFFICIENT TIME": ['{"_sheet": "N/A", "_status": "SUFFICIENT TIME (GREEN)"}']}, master_columns, title="⚠ Initial Scan Results", is_auto=True)"""

    code = code.replace(old_initial, new_initial)

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    execute_repair()
