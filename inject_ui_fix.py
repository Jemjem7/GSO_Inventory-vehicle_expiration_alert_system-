import os
import re

def execute_final_ui_fix():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # The dynamic logic we injected before got swallowed on some fallback loops.
    # The actual crash is on `tree.insert("", tk.END, values=(office, plate, engine...))`
    # It must be purely dynamic: `values=tuple(row_values)`

    old_row_loop = """                    office = data.get("office", "")
                    plate = data.get("plate", "Unknown")
                    engine = data.get("engine", "")
                    chassis = data.get("chassis", "")
                    brand = data.get("brand", "")
                    year = data.get("year", "")
                    date_val = data.get("date", "N/A")
                    cost = data.get("cost", "")
                    acq_date = data.get("acq_date", "")
                    owner = data.get("owner", "Unknown")
                    phys_status = data.get("status", "")
                    alert_val = data.get("alert", status_key)
                    sheet_name = data.get("sheet", "Unknown")
                        
                    # Insert row
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    try:
                        tree.insert("", tk.END, values=(office, plate, engine, chassis, brand, year, date_val, cost, acq_date, owner, phys_status, alert_val, sheet_name, sheet_name), tags=(status_key, stripe_tag))
                    except:
                        tree.insert("", tk.END, values=(office, plate, engine, chassis, brand, year, date_val, cost, acq_date, owner, phys_status, alert_val, sheet_name, sheet_name), tags=(status_key, stripe_tag))
                    row_count += 1
                    has_alerts = True"""

    # If `old_row_loop` string match fails, let's use exact regex
    pattern = re.compile(r'office = data\.get\("office", ""\).*?has_alerts = True', re.DOTALL)
    
    new_row_loop = """row_values = []
                    # Fallback columns if none passed
                    active_cols = columns if columns else ["office", "plate", "engine", "chassis", "brand", "year", "date", "cost", "acq", "owner", "status", "alert", "month", "sheet"]
                    
                    for col in active_cols:
                        row_values.append(data.get(col, ""))
                        
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    try:
                        tree.insert("", tk.END, values=tuple(row_values), tags=(status_key, stripe_tag))
                    except Exception as e:
                        print(f"DEBUG: TUPLE INSERTION CRASH: {e}")
                    row_count += 1
                    has_alerts = True"""
                    
    code = pattern.sub(new_row_loop, code)
    
    # Finally, make sure the `AlertWindow` physically pops up instantly when called via updating the queue listener.
    
    old_show = """                    self.build_ui(msg['alerts'], msg.get('columns', []), msg['title'])
                    self.deiconify()
                    self.lift()
                    self.attributes('-topmost', True)"""

    new_show = """                    self.build_ui(msg['alerts'], msg.get('columns', []), msg['title'])
                    if not self.winfo_ismapped():
                        self.deiconify()
                    self.state('normal')
                    self.lift()
                    self.attributes('-topmost', True)"""
                    
    code = code.replace(old_show, new_show)

    # Let's clean up all those debug prints!
    code = re.sub(r'print\(f?[\'"]DEBUG:.*?[\'"]\)', '', code)

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    execute_final_ui_fix()
