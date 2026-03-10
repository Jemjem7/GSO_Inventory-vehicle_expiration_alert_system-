import os

with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Update format_plate_with_data
old_format = '''def format_plate_with_data(plate, exp_date, sheet_name="Unknown", owner="Unknown", office="", engine="", chassis="", brand="", year="", cost="", acq_date="", phys_status="", alert="", driver=""):'''
new_format = '''def format_plate_with_data(plate, exp_date, sheet_name="Unknown", owner="Unknown", office="", make="", type_val="", emission="", gsis="", lto="", last_reg="", cost="", acq_date="", phys_status="", alert="", insurance="", driver=""):'''

old_format_dict = '''    return json.dumps({
        "plate": plate,
        "owner": owner,
        "driver": driver,
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
new_format_dict = '''    return json.dumps({
        "plate": plate,
        "owner": owner,
        "driver": driver,
        "date": dt_str,
        "sheet": sheet_name,
        "office": office,
        "make": make,
        "type": type_val,
        "emission": emission,
        "gsis": gsis,
        "lto": lto,
        "last_reg": last_reg,
        "insurance": insurance,
        "cost": cost,
        "acq_date": acq_date,
        "status": phys_status,
        "alert": alert
    })'''

text = text.replace(old_format, new_format)
text = text.replace(old_format_dict, new_format_dict)

# 2. Update Treeview Setup
old_tree_cols = '''        # Setup Treeview Table
        columns = ("office", "plate", "engine", "chassis", "brand", "year", "date", "cost", "acq", "owner", "driver", "status", "alert", "month", "sheet")
        tree = ttk.Treeview(summary_frame, columns=columns, show="headings", style="Custom.Treeview", height=15)
        
        tree.heading("office", text="OFFICE", anchor=tk.W)
        tree.heading("plate", text="PLATE NUMBER", anchor=tk.W)
        tree.heading("engine", text="ENGINE NUMBER", anchor=tk.W)
        tree.heading("chassis", text="CHASSIS NO.", anchor=tk.W)
        tree.heading("brand", text="BRAND/ BODY TYPE", anchor=tk.W)
        tree.heading("year", text="YEAR MODEL", anchor=tk.W)
        tree.heading("date", text="EXPIRATION DATE", anchor=tk.W)
        tree.heading("cost", text="ACQUISITION COST", anchor=tk.W)
        tree.heading("acq", text="ACQUISITION DATE", anchor=tk.W)
        tree.heading("owner", text="ACCOUNTABLE PERSON", anchor=tk.W)
        tree.heading("driver", text="DRIVER", anchor=tk.W)
        tree.heading("status", text="STATUS", anchor=tk.W)
        tree.heading("alert", text="ALERT", anchor=tk.W)
        tree.heading("month", text="MONTH", anchor=tk.W)
        tree.heading("sheet", text="Sheet", anchor=tk.W) 
        
        tree.column("office", width=80, minwidth=60, stretch=tk.NO)
        tree.column("plate", width=120, minwidth=100, stretch=tk.NO)
        tree.column("engine", width=140, minwidth=100, stretch=tk.YES)
        tree.column("chassis", width=140, minwidth=100, stretch=tk.YES)
        tree.column("brand", width=140, minwidth=100, stretch=tk.YES)
        tree.column("year", width=80, minwidth=50, stretch=tk.NO)
        tree.column("date", width=120, minwidth=100, stretch=tk.NO)
        tree.column("cost", width=120, minwidth=70, stretch=tk.NO)
        tree.column("acq", width=120, minwidth=90, stretch=tk.NO)
        tree.column("owner", width=160, minwidth=110, stretch=tk.YES)
        tree.column("driver", width=140, minwidth=100, stretch=tk.YES)
        tree.column("status", width=100, minwidth=80, stretch=tk.NO)
        tree.column("alert", width=180, minwidth=120, stretch=tk.YES)
        tree.column("month", width=100, minwidth=80, stretch=tk.NO)
        tree.column("sheet", width=0, minwidth=0, stretch=tk.NO)'''

new_tree_cols = '''        # Setup Treeview Table
        columns = ("phys_status", "office", "plate", "make", "type", "emission", "gsis", "lto", "last_reg", "reminder", "alert", "insurance", "driver", "cost", "acq", "month", "sheet")
        tree = ttk.Treeview(summary_frame, columns=columns, show="headings", style="Custom.Treeview", height=15)
        
        tree.heading("phys_status", text="STATUS (YES/NO)", anchor=tk.W)
        tree.heading("office", text="OFFICE", anchor=tk.W)
        tree.heading("plate", text="PLATE #", anchor=tk.W)
        tree.heading("make", text="MAKE", anchor=tk.W)
        tree.heading("type", text="TYPE", anchor=tk.W)
        tree.heading("emission", text="EMISSION", anchor=tk.W)
        tree.heading("gsis", text="GSIS", anchor=tk.W)
        tree.heading("lto", text="LTO", anchor=tk.W)
        tree.heading("last_reg", text="LAST REG.", anchor=tk.W)
        tree.heading("reminder", text="REMINDER", anchor=tk.W)
        tree.heading("alert", text="ALERT", anchor=tk.W)
        tree.heading("insurance", text="INSURANCE", anchor=tk.W)
        tree.heading("driver", text="DRIVER", anchor=tk.W)
        tree.heading("cost", text="ACQUISITION COST", anchor=tk.W)
        tree.heading("acq", text="DATE ACQUIRED", anchor=tk.W)
        tree.heading("month", text="MONTH", anchor=tk.W)
        tree.heading("sheet", text="Sheet", anchor=tk.W) 
        
        tree.column("phys_status", width=120, minwidth=80, stretch=tk.NO)
        tree.column("office", width=80, minwidth=60, stretch=tk.NO)
        tree.column("plate", width=100, minwidth=80, stretch=tk.NO)
        tree.column("make", width=120, minwidth=80, stretch=tk.YES)
        tree.column("type", width=120, minwidth=80, stretch=tk.YES)
        tree.column("emission", width=100, minwidth=80, stretch=tk.YES)
        tree.column("gsis", width=100, minwidth=80, stretch=tk.YES)
        tree.column("lto", width=100, minwidth=80, stretch=tk.YES)
        tree.column("last_reg", width=100, minwidth=80, stretch=tk.NO)
        tree.column("reminder", width=100, minwidth=80, stretch=tk.NO)
        tree.column("alert", width=160, minwidth=100, stretch=tk.YES)
        tree.column("insurance", width=100, minwidth=80, stretch=tk.NO)
        tree.column("driver", width=120, minwidth=100, stretch=tk.YES)
        tree.column("cost", width=120, minwidth=80, stretch=tk.NO)
        tree.column("acq", width=100, minwidth=80, stretch=tk.NO)
        tree.column("month", width=100, minwidth=80, stretch=tk.NO)
        tree.column("sheet", width=0, minwidth=0, stretch=tk.NO)'''

text = text.replace(old_tree_cols, new_tree_cols)

# 3. Update JSON extraction for Tree items
old_tree_vars = '''                    office = data.get("office", "")
                    plate = data.get("plate", "Unknown")
                    engine = data.get("engine", "")
                    chassis = data.get("chassis", "")
                    brand = data.get("brand", "")
                    year = data.get("year", "")
                    date_val = data.get("date", "N/A")
                    cost = data.get("cost", "")
                    acq_date = data.get("acq_date", "")
                    owner = data.get("owner", "Unknown")
                    driver = data.get("driver", "Unknown")
                    phys_status = data.get("status", "")
                    alert_val = data.get("alert", status_key)
                    sheet_name = data.get("sheet", "Unknown")
                        
                    # Insert row
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    tree.insert("", tk.END, values=(office, plate, engine, chassis, brand, year, date_val, cost, acq_date, owner, driver, phys_status, alert_val, sheet_name, sheet_name), tags=(status_key, stripe_tag))'''
new_tree_vars = '''                    office = data.get("office", "")
                    plate = data.get("plate", "Unknown")
                    make = data.get("make", "")
                    type_val = data.get("type", "")
                    emission = data.get("emission", "")
                    gsis = data.get("gsis", "")
                    lto = data.get("lto", "")
                    last_reg = data.get("last_reg", "")
                    reminder = data.get("date", "N/A")
                    alert_val = data.get("alert", status_key)
                    insurance = data.get("insurance", "")
                    driver = data.get("driver", "Unknown")
                    cost = data.get("cost", "")
                    acq_date = data.get("acq_date", "")
                    phys_status = data.get("status", "")
                    sheet_name = data.get("sheet", "Unknown")
                        
                    # Insert row
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    tree.insert("", tk.END, values=(phys_status, office, plate, make, type_val, emission, gsis, lto, last_reg, reminder, alert_val, insurance, driver, cost, acq_date, sheet_name, sheet_name), tags=(status_key, stripe_tag))'''
text = text.replace(old_tree_vars, new_tree_vars)

# 4. Sheet column parsing replacing chunk
import re

parse_start = """        office_c = [c for c in df_sheet.columns if 'OFFICE' in str(c).upper()]"""
parse_end = """                   val_acq_date = str(acq_d)"""

# Instead of complex regex matching that might fail, lets use string slice
idx1 = text.find(parse_start)
idx2 = text.find(parse_end) + len(parse_end)
if idx1 != -1 and idx2 != -1:
    old_parsing = text[idx1:idx2]
    new_parsing = """        office_c = [c for c in df_sheet.columns if 'OFFICE' in str(c).upper()]
        make_c = [c for c in df_sheet.columns if 'MAKE' in str(c).upper()]
        type_c = [c for c in df_sheet.columns if 'TYPE' in str(c).upper() and 'BODY' not in str(c).upper()]
        emission_c = [c for c in df_sheet.columns if 'EMISSION' in str(c).upper()]
        gsis_c = [c for c in df_sheet.columns if 'GSIS' in str(c).upper()]
        lto_c = [c for c in df_sheet.columns if 'LTO' in str(c).upper()]
        last_reg_c = [c for c in df_sheet.columns if 'LAST REG' in str(c).upper()]
        insurance_c = [c for c in df_sheet.columns if 'INSURANCE' in str(c).upper()]
        cost_c = [c for c in df_sheet.columns if 'COST' in str(c).upper()]
        acq_date_c = [c for c in df_sheet.columns if 'ACQUIRED' in str(c).upper() or 'ACQUISITION DATE' in str(c).upper()]
        driver_c = [c for c in df_sheet.columns if 'DRIVER' in str(c).upper()]
        
        office_col = office_c[0] if office_c else None
        make_col = make_c[0] if make_c else None
        type_col = type_c[0] if type_c else None
        emission_col = emission_c[0] if emission_c else None
        gsis_col = gsis_c[0] if gsis_c else None
        lto_col = lto_c[0] if lto_c else None
        last_reg_col = last_reg_c[0] if last_reg_c else None
        insurance_col = insurance_c[0] if insurance_c else None
        cost_col = cost_c[0] if cost_c else None
        acq_date_col = acq_date_c[0] if acq_date_c else None
        driver_col = driver_c[0] if driver_c else None

        if plate_col not in df_sheet.columns:
            continue
            
        current_state = {}
        changed_records = []
        
        for index, row in df_sheet.iterrows():
            plate = row[plate_col]
            owner = str(row[owner_col]).strip() if owner_col and pd.notna(row[owner_col]) else "Unknown"
            val_driver = str(row[driver_col]).strip() if driver_col and pd.notna(row[driver_col]) else ""
            
            val_office = str(row[office_col]).strip() if office_col and pd.notna(row[office_col]) else ""
            val_make = str(row[make_col]).strip() if make_col and pd.notna(row[make_col]) else ""
            val_type = str(row[type_col]).strip() if type_col and pd.notna(row[type_col]) else ""
            val_emission = str(row[emission_col]).strip() if emission_col and pd.notna(row[emission_col]) else ""
            val_gsis = str(row[gsis_col]).strip() if gsis_col and pd.notna(row[gsis_col]) else ""
            val_lto = str(row[lto_col]).strip() if lto_col and pd.notna(row[lto_col]) else ""
            val_last_reg = str(row[last_reg_col]).strip() if last_reg_col and pd.notna(row[last_reg_col]) else ""
            val_insurance = str(row[insurance_col]).strip() if insurance_col and pd.notna(row[insurance_col]) else ""
            val_cost = str(row[cost_col]).strip() if cost_col and pd.notna(row[cost_col]) else ""
            
            acq_d = row[acq_date_col] if acq_date_col and pd.notna(row[acq_date_col]) else ""
            val_acq_date = ""
            if acq_d != "":
                try:
                    if hasattr(acq_d, 'strftime'): val_acq_date = acq_d.strftime('%Y-%m-%d')
                    else: val_acq_date = str(acq_d).split(" ")[0]
                except:
                   val_acq_date = str(acq_d)"""
                   
    text = text.replace(old_parsing, new_parsing)


# 5. Tuple creation unpacking 
# Re-do phys_status_col search 
t_find = "phys_status_keys = [c for c in df_sheet.columns if 'STATUS' in str(c).upper() and 'NOT' not in str(c).upper()]"
t_rep = "phys_status_keys = [c for c in df_sheet.columns if 'YES' in str(c).upper() and 'NOT' not in str(c).upper()]"
text = text.replace(t_find, t_rep)

t1 = r"current_state[plate] = (status, exp_date, sheet_name, owner, val_office, val_engine, val_chassis, val_brand, val_year, val_cost, val_acq_date, val_phys_status, val_driver)"
t2 = r"current_state[plate] = (status, exp_date, sheet_name, owner, val_office, val_make, val_type, val_emission, val_gsis, val_lto, val_last_reg, val_insurance, val_cost, val_acq_date, val_phys_status, val_driver)"
text = text.replace(t1, t2)

# Unpacking loops
loop_old = '''                 status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
                 owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
                 office = state_tuple[4] if len(state_tuple) > 4 else ""
                 engine = state_tuple[5] if len(state_tuple) > 5 else ""
                 chassis = state_tuple[6] if len(state_tuple) > 6 else ""
                 brand = state_tuple[7] if len(state_tuple) > 7 else ""
                 year = state_tuple[8] if len(state_tuple) > 8 else ""
                 cost = state_tuple[9] if len(state_tuple) > 9 else ""
                 acq_date = state_tuple[10] if len(state_tuple) > 10 else ""
                 phys_status = state_tuple[11] if len(state_tuple) > 11 else ""
                 driver = state_tuple[12] if len(state_tuple) > 12 else ""'''

loop_new = '''                 status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
                 owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
                 office = state_tuple[4] if len(state_tuple) > 4 else ""
                 make = state_tuple[5] if len(state_tuple) > 5 else ""
                 type_val = state_tuple[6] if len(state_tuple) > 6 else ""
                 emission = state_tuple[7] if len(state_tuple) > 7 else ""
                 gsis = state_tuple[8] if len(state_tuple) > 8 else ""
                 lto = state_tuple[9] if len(state_tuple) > 9 else ""
                 last_reg = state_tuple[10] if len(state_tuple) > 10 else ""
                 insurance = state_tuple[11] if len(state_tuple) > 11 else ""
                 cost = state_tuple[12] if len(state_tuple) > 12 else ""
                 acq_date = state_tuple[13] if len(state_tuple) > 13 else ""
                 phys_status = state_tuple[14] if len(state_tuple) > 14 else ""
                 driver = state_tuple[15] if len(state_tuple) > 15 else ""'''

text = text.replace(loop_old, loop_new)

# Reformat append
app_old = '''initial_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status, driver))'''
app_new = '''initial_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, make, type_val, emission, gsis, lto, last_reg, cost, acq_date, phys_status, status, insurance, driver))'''
text = text.replace(app_old, app_new)

app_old_2 = '''manual_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status, driver))'''
app_new_2 = '''manual_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, make, type_val, emission, gsis, lto, last_reg, cost, acq_date, phys_status, status, insurance, driver))'''
text = text.replace(app_old_2, app_new_2)

app_old_3 = '''full_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status, driver))'''
app_new_3 = '''full_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, make, type_val, emission, gsis, lto, last_reg, cost, acq_date, phys_status, status, insurance, driver))'''
text = text.replace(app_old_3, app_new_3)

# And one special loop for initial alerts where variables were named by me without "val_"
init_old = '''            status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
            owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
            office = state_tuple[4] if len(state_tuple) > 4 else ""
            engine = state_tuple[5] if len(state_tuple) > 5 else ""
            chassis = state_tuple[6] if len(state_tuple) > 6 else ""
            brand = state_tuple[7] if len(state_tuple) > 7 else ""
            year = state_tuple[8] if len(state_tuple) > 8 else ""
            cost = state_tuple[9] if len(state_tuple) > 9 else ""
            acq_date = state_tuple[10] if len(state_tuple) > 10 else ""
            phys_status = state_tuple[11] if len(state_tuple) > 11 else ""
            driver = state_tuple[12] if len(state_tuple) > 12 else ""'''

init_new = '''            status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
            owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
            office = state_tuple[4] if len(state_tuple) > 4 else ""
            make = state_tuple[5] if len(state_tuple) > 5 else ""
            type_val = state_tuple[6] if len(state_tuple) > 6 else ""
            emission = state_tuple[7] if len(state_tuple) > 7 else ""
            gsis = state_tuple[8] if len(state_tuple) > 8 else ""
            lto = state_tuple[9] if len(state_tuple) > 9 else ""
            last_reg = state_tuple[10] if len(state_tuple) > 10 else ""
            insurance = state_tuple[11] if len(state_tuple) > 11 else ""
            cost = state_tuple[12] if len(state_tuple) > 12 else ""
            acq_date = state_tuple[13] if len(state_tuple) > 13 else ""
            phys_status = state_tuple[14] if len(state_tuple) > 14 else ""
            driver = state_tuple[15] if len(state_tuple) > 15 else ""'''
text = text.replace(init_old, init_new)

# Save
with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Patching complete.")
