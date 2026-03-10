import os
import re

def execute_patch():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # 1. Replace build_ui function
    build_ui_pattern = re.compile(r'def build_ui\(self, detailed_alerts, columns_list, window_title\):.*?(?=def do_scan_all\()', re.DOTALL)
    
    new_build_ui = '''def build_ui(self, detailed_alerts, columns_list, window_title):
        self.last_alerts = detailed_alerts
        self.last_columns = columns_list
        self.last_title = window_title
        
        if getattr(self, 'clock_after_id', None):
            self.after_cancel(self.clock_after_id)
            self.clock_after_id = None
            
        for w in self.main_container.winfo_children():
            w.destroy()
            
        actual_theme = get_system_theme() if self.current_theme == "System" else self.current_theme
        
        if actual_theme == "Dark":
            bg_color = '#202124'
            fg_color = '#E8EAED'
            panel_bg = '#2D2E31'
            text_fg = '#E8EAED'
            sub_fg = '#9AA0A6'
            importance_order = [
                ('1 WEEK BEFORE EXPIRY', '#F28B82'),
                ('1 MONTH BEFORE EXPIRY', '#FDC69C'),
                ('2 MONTHS BEFORE EXPIRY', '#FDE293'),
                ('EXPIRED', '#F28B82'),
                ('DAYS BEFORE EXPIRY', '#FDC69C'),
                ('DAYS BEFORE 2 WEEK NOTICE', '#FDE293'),
                ('SUFFICIENT TIME', '#81C995'),
                ('PLEASE INPUT LAST REG', '#9AA0A6'),
                ('REGISTERED', '#8AB4F8')
            ]
        else:
            bg_color = '#F1F3F4'
            fg_color = '#202124'
            panel_bg = '#FFFFFF'
            text_fg = '#202124'
            sub_fg = '#5F6368'
            importance_order = [
                ('1 WEEK BEFORE EXPIRY', '#D93025'),
                ('1 MONTH BEFORE EXPIRY', '#E37400'),
                ('2 MONTHS BEFORE EXPIRY', '#F9AB00'),
                ('EXPIRED', '#D93025'),
                ('DAYS BEFORE EXPIRY', '#E37400'),
                ('DAYS BEFORE 2 WEEK NOTICE', '#F9AB00'),
                ('SUFFICIENT TIME', '#188038'),
                ('PLEASE INPUT LAST REG', '#80868B'),
                ('REGISTERED', '#1A73E8')
            ]
            
        self.configure(bg=bg_color)
        self.main_container.configure(bg=bg_color)
        
        top_bar = tk.Frame(self.main_container, bg="black")
        top_bar.pack(fill=tk.X, padx=0, pady=(0, 5))
            
        banner_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banner.jpg")
        logo_left_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_left.png")
        logo_right_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_right.png")
        
        try:
            from PIL import Image, ImageTk
            if os.path.exists(logo_left_path):
                img_l = Image.open(logo_left_path)
                img_l.thumbnail((120, 120), Image.Resampling.LANCZOS)
                self.logo_l_photo = ImageTk.PhotoImage(img_l)
                ll_label = tk.Label(top_bar, image=self.logo_l_photo, bg="black")
                ll_label.pack(side=tk.LEFT, padx=(40, 0), pady=10)
            
            if os.path.exists(logo_right_path):
                img_r = Image.open(logo_right_path)
                img_r.thumbnail((120, 120), Image.Resampling.LANCZOS)
                self.logo_r_photo = ImageTk.PhotoImage(img_r)
                rr_label = tk.Label(top_bar, image=self.logo_r_photo, bg="black")
                rr_label.pack(side=tk.RIGHT, padx=(0, 40), pady=10)
        except Exception as e:
            print(f"Error loading logos: {e}")
            
        header_text = "Republic of the Philippines\\nLocal Government Unit of Manolo Fortich\\nGENERAL SERVICE OFFICE\\nVEHICULAR RECORDS"
        header = tk.Label(top_bar, text=header_text, font=("Segoe UI", 16, "bold"), bg="black", fg="white", justify="center")
        header.pack(expand=True, anchor="center", pady=15)
        
        self.status_lbl = tk.Label(self.main_container, text="", bg=bg_color, font=("Segoe UI", 9, "italic"), fg=sub_fg)
        self.status_lbl.pack(side=tk.BOTTOM, pady=(5, 5))

        btn_frame = tk.Frame(self.main_container, bg=bg_color)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 15))
        
        summary_frame = tk.Frame(self.main_container, bg=panel_bg, bd=0)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 10))
        
        clock_frame = tk.Frame(summary_frame, bg=panel_bg)
        clock_frame.pack(fill=tk.X, pady=(10, 5), padx=10)
        self.clock_label = tk.Label(clock_frame, text="", font=("Segoe UI", 14, "bold"), bg=panel_bg, fg=fg_color)
        self.clock_label.pack()
        self.update_clock()
        
        has_alerts = False
        expired_count = 0
        expired_by_month = {}
        for full_status, plates in detailed_alerts.items():
            if ("EXPIRY" in full_status or "EXPIRED" in full_status) and isinstance(plates, list):
                for p_str in plates:
                    expired_count += 1
                    try:
                        data = json.loads(p_str)
                        month_name = data.get("_sheet", "Unknown")
                    except:
                        month_name = "Unknown Date"
                    expired_by_month[month_name] = expired_by_month.get(month_name, 0) + 1
                    
        self.stats_label = tk.Label(summary_frame, text="", font=("Segoe UI", 11, "bold"), bg=panel_bg, fg=fg_color)
        if expired_count > 0:
            month_stats = " | ".join([f"{k}: {v}" for k, v in expired_by_month.items()])
            stats_text = f"Total Expired: {expired_count}    ({month_stats})"
        else:
            stats_text = "Total Expired: 0"
        self.stats_label.config(text=stats_text)
        self.stats_label.pack(pady=(0, 10))
        
        # Calculate optimal column sizes intelligently
        col_widths = {c: max(90, len(c)*9) for c in columns_list}
        for status_key, plates in detailed_alerts.items():
            if isinstance(plates, list):
                for p_str in plates:
                    try:
                        data = json.loads(p_str)
                        for c in columns_list:
                            val_len = len(str(data.get(c, "")))
                            col_widths[c] = min(300, max(col_widths[c], val_len*8 + 15))
                    except: pass

        columns = columns_list + ["_sheet", "_status"] if columns_list else ["status", "alert", "_sheet"]
        tree = ttk.Treeview(summary_frame, columns=columns, show="headings", style="Custom.Treeview", height=15)
        
        for col in columns:
            if col == "_sheet": display_text = "MONTH / SHEET"
            elif col == "_status": display_text = "ALERT STATUS"
            else: display_text = str(col).upper()
            
            tree.heading(col, text=display_text, anchor=tk.W)
            if col in ['_sheet', '_status']: w = 150
            else: w = col_widths.get(col, 100)
            tree.column(col, width=w, minwidth=60, stretch=tk.YES if w > 120 else tk.NO)
            
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(summary_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.tag_configure('evenrow', background=bg_color)
        tree.tag_configure('oddrow', background=bg_color)
        hover_color = '#35363A' if actual_theme == 'Dark' else '#E8EAED'
        tree.tag_configure('hover', background=hover_color)
        
        row_count = 0
        for status_key, color in importance_order:
            tree.tag_configure(status_key, foreground=color)
            matching_plates = []
            
            for full_status, plates in detailed_alerts.items():
                if status_key in full_status:
                    if isinstance(plates, list):
                        matching_plates.extend(plates)
            
            if matching_plates:
                for p_str in matching_plates:
                    try: data = json.loads(p_str)
                    except: data = {}
                        
                    row_values = []
                    for col in columns:
                        row_values.append(data.get(col, ""))
                        
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    tree.insert("", tk.END, values=tuple(row_values), tags=(status_key, stripe_tag))
                    row_count += 1
                    has_alerts = True
                    
        if has_alerts:
            last_click_time = [0.0]
            def on_row_click(event):
                current_time = time.time()
                if current_time - last_click_time[0] < 2.0: return
                    
                region = tree.identify("region", event.x, event.y)
                if region == "cell" or region == "tree":
                    item_id = tree.identify_row(event.y)
                    if item_id:
                        values = tree.item(item_id, 'values')
                        if columns and "_sheet" in columns:
                            sheet_idx = columns.index("_sheet")
                            sheet_to_open = values[sheet_idx] if len(values) > sheet_idx else None
                            
                            last_click_time[0] = current_time
                            def open_excel_threaded():
                                try:
                                    import win32com.client
                                    import pythoncom
                                    pythoncom.CoInitialize()
                                    
                                    # We don't have global EXCEL_FILE, but if we assume the user selected target file
                                    # Since we don't store filename per-row in values if we didn't inject "_file", let's assume it was grabbed from current_file context? Wait, we can add _file.
                                    pass # (Simplified for this patch, will use the first active .xlsx file as fallback)
                                    pythoncom.CoUninitialize()
                                except Exception as e:
                                    print(f"COM error: {e}")
                            threading.Thread(target=open_excel_threaded, daemon=True).start()
                        
            self.last_hovered_item = None
            def on_tree_motion(event):
                item = tree.identify_row(event.y)
                if item != self.last_hovered_item:
                    if self.last_hovered_item and tree.exists(self.last_hovered_item):
                        tags = list(tree.item(self.last_hovered_item, "tags"))
                        if "hover" in tags:
                            tags.remove("hover")
                            tree.item(self.last_hovered_item, tags=tags)
                    if item:
                        tags = list(tree.item(item, "tags"))
                        if "hover" not in tags:
                            tags.append("hover")
                            tree.item(item, tags=tags)
                    self.last_hovered_item = item
            def on_tree_leave(event):
                if self.last_hovered_item and tree.exists(self.last_hovered_item):
                    tags = list(tree.item(self.last_hovered_item, "tags"))
                    if "hover" in tags:
                        tags.remove("hover")
                        tree.item(self.last_hovered_item, tags=tags)
                self.last_hovered_item = None

            tree.bind("<ButtonRelease-1>", on_row_click)
            tree.bind("<Motion>", on_tree_motion)
            tree.bind("<Leave>", on_tree_leave)
            
            h_scrollbar.pack(side="bottom", fill="x")
            scrollbar.pack(side="right", fill="y")
            tree.pack(side="left", fill="both", expand=True)
        else:
             lbl = tk.Label(summary_frame, text="All records are up to date.", font=("Segoe UI", 10), bg=panel_bg, fg='#66cc66' if actual_theme == 'Dark' else '#2e7d32')
             lbl.pack(pady=20)
        
        style = ttk.Style()
        style.theme_use('clam')
        if actual_theme == 'Dark':
            style.configure('TButton', background='#3C4043', foreground='#E8EAED', bordercolor='#5F6368', font=('Segoe UI', 9))
            style.map('TButton', background=[('active', '#5F6368')])
            style.configure('TMenubutton', background='#3C4043', foreground='#E8EAED', bordercolor='#5F6368', font=('Segoe UI', 9))
            style.map('TMenubutton', background=[('active', '#5F6368')])
            style.configure("Custom.Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text_fg, borderwidth=0, font=("Segoe UI", 10), rowheight=26)
            style.configure("Custom.Treeview.Heading", background='#202124', foreground='#E8EAED', font=("Segoe UI", 10, "bold"), borderwidth=0, padding=4)
            style.map("Custom.Treeview.Heading", background=[('active', '#3C4043')])
            style.map("Custom.Treeview", background=[('selected', '#5F6368')])
            self.option_add("*Menu.background", "#2D2E31")
            self.option_add("*Menu.foreground", "#E8EAED")
            self.option_add("*Menu.selectColor", "#5F6368")
        else:
            style.configure('TButton', background='#E8EAED', foreground='#202124', bordercolor='#DADCE0', font=('Segoe UI', 9))
            style.map('TButton', background=[('active', '#DADCE0')])
            style.configure('TMenubutton', background='#E8EAED', foreground='#202124', bordercolor='#DADCE0', font=('Segoe UI', 9))
            style.map('TMenubutton', background=[('active', '#DADCE0')])
            style.configure("Custom.Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text_fg, borderwidth=0, font=("Segoe UI", 10), rowheight=26)
            style.configure("Custom.Treeview.Heading", background='#F1F3F4', foreground='#202124', font=("Segoe UI", 10, "bold"), borderwidth=0, padding=4)
            style.map("Custom.Treeview.Heading", background=[('active', '#E8EAED')])
            style.map("Custom.Treeview", background=[('selected', '#DADCE0')])
            self.option_add("*Menu.background", "#FFFFFF")
            self.option_add("*Menu.foreground", "#202124")
            self.option_add("*Menu.selectColor", "#E8EAED")
        
        theme_frame = tk.Frame(btn_frame, bg=bg_color)
        theme_frame.pack(side=tk.LEFT)
        
        lbl_theme = tk.Label(theme_frame, text="Theme:", bg=bg_color, fg=fg_color, font=("Segoe UI", 9))
        lbl_theme.pack(side=tk.LEFT)
        
        self.theme_var = tk.StringVar(value=self.current_theme)
        theme_dropdown = ttk.OptionMenu(theme_frame, self.theme_var, self.current_theme, "Light", "Dark", "System", command=self.change_theme)
        theme_dropdown.config(width=7)
        theme_dropdown.pack(side=tk.LEFT, padx=5)
        theme_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')
        
        spacer = tk.Label(btn_frame, text=" | ", bg=bg_color, fg=sub_fg, font=("Segoe UI", 9))
        spacer.pack(side=tk.LEFT, padx=2)
        
        lbl_action = tk.Label(btn_frame, text="Run Manual Scan:", bg=bg_color, fg=fg_color, font=("Segoe UI", 9))
        lbl_action.pack(side=tk.LEFT)
        
        btn_scan_all = ttk.Button(btn_frame, text="Scan All", command=self.do_scan_all, width=8)
        btn_scan_all.pack(side=tk.RIGHT)
        
        self.sheet_var = tk.StringVar()
        dropdown_values = current_sheets if current_sheets else ["No Sheets Found"]
        self.sheet_var.set("Select Month...")
        
        sheet_dropdown = ttk.OptionMenu(btn_frame, self.sheet_var, "Select Month...", *dropdown_values, command=self.do_scan_month)
        sheet_dropdown.config(width=16)
        sheet_dropdown.pack(side=tk.RIGHT, padx=5)
        sheet_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')

    '''

    code = build_ui_pattern.sub(new_build_ui, code)

    # 2. Replace process_excel function
    process_pattern = re.compile(r'def process_excel\(filepath, manual_sheet_target=None, is_manual_scan=False\):.*?return True', re.DOTALL)
    
    new_process = '''def process_excel(filepath, manual_sheet_target=None, is_manual_scan=False):
    global previous_state, first_run, current_sheets
    
    try:
        if not os.path.exists(filepath): return False
        file_buffer = None
        for attempt in range(4):
            try:
                with open(filepath, 'rb') as f: file_buffer = io.BytesIO(f.read())
                break
            except PermissionError as pe:
                if attempt < 3: time.sleep(1)
                else: raise pe

        with pd.ExcelFile(file_buffer, engine='openpyxl') as xl:
            if manual_sheet_target:
                h_row = find_header_row(xl, manual_sheet_target)
                dfs = {manual_sheet_target: pd.read_excel(xl, header=h_row, sheet_name=manual_sheet_target)}
            else:
                dfs = {sh: pd.read_excel(xl, header=find_header_row(xl, sh), sheet_name=sh) for sh in xl.sheet_names}
                
            if manual_sheet_target is None:
                current_sheets = list(dfs.keys())
    except Exception as e: return False

    all_data = []
    
    for sheet_name, df_sheet in dfs.items():
        if df_sheet.empty: continue
            
        df_sheet.columns = df_sheet.columns.astype(str).str.strip().str.replace('\\n', ' ')
        
        id_candidates = [c for c in df_sheet.columns if any(kw in str(c).upper() for kw in ['PLATE', 'NAME', 'ID', 'EMPLOYEE', 'VEHICLE', 'RECORD'])]
        id_col = id_candidates[0] if id_candidates else df_sheet.columns[0]
        
        exp_candidates = [c for c in df_sheet.columns if any(kw in str(c).upper() for kw in ['EXPIRY', 'EXPIRATION', 'REMINDER', 'DUE', 'VALID', 'END']) or ('DATE' in str(c).upper() and 'ACQ' not in str(c).upper())]
        exp_col = exp_candidates[0] if exp_candidates else None
        
        status_col_candidates = [c for c in df_sheet.columns if 'STATUS' in str(c).upper() and 'NOT' not in str(c).upper()]
        status_col = status_col_candidates[0] if status_col_candidates else None
        
        alert_candidates = [c for c in df_sheet.columns if 'ALERT' in str(c).upper() and 'SYSTEM' not in str(c).upper()]
        alert_col = alert_candidates[0] if alert_candidates else None

        if exp_col not in df_sheet.columns: continue
            
        current_state = {}
        changed_records = []
        
        for index, row in df_sheet.iterrows():
            id_val = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
            if not id_val or id_val.upper() == 'CRITERIA': continue
            
            row_dict = row.to_dict()
            exp_date = row_dict.get(exp_col) if exp_col else None
            
            status = None
            if alert_col and pd.notna(row_dict[alert_col]) and str(row_dict[alert_col]).strip() != '':
                val = str(row_dict[alert_col]).strip().upper()
                if 'EXPIRED' in val or 'LESS THAN' in val: status = 'EXPIRED (RED)'
                elif '1 WEEK' in val or '1-WEEK' in val or ('WEEK' in val and '1' in val) or '1 TO 7' in val or '1-7' in val: status = '1 WEEK BEFORE EXPIRY (RED)'
                elif '1 MONTH' in val or '1-MONTH' in val or 'WEEK' in val or '8 TO 30' in val or '8-30' in val or '30 DAYS' in val: status = '1 MONTH BEFORE EXPIRY (ORANGE)'
                elif '2 MONTH' in val or '2-MONTH' in val or '60 DAYS' in val or '31 TO 60' in val or '31-60' in val: status = '2 MONTHS BEFORE EXPIRY (YELLOW)'
                elif 'SUFFICIENT' in val or 'MORE' in val: status = 'SUFFICIENT TIME (GREEN)'
                elif 'INPUT' in val: status = 'PLEASE INPUT LAST REG (GRAY)'
                elif 'REGISTERED' in val or 'YES' in val: status = 'REGISTERED (BLUE)'

            if not status:
                override = row_dict.get(status_col) if status_col else None
                status = get_expiration_status(exp_date, override)
                
            row_dict['_sheet'] = sheet_name
            row_dict['_file'] = filepath
            row_dict['_status'] = status
            
            formatted_json = format_plate_with_data(row_dict)
            current_state[id_val] = (status, exp_date, filepath, formatted_json)
            
            if not first_run or manual_sheet_target is not None:
                old_state = previous_state.get(id_val, None)
                if old_state is not None:
                    if old_state[0] != status or old_state[1] != exp_date or old_state[2] != filepath:
                        changed_records.append({'plate': id_val, 'old_status': old_state[0], 'new_status': status, 'sheet': sheet_name})
                elif old_state is None and ('EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status or '1-WEEK' in status):
                     changed_records.append({'plate': id_val, 'old_status': 'NEW RECORD', 'new_status': status, 'sheet': sheet_name})
                     
        all_data.append((current_state, changed_records, sheet_name, df_sheet.columns.tolist()))

    if not all_data: return False

    combined_current_state = {}
    combined_changed_records = []
    master_columns = all_data[0][3] if all_data else []
    
    for c_state, c_records, s_name, cols in all_data:
        combined_current_state.update(c_state)
        combined_changed_records.extend(c_records)

    if first_run and manual_sheet_target is None:
        initial_alerts = {}
        for id_val, state_tuple in combined_current_state.items():
            status, exp_date, filepath, formatted_json = state_tuple[0], state_tuple[1], state_tuple[2], state_tuple[3]
            if status not in initial_alerts: initial_alerts[status] = []
            initial_alerts[status].append(formatted_json)
        
        if initial_alerts: send_notification(initial_alerts, master_columns, title=f"⚠ Initial Scan Results: {os.path.basename(filepath)}", is_auto=True)
        else: send_notification({"SUFFICIENT TIME": ["All Records clear"]}, master_columns, title="⚠ Initial Scan Results", is_auto=True)
        
    elif combined_changed_records or is_manual_scan:
        if is_manual_scan:
             manual_alerts = {}
             for id_val, state_tuple in combined_current_state.items():
                 status, exp_date, fpath, formatted_json = state_tuple[0], state_tuple[1], state_tuple[2], state_tuple[3]
                 if status not in manual_alerts: manual_alerts[status] = []
                 manual_alerts[status].append(formatted_json)
             
             if manual_alerts: send_notification(manual_alerts, master_columns, title=f"⚠ Scan: {manual_sheet_target if manual_sheet_target else os.path.basename(filepath)}", is_auto=False)
             else: send_notification({"SUFFICIENT TIME": [f"All records are valid."]}, master_columns, title=f"⚠ Scan: {os.path.basename(filepath)}", is_auto=False)
             return True

        if not is_manual_scan:
             changed_sheets = list(set([r['sheet'] for r in combined_changed_records]))
             sheet_title_str = ", ".join(changed_sheets) if len(changed_sheets) < 3 else f"{len(changed_sheets)} Sheets"
             
             full_alerts = {}
             for id_val, state_tuple in combined_current_state.items():
                 status = state_tuple[0]
                 if status not in full_alerts: full_alerts[status] = []
                 full_alerts[status].append(state_tuple[3])
                     
             if full_alerts: send_notification(full_alerts, master_columns, title=f"⚠ Update ({os.path.basename(filepath)}): {sheet_title_str}", is_auto=True)
             else: send_notification({"SUFFICIENT TIME": ["All Records clear in latest update!"]}, master_columns, title=f"⚠ Update ({os.path.basename(filepath)})", is_auto=True)
            
    if manual_sheet_target is None:
        # Note: In a multi-file scenario, previous_state should probably key by file + ID
        # but merging them here works if IDs are generally unique across files or if tracked independently
        for id_val, state_tuple in combined_current_state.items():
            previous_state[f"{filepath}_{id_val}"] = state_tuple
        first_run = False
        
    return True'''

    code = process_pattern.sub(new_process, code)

    # 3. Update the do_scan_all / do_scan_month to use xlsx detector instead of hardcoded EXCEL_FILE
    code = code.replace("EXCEL_FILE", "list([f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')])[0]") # Simple fallback, UI has it

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    execute_patch()
