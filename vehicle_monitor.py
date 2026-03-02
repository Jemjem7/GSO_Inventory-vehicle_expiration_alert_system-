import os
import sys
import time
import threading
import queue
import json
import winreg
import pandas as pd
import io
from datetime import datetime
import colorama
from colorama import Fore, Style
import pystray
from PIL import Image, ImageDraw
import traceback
import tkinter as tk
from tkinter import ttk
import winsound

colorama.init(autoreset=True)

# Configuration
EXCEL_FILE = 'VehicleMonitoring.xlsx'
CHECK_INTERVAL_SECONDS = 5

def get_system_theme():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "Light" if value else "Dark"
    except Exception:
        return "Light"

def load_settings():
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r") as f:
                return json.load(f)
    except:
        pass
    return {"theme": "System"}

def save_settings(settings):
    try:
        with open("settings.json", "w") as f:
            json.dump(settings, f)
    except:
        pass

app_settings = load_settings()

# State
previous_state = {}
first_run = True
current_sheets = []
monitor_active = True
tray_icon = None

gui_queue = queue.Queue()

def create_image():
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 4, height // 4, width * 3 // 4, height * 3 // 4), fill=(0, 120, 215)) # Blue square
    dc.text((width // 3 + 2, height // 3 + 5), "V", fill=(255, 255, 255))
    return image

def print_status(message, status_col=""):
    color = Style.RESET_ALL
    if "EXPIRED" in status_col:
        color = Fore.RED
    elif "DAYS BEFORE EXPIRY" in status_col:
        color = Fore.YELLOW
    elif "2-WEEK NOTICE" in status_col:
        color = Fore.LIGHTYELLOW_EX
    elif "SUFFICIENT TIME" in status_col:
        color = Fore.GREEN
    elif "PLEASE INPUT LAST REG" in status_col:
        color = Fore.LIGHTBLACK_EX
    elif "REGISTERED" in status_col:
        color = Fore.CYAN
        
    print(f"{color}{message}{Style.RESET_ALL}")

def get_expiration_status(exp_date, status_override):
    # Fallback status generator if the user's Excel sheet doesn't calculate the ALERT column
    if pd.notna(status_override) and str(status_override).strip().upper() in ['YES', 'REGISTERED']:
        return 'REGISTERED (BLUE)'
    if pd.isna(exp_date) or str(exp_date).strip() == '':
        return 'PLEASE INPUT LAST REG (GRAY)'
    try:
        if isinstance(exp_date, pd.Timestamp) or isinstance(exp_date, datetime):
            target_date = exp_date.date()
        else:
            exp_date_str = str(exp_date).replace('\\', '/')
            target_date = pd.to_datetime(exp_date_str, dayfirst=True).date()
            
        today = datetime.now().date()
        delta_days = (target_date - today).days

        if delta_days <= 0:
            return 'EXPIRED (RED)'
        elif 1 <= delta_days <= 14:
            return 'DAYS BEFORE EXPIRY (ORANGE)'
        elif 15 <= delta_days <= 29:
            return 'DAYS BEFORE 2 WEEK NOTICE (YELLOW)'
        else:
            return 'SUFFICIENT TIME (GREEN)'
    except Exception as e:
        return 'PLEASE INPUT LAST REG (GRAY)'

class AlertWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("⚠ Vehicle Expiration Alert")
        self.attributes('-topmost', True)
        self.current_theme = app_settings.get("theme", "System")
        self.last_alerts = {}
        self.last_title = ""
        self.withdraw() # Hide immediately on launch
        
        window_width = 580
        window_height = 380
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = screen_width - window_width - 20
        y = screen_height - window_height - 60
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.main_container = tk.Frame(self)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # Trap protocol so the user simply HIDES the window on X, no destroying!
        self.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        # Start checking the queue
        self.check_queue()

    def hide_window(self):
        self.withdraw()

    def check_queue(self):
        try:
            while True:
                msg = gui_queue.get_nowait()
                if msg['type'] == 'show':
                    # Play sound without blocking
                    def play_alert():
                        try:
                            winsound.Beep(1500, 150)
                            winsound.Beep(2000, 250)
                        except:
                            pass
                    threading.Thread(target=play_alert, daemon=True).start()
                    
                    self.build_ui(msg['alerts'], msg['title'])
                    self.deiconify()
                    self.lift()
                    self.attributes('-topmost', True)
                    
                elif msg['type'] == 'exit':
                    self.quit()
                    self.destroy()
                    return
        except queue.Empty:
            pass
        self.after(200, self.check_queue)
        
    def change_theme(self, selection):
        self.current_theme = selection
        app_settings["theme"] = selection
        save_settings(app_settings)
        if self.last_alerts:
            self.build_ui(self.last_alerts, self.last_title)

    def build_ui(self, detailed_alerts, window_title):
        self.last_alerts = detailed_alerts
        self.last_title = window_title
        
        for w in self.main_container.winfo_children():
            w.destroy()
            
        actual_theme = get_system_theme() if self.current_theme == "System" else self.current_theme
        
        if actual_theme == "Dark":
            bg_color = '#202124'
            fg_color = '#E8EAED'
            panel_bg = '#2D2E31'
            text_fg = '#E8EAED'
            sub_fg = '#9AA0A6'
            stripe_1 = '#2D2E31'
            stripe_2 = '#35363A'
            importance_order = [
                ('EXPIRED', '#F28B82'),
                ('DAYS BEFORE EXPIRY', '#FDC69C'),
                ('DAYS BEFORE 2 WEEK NOTICE', '#FCE8E6'),
                ('PLEASE INPUT LAST REG', '#9AA0A6')
            ]
        else:
            bg_color = '#F1F3F4'
            fg_color = '#202124'
            panel_bg = '#FFFFFF'
            text_fg = '#202124'
            sub_fg = '#5F6368'
            stripe_1 = '#FFFFFF'
            stripe_2 = '#F8F9FA'
            importance_order = [
                ('EXPIRED', '#D93025'),
                ('DAYS BEFORE EXPIRY', '#E37400'),
                ('DAYS BEFORE 2 WEEK NOTICE', '#F29900'),
                ('PLEASE INPUT LAST REG', '#80868B')
            ]
            
        self.configure(bg=bg_color)
        self.main_container.configure(bg=bg_color)
            
        header = tk.Label(self.main_container, text=window_title, font=("Segoe UI", 14, "bold"), bg=bg_color, fg=fg_color)
        header.pack(pady=(20, 10))
        
        summary_frame = tk.Frame(self.main_container, bg=panel_bg, bd=0)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))
        
        has_alerts = False
        
        # Setup Treeview Table
        columns = ("plate", "date", "status")
        tree = ttk.Treeview(summary_frame, columns=columns, show="headings", style="Custom.Treeview", height=10)
        
        tree.heading("plate", text="Plate Number", anchor=tk.W)
        tree.heading("date", text="Expiration / Status", anchor=tk.W)
        tree.heading("status", text="Condition", anchor=tk.W)
        
        tree.column("plate", width=150, minwidth=120, stretch=tk.NO)
        tree.column("date", width=150, minwidth=120, stretch=tk.NO)
        tree.column("status", width=220, minwidth=180, stretch=tk.YES)
        
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Setup row striping tags
        tree.tag_configure('evenrow', background=stripe_1)
        tree.tag_configure('oddrow', background=stripe_2)
        
        # Sort out and display
        row_count = 0
        for status_key, color in importance_order:
            tree.tag_configure(status_key, foreground=color)
            matching_plates = []
            
            for full_status, plates in detailed_alerts.items():
                if status_key in full_status:
                    if isinstance(plates, list):
                        matching_plates.extend(plates)
            
            if matching_plates:
                def extract_date(p_str):
                    try:
                        if '||' in str(p_str):
                            part = str(p_str).split('||')[-1].strip()
                            return datetime.strptime(part, '%Y-%m-%d')
                        return datetime.max
                    except Exception:
                        return datetime.max
                        
                matching_plates.sort(key=extract_date)
                
                for p_str in matching_plates:
                    p_str = str(p_str)
                    if '||' in p_str:
                        plate, date_val = p_str.split('||', 1)
                    else:
                        plate = p_str
                        date_val = "N/A"
                        
                    # Insert row
                    stripe_tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                    tree.insert("", tk.END, values=(plate, date_val, status_key), tags=(status_key, stripe_tag))
                    row_count += 1
                    has_alerts = True
                    
        if has_alerts:
            tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        else:
             lbl = tk.Label(summary_frame, text="All vehicles are up to date.", font=("Segoe UI", 10), bg=panel_bg, fg='#66cc66' if actual_theme == 'Dark' else '#2e7d32')
             lbl.pack(pady=20)
        
        self.status_lbl = tk.Label(self.main_container, text="", bg=bg_color, font=("Segoe UI", 9, "italic"), fg=sub_fg)
        self.status_lbl.pack(pady=(5, 5))

        btn_frame = tk.Frame(self.main_container, bg=bg_color)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        # Stylization for ttk buttons
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
        
        # Theme Dropdown
        theme_frame = tk.Frame(btn_frame, bg=bg_color)
        theme_frame.pack(side=tk.LEFT)
        
        lbl_theme = tk.Label(theme_frame, text="Theme:", bg=bg_color, fg=fg_color, font=("Segoe UI", 9))
        lbl_theme.pack(side=tk.LEFT)
        
        self.theme_var = tk.StringVar(value=self.current_theme)
        
        theme_dropdown = ttk.OptionMenu(theme_frame, self.theme_var, self.current_theme, "Light", "Dark", "System", command=self.change_theme)
        theme_dropdown.config(width=7)
        theme_dropdown.pack(side=tk.LEFT, padx=5)
        # Apply menu styling
        theme_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')
        
        # Spacer
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
        # Apply menu styling
        sheet_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')

    def do_scan_all(self):
        self.status_lbl.config(text="Scanning all sheets in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE, None, True), daemon=True).start()

    def do_scan_month(self, selection):
        if selection and selection != "Select Month..." and selection != "No Sheets Found":
            self.status_lbl.config(text=f"Scanning {selection} in background...")
            threading.Thread(target=process_excel, args=(EXCEL_FILE, selection, True), daemon=True).start()

def send_notification(detailed_alerts, title="⚠ Vehicle Update Detected"):
    if not detailed_alerts:
        return
    gui_queue.put({'type': 'show', 'alerts': detailed_alerts, 'title': title})

def format_plate_with_date(plate, exp_date):
    if pd.isna(exp_date) or str(exp_date).strip() == '':
        return f"{plate}||N/A"
    try:
        if not hasattr(exp_date, 'strftime'):
            exp_date_str = str(exp_date).replace('\\', '/')
            exp_date = pd.to_datetime(exp_date_str, dayfirst=True)
            
        dt_str = exp_date.strftime('%Y-%m-%d')
        return f"{plate}||{dt_str}"
    except:
        return f"{plate}||{exp_date}"

def find_header_row(excel_file_obj, sheet_name):
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
    return 3 # fallback default

def process_excel(filepath, manual_sheet_target=None, is_manual_scan=False):
    global previous_state, first_run, current_sheets
    
    try:
        if not os.path.exists(filepath):
            if is_manual_scan:
                print(f"{Fore.RED}File not found. Cannot scan.{Style.RESET_ALL}")
            return False

        # To avoid file lock/sharing violations, read file into memory first
        with open(filepath, 'rb') as f:
            file_buffer = io.BytesIO(f.read())

        # Load specific sheet or all sheets
        with pd.ExcelFile(file_buffer, engine='openpyxl') as xl:
            if manual_sheet_target:
                h_row = find_header_row(xl, manual_sheet_target)
                dfs = pd.read_excel(xl, header=h_row, sheet_name=manual_sheet_target)
                if isinstance(dfs, pd.DataFrame):
                    dfs = {manual_sheet_target: dfs}
            else:
                dfs = {}
                for sh in xl.sheet_names:
                    h_row = find_header_row(xl, sh)
                    dfs[sh] = pd.read_excel(xl, header=h_row, sheet_name=sh)
                
            if manual_sheet_target is None:
                current_sheets = list(dfs.keys())
                
    except Exception as e:
        if is_manual_scan:
             print(f"{Fore.RED}Error loading Excel: {e}{Style.RESET_ALL}")
        return False

    all_data = []
    
    for sheet_name, df_sheet in dfs.items():
        if df_sheet.empty:
            continue
            
        df_sheet.columns = df_sheet.columns.astype(str).str.strip().str.replace('\n', ' ')

        # Find the dynamic columns
        plate_col_candidates = [c for c in df_sheet.columns if 'PLATE' in str(c).upper()]
        plate_col = plate_col_candidates[0] if plate_col_candidates else 'PLATE #'
        
        exp_col_candidates = [c for c in df_sheet.columns if 'REMINDER' in str(c).upper()]
        exp_col = exp_col_candidates[0] if exp_col_candidates else 'REMINDER'
        
        status_col_keys = [c for c in df_sheet.columns if 'REGISTERED' in str(c).upper()]
        status_col = status_col_keys[0] if status_col_keys else None
        
        alert_col_candidates = [c for c in df_sheet.columns if 'ALERT' in str(c).upper() and 'SYSTEM' not in str(c).upper()]
        alert_col = alert_col_candidates[0] if alert_col_candidates else None

        if plate_col not in df_sheet.columns:
            continue
            
        current_state = {}
        changed_records = []
        
        for index, row in df_sheet.iterrows():
            plate = row[plate_col]
            if pd.isna(plate) or str(plate).strip() == '' or str(plate).upper() == 'CRITERIA':
                # Avoid breaking fully if there is just an empty row, unless it explicitly says CRITERIA
                if str(plate).upper() == 'CRITERIA':
                    break
                continue
                
            plate = str(plate).strip()
            exp_date = row[exp_col] if exp_col in df_sheet.columns and pd.notna(row[exp_col]) else None
            
            status = None
            # NATIVE EXCEL ALERT READING
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

            # Fallback if no alert mapped from Excel
            if not status:
                status_override = None
                if status_col and pd.notna(row[status_col]):
                    val = str(row[status_col]).strip().upper()
                    if val in ['YES', 'REGISTERED']:
                        status_override = 'REGISTERED'
                status = get_expiration_status(exp_date, status_override)
                
            current_state[plate] = (status, exp_date)
            
            if not first_run or manual_sheet_target is not None:
                old_state = previous_state.get(plate, None)
                if old_state is not None:
                    old_status, old_exp = old_state
                    if old_status != status or old_exp != exp_date:
                        changed_records.append({
                            'plate': plate,
                            'old_status': old_status,
                            'new_status': status,
                            'sheet': sheet_name,
                            'exp_date': exp_date
                        })
                elif old_state is None and ('EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status):
                     changed_records.append({
                        'plate': plate,
                        'old_status': 'NEW RECORD',
                        'new_status': status,
                        'sheet': sheet_name,
                        'exp_date': exp_date
                    })
                    
        all_data.append((current_state, changed_records, sheet_name))

    if not all_data:
        if is_manual_scan:
            print(f"{Fore.RED}No matching plates found.{Style.RESET_ALL}")
        return False

    combined_current_state = {}
    combined_changed_records = []
    
    for c_state, c_records, s_name in all_data:
        combined_current_state.update(c_state)
        combined_changed_records.extend(c_records)

    if first_run and manual_sheet_target is None:
        print(f"{Fore.CYAN}--- Initial Scan Results ({len(dfs)} sheets checked) ---{Style.RESET_ALL}")
        initial_alerts = {}
        for plate, state_tuple in combined_current_state.items():
            status, exp_date = state_tuple[0], state_tuple[1]
            print_status(f"[{plate}] {status}", status)
            if status not in initial_alerts:
                initial_alerts[status] = []
            initial_alerts[status].append(format_plate_with_date(plate, exp_date))
        
        print(f"{Fore.CYAN}--- End Initial Scan ---{Style.RESET_ALL}")
        
        if initial_alerts:
             send_notification(initial_alerts, title="⚠ Initial Scan Results")
        else:
             send_notification({"SUFFICIENT TIME (GREEN)": ["All Plates inside Excel File"]}, title="⚠ Initial Scan Results")
        
    elif combined_changed_records or is_manual_scan:
        # User requested a specific sheet or requested "Scan All"
        if is_manual_scan:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Manual Scan Triggered{Style.RESET_ALL}")
             title_text = f"⚠ Manual Scan: {manual_sheet_target if manual_sheet_target else 'All Sheets'}"
             
             manual_alerts = {}
             # Just pull from the results of what we read!
             for plate, state_tuple in combined_current_state.items():
                 status, exp_date = state_tuple[0], state_tuple[1]
                 if status not in manual_alerts:
                     manual_alerts[status] = []
                 manual_alerts[status].append(format_plate_with_date(plate, exp_date))
             
             if manual_alerts:
                 send_notification(manual_alerts, title=title_text)
             else:
                 send_notification({"SUFFICIENT TIME (GREEN)": [f"All vehicles checked are valid."]}, title=title_text)
             return True

        if not is_manual_scan:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Background Change Detected!{Style.RESET_ALL}")
             changed_sheets = list(set([r['sheet'] for r in combined_changed_records]))
             sheet_title_str = ", ".join(changed_sheets) if len(changed_sheets) < 3 else f"{len(changed_sheets)} Sheets"
             
             for record in combined_changed_records:
                 plate = record['plate']
                 old = record['old_status']
                 new = record['new_status']
                 sheet = record['sheet']
                 print_status(f"Real-time Update ({sheet}): [{plate}] {old} -> {new}", new)
                 
             # Send comprehensive updated state so UI refreshes real-time
             full_alerts = {}
             for plate, state_tuple in combined_current_state.items():
                 status, exp_date = state_tuple[0], state_tuple[1]
                 if status not in full_alerts:
                     full_alerts[status] = []
                 full_alerts[status].append(format_plate_with_date(plate, exp_date))
                     
             if full_alerts:
                 send_notification(full_alerts, title=f"⚠ Real-time File Update: {sheet_title_str}")
             else:
                 send_notification({"SUFFICIENT TIME (GREEN)": ["All Vehicles clear in latest update!"]}, title=f"⚠ Real-time File Update: {sheet_title_str}")
            
    if manual_sheet_target is None:
        previous_state = combined_current_state
        first_run = False
        
    return True

def background_monitor():
    global monitor_active
    last_mtime = 0
    
    while monitor_active:
        try:
            if os.path.exists(EXCEL_FILE):
                current_mtime = os.path.getmtime(EXCEL_FILE)
                if current_mtime != last_mtime:
                    # Added slightly more sleep to avoid lock race conditions with heavy Excel saves
                    time.sleep(2)
                    process_excel(EXCEL_FILE)
                    try:
                        last_mtime = os.path.getmtime(EXCEL_FILE)
                    except WindowsError:
                        pass
            time.sleep(CHECK_INTERVAL_SECONDS)
        except Exception as e:
            time.sleep(CHECK_INTERVAL_SECONDS)

# Manual Scan All via Tray (Sends entire overview)
def on_scan_all(icon, item):
    print("Manually Scanning All...")
    process_excel(EXCEL_FILE, is_manual_scan=True)
    
def make_scan_sheet_callback(sheet_name):
    def callback(icon, item):
        print(f"Manually Scanning: {sheet_name}")
        process_excel(EXCEL_FILE, manual_sheet_target=sheet_name, is_manual_scan=True)
    return callback

def on_exit(icon, item):
    global monitor_active
    monitor_active = False
    icon.stop()
    print("Exiting...")
    gui_queue.put({'type': 'exit'})
    
def pystray_runner():
    global tray_icon
    image = create_image()
    tray_icon = pystray.Icon("VehicleMonitor", image, "Vehicle Alert System")
    
    def setup_menu():
        items = [pystray.MenuItem('Scan All', on_scan_all), pystray.Menu.SEPARATOR]
        if current_sheets:
            sheet_menus = []
            for sheet in current_sheets:
                sheet_menus.append(pystray.MenuItem(f"Scan {sheet}", make_scan_sheet_callback(sheet)))
            items.append(pystray.MenuItem('Scan Month...', pystray.Menu(*sheet_menus)))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem('Exit', on_exit))
        return items

    tray_icon.menu = pystray.Menu(setup_menu)
    tray_icon.run()

def main():
    print(f"{Fore.GREEN}Starting Vehicle Monitor Dashboard...{Style.RESET_ALL}")
    
    # Pre-scan the initial file so tray menu builds immediately
    process_excel(EXCEL_FILE)

    monitor_thread = threading.Thread(target=background_monitor, daemon=True)
    monitor_thread.start()
    
    tray_thread = threading.Thread(target=pystray_runner, daemon=True)
    tray_thread.start()
    
    # TKinter Main Window must be in main thread
    window = AlertWindow()
    window.mainloop()

if __name__ == "__main__":
    main()
