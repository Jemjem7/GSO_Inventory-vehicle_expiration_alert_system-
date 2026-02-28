import os
import sys
import time
import threading
import queue
import pandas as pd
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
            target_date = pd.to_datetime(exp_date, dayfirst=True).date()
            
        today = datetime.now().date()
        delta_days = (target_date - today).days

        if delta_days <= 0:
            return 'EXPIRED (RED)'
        elif 1 <= delta_days <= 14:
            return 'DAYS BEFORE EXPIRY (ORANGE)'
        elif 15 <= delta_days <= 29:
            return '2-WEEK NOTICE (YELLOW)'
        else:
            return 'SUFFICIENT TIME (GREEN)'
    except Exception as e:
        return 'PLEASE INPUT LAST REG (GRAY)'

class AlertWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("⚠ Vehicle Expiration Alert")
        self.attributes('-topmost', True)
        self.configure(bg='#f0f0f0')
        self.withdraw() # Hide immediately on launch
        
        window_width = 440
        window_height = 300
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = screen_width - window_width - 20
        y = screen_height - window_height - 60
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.main_container = tk.Frame(self, bg='#f0f0f0')
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
        
    def build_ui(self, detailed_alerts, window_title):
        for w in self.main_container.winfo_children():
            w.destroy()
            
        header = tk.Label(self.main_container, text=window_title, font=("Segoe UI", 12, "bold"), bg='#f0f0f0', fg='#333333')
        header.pack(pady=(15, 5))
        
        summary_frame = tk.Frame(self.main_container, bg='#ffffff', bd=1, relief=tk.SUNKEN)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        importance_order = [
            ('EXPIRED', '#d32f2f'),
            ('DAYS BEFORE EXPIRY', '#e65100'),
            ('2-WEEK NOTICE', '#f57f17'),
            ('SUFFICIENT TIME', '#2e7d32'),
            ('PLEASE INPUT LAST REG', '#616161'),
            ('REGISTERED', '#1565c0')
        ]
        
        has_alerts = False
        canvas = tk.Canvas(summary_frame, bg='#ffffff', highlightthickness=0)
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#ffffff')
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Sort out and display
        for status_key, color in importance_order:
            matching_plates = []
            for full_status, plates in detailed_alerts.items():
                if status_key in full_status:
                    if isinstance(plates, list):
                        matching_plates.extend(plates)
            
            if matching_plates:
                # Helper for sorting plates by exact date "PLATE (YYYY-MM-DD)"
                def extract_date(p_str):
                    try:
                        part = str(p_str).split('(')[-1].strip(')')
                        return datetime.strptime(part, '%Y-%m-%d')
                    except Exception:
                        return datetime.max
                        
                matching_plates.sort(key=extract_date)

                text = f"• {len(matching_plates)} {status_key}"
                lbl = tk.Label(scrollable_frame, text=text, font=("Segoe UI", 10, "bold"), bg='#ffffff', fg=color, anchor="w")
                lbl.pack(fill=tk.X, padx=10, pady=(5, 0))
                
                plates_text = ", ".join(str(p) for p in matching_plates)
                plates_lbl = tk.Message(scrollable_frame, text=plates_text, font=("Segoe UI", 8), bg='#ffffff', fg='#555555', width=300, justify=tk.LEFT)
                plates_lbl.pack(fill=tk.X, padx=25, pady=(0, 5))
                has_alerts = True
                
        if not has_alerts:
             lbl = tk.Label(scrollable_frame, text="All vehicles are up to date.", font=("Segoe UI", 10), bg='#ffffff', fg='#2e7d32')
             lbl.pack(pady=10)
        
        self.status_lbl = tk.Label(self.main_container, text="", bg='#f0f0f0', font=("Segoe UI", 8, "italic"), fg='#666666')
        self.status_lbl.pack(pady=(0, 2))

        btn_frame = tk.Frame(self.main_container, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        lbl_action = tk.Label(btn_frame, text="Run Manual Scan:", bg='#f0f0f0', font=("Segoe UI", 9))
        lbl_action.pack(side=tk.LEFT)
        
        btn_scan_all = ttk.Button(btn_frame, text="Scan All", command=self.do_scan_all, width=10)
        btn_scan_all.pack(side=tk.RIGHT)
        
        self.sheet_var = tk.StringVar()
        dropdown_values = current_sheets if current_sheets else ["No Sheets Found"]
        self.sheet_var.set("Select Month...")
        
        style = ttk.Style()
        style.theme_use('clam')
        sheet_dropdown = ttk.OptionMenu(btn_frame, self.sheet_var, "Select Month...", *dropdown_values, command=self.do_scan_month)
        sheet_dropdown.config(width=18)
        sheet_dropdown.pack(side=tk.RIGHT, padx=5)

    def do_scan_all(self):
        self.status_lbl.config(text="Scanning all sheets in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE,), daemon=True).start()

    def do_scan_month(self, selection):
        if selection and selection != "Select Month..." and selection != "No Sheets Found":
            self.status_lbl.config(text=f"Scanning {selection} in background...")
            threading.Thread(target=process_excel, args=(EXCEL_FILE, selection), daemon=True).start()

def send_notification(detailed_alerts, title="⚠ Vehicle Update Detected"):
    if not detailed_alerts:
        return
    gui_queue.put({'type': 'show', 'alerts': detailed_alerts, 'title': title})

def format_plate_with_date(plate, exp_date):
    if pd.isna(exp_date):
        return plate
    try:
        dt_str = exp_date.strftime('%Y-%m-%d') if hasattr(exp_date, 'strftime') else str(exp_date).split(' ')[0]
        return f"{plate} ({dt_str})"
    except:
        return f"{plate} ({exp_date})"

def find_header_row(filepath, sheet_name):
    """
    Scans the first 15 rows looking for "PLATE". 
    Returns the integer index of the row to use as the header.
    """
    try:
        df_test = pd.read_excel(filepath, engine='openpyxl', nrows=15, header=None, sheet_name=sheet_name)
        for i, row in df_test.iterrows():
            if any(isinstance(v, str) and 'PLATE' in v.upper() for v in row.values):
                return i
    except:
        pass
    return 3 # fallback default

def process_excel(filepath, manual_sheet_target=None):
    global previous_state, first_run, current_sheets
    
    try:
        if not os.path.exists(filepath):
            if manual_sheet_target is not None:
                print(f"{Fore.RED}File not found. Cannot scan.{Style.RESET_ALL}")
            return False

        # Load specific sheet or all sheets
        if manual_sheet_target:
            h_row = find_header_row(filepath, manual_sheet_target)
            dfs = pd.read_excel(filepath, engine='openpyxl', header=h_row, sheet_name=manual_sheet_target)
            if isinstance(dfs, pd.DataFrame):
                dfs = {manual_sheet_target: dfs}
        else:
            xl = pd.ExcelFile(filepath, engine='openpyxl')
            dfs = {}
            for sh in xl.sheet_names:
                h_row = find_header_row(filepath, sh)
                dfs[sh] = pd.read_excel(filepath, engine='openpyxl', header=h_row, sheet_name=sh)
            
        if manual_sheet_target is None:
            current_sheets = list(dfs.keys())
                
    except Exception as e:
        if manual_sheet_target is not None:
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
                elif 'DAYS BEFORE' in val and 'NOTICE' not in val and '2-WEEK' not in val:
                    status = 'DAYS BEFORE EXPIRY (ORANGE)'
                elif '2-WEEK' in val or '2 WEEK' in val or '15 TO' in val:
                    status = '2-WEEK NOTICE (YELLOW)'
                elif 'SUFFICIENT' in val:
                    status = 'SUFFICIENT TIME (GREEN)'
                elif 'INPUT' in val:
                    status = 'PLEASE INPUT LAST REG (GRAY)'
                elif 'REGISTERED' in val or 'YES' in val:
                    status = 'REGISTERED (BLUE)'
                else: 
                     # If the text is weird, assume registered to be safe
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
        if manual_sheet_target is not None:
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
            if 'EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status:
                if status not in initial_alerts:
                    initial_alerts[status] = []
                initial_alerts[status].append(format_plate_with_date(plate, exp_date))
        
        print(f"{Fore.CYAN}--- End Initial Scan ---{Style.RESET_ALL}")
        
        if initial_alerts:
             send_notification(initial_alerts, title="⚠ Initial Scan Results")
        else:
             send_notification({"SUFFICIENT TIME (GREEN)": ["All Plates inside Excel File"]}, title="⚠ Initial Scan Results")
        
    elif combined_changed_records or manual_sheet_target is not None:
        # User requested a specific sheet or requested "Scan All"
        if manual_sheet_target is not None:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Manual Scan: {manual_sheet_target}{Style.RESET_ALL}")
             title_text = f"⚠ Manual Scan: {manual_sheet_target}"
             
             manual_alerts = {}
             for plate, state_tuple in combined_current_state.items():
                 status, exp_date = state_tuple[0], state_tuple[1]
                 if 'EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status:
                     if status not in manual_alerts:
                         manual_alerts[status] = []
                     manual_alerts[status].append(format_plate_with_date(plate, exp_date))
             
             if manual_alerts:
                 send_notification(manual_alerts, title=title_text)
             else:
                 send_notification({"SUFFICIENT TIME (GREEN)": [f"All vehicles in {manual_sheet_target} are valid."]}, title=title_text)
             return True

        if manual_sheet_target is None:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Change Detected!{Style.RESET_ALL}")
             
        summary_alerts = {}
        for record in combined_changed_records:
            plate = record['plate']
            old = record['old_status']
            new = record['new_status']
            sheet = record['sheet']
            exp_date = record['exp_date']
            if manual_sheet_target is None:
                 print_status(f"Update ({sheet}): [{plate}] {old} -> {new}", new)
            
            if 'EXPIRED' in new or 'DAYS BEFORE' in new or '2-WEEK' in new:
                if new not in summary_alerts:
                    summary_alerts[new] = []
                summary_alerts[new].append(format_plate_with_date(plate, exp_date))
                
        if summary_alerts:
            # If the user did "Scan All", manual_sheet_target is None, but combined_changed_records would be empty since it's just checking states
            # Wait, "Scan All" triggers a process_excel(EXCEL_FILE) which just recalculates states. We must send a summary of EVERYTHING if they clicked it!
            pass # See check below
            
    # CRITICAL: If they clicked Scan All, it passes manual_sheet_target=None, but it's NOT the first run! 
    # Therefore changed_records will be empty if no files changed. 
    # But they just clicked the button, so they expect a popup immediately showing all expired!
    # Let's check if there are no changes but it was triggered manually.
    if manual_sheet_target is None and not first_run:
        # Check if we were called rapidly from the tray or button without file change?
        # A bit tricky. We assume if it came from the thread, it's either the file watcher or Scan All.
        # It's safer if Scan All prints the whole state again.
        # But for now, let's keep the file watcher working mostly.
        # For a manually triggered "Scan All", we will send the full list!
        pass

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
                    time.sleep(1)
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
    # Generate full report of current state!
    if previous_state:
        full_alerts = {}
        for plate, state_tuple in previous_state.items():
            status, exp_date = state_tuple[0], state_tuple[1]
            if 'EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status:
                if status not in full_alerts:
                    full_alerts[status] = []
                full_alerts[status].append(format_plate_with_date(plate, exp_date))
        
        if full_alerts:
             send_notification(full_alerts, title="⚠ Scan All Results")
        else:
             send_notification({"SUFFICIENT TIME (GREEN)": ["All Vehicles Registered"]}, title="⚠ Scan All Results")
    else:
        # If not initialized, process it
        process_excel(EXCEL_FILE)
    
def make_scan_sheet_callback(sheet_name):
    def callback(icon, item):
        print(f"Manually Scanning: {sheet_name}")
        process_excel(EXCEL_FILE, manual_sheet_target=sheet_name)
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
