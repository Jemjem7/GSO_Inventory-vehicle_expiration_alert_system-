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
import socket
from PIL import Image, ImageDraw
import traceback
import tkinter as tk
from tkinter import ttk
import winsound
import win32com.client

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
    elif "1-WEEK ADVANCE" in status_col:
        color = Fore.LIGHTRED_EX
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

        if delta_days < 0:
            return 'EXPIRED (RED)'
        elif 0 <= delta_days <= 7:
            return '1 WEEK BEFORE EXPIRY (RED)'
        elif 8 <= delta_days <= 30:
            return '1 MONTH BEFORE EXPIRY (ORANGE)'
        elif 31 <= delta_days <= 60:
            return '2 MONTHS BEFORE EXPIRY (YELLOW)'
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
        
        window_width = 1300
        window_height = 750
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.main_container = tk.Frame(self)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.clock_after_id = None

        # Trap protocol so the user simply HIDES the window on X, no destroying!
        self.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        # Start checking the queue
        self.check_queue()

    def update_clock(self):
        if hasattr(self, 'clock_label') and self.clock_label.winfo_exists():
            now = datetime.now()
            time_str = now.strftime("%I:%M:%S %p")
            date_str = now.strftime("%m/%d/%Y")
            self.clock_label.config(text=f"{date_str}   |   {time_str}")
            self.clock_after_id = self.after(1000, self.update_clock)

    def hide_window(self):
        self.withdraw()

    def check_queue(self):
        try:
            while True:
                msg = gui_queue.get_nowait()
                if msg['type'] == 'show':
                    # Play sound without blocking if it was an automatic background scan
                    if msg.get('is_auto', False):
                        def play_alert():
                            try:
                                # Bell-like chime (Single strike with slight resonance)
                                winsound.Beep(1200, 300) 
                                winsound.Beep(800, 200) 
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
            stripe_1 = '#2D2E31'
            stripe_2 = '#35363A'
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
            stripe_1 = '#FFFFFF'
            stripe_2 = '#F8F9FA'
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
        
        # Use the main container for placing elements securely
        
        # Dedicated top bar 
        top_bar = tk.Frame(self.main_container, bg="black")
        top_bar.pack(fill=tk.X, padx=0, pady=(0, 5))
            
        banner_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banner.jpg")
        logo_left_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_left.png")
        logo_right_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_right.png")
        
        try:
            from PIL import Image, ImageTk
            # Left Logo
            if os.path.exists(logo_left_path):
                img_l = Image.open(logo_left_path)
                img_l.thumbnail((120, 120), Image.Resampling.LANCZOS)
                self.logo_l_photo = ImageTk.PhotoImage(img_l)
                ll_label = tk.Label(top_bar, image=self.logo_l_photo, bg="black")
                ll_label.pack(side=tk.LEFT, padx=(40, 0), pady=10)
            
            # Right Logo
            if os.path.exists(logo_right_path):
                img_r = Image.open(logo_right_path)
                img_r.thumbnail((120, 120), Image.Resampling.LANCZOS)
                self.logo_r_photo = ImageTk.PhotoImage(img_r)
                rr_label = tk.Label(top_bar, image=self.logo_r_photo, bg="black")
                rr_label.pack(side=tk.RIGHT, padx=(0, 40), pady=10)
        except Exception as e:
            print(f"Error loading logos: {e}")
            
        header_text = "Republic of the Philippines\nLocal Government Unit of Manolo Fortich\nGENERAL SERVICE OFFICE\nVEHICULAR RECORDS"
        header = tk.Label(top_bar, text=header_text, font=("Segoe UI", 16, "bold"), bg="black", fg="white", justify="center")
        header.pack(expand=True, anchor="center", pady=15)
        
        # Pack bottom controls FIRST so they claim the bottom edge securely
        self.status_lbl = tk.Label(self.main_container, text="", bg=bg_color, font=("Segoe UI", 9, "italic"), fg=sub_fg)
        self.status_lbl.pack(side=tk.BOTTOM, pady=(5, 5))

        btn_frame = tk.Frame(self.main_container, bg=bg_color)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 15))
        
        # Now pack the summary frame to expand and claim remaining space
        summary_frame = tk.Frame(self.main_container, bg=panel_bg, bd=0)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 10))
        
        # Clock inside the panel directly
        clock_frame = tk.Frame(summary_frame, bg=panel_bg)
        clock_frame.pack(fill=tk.X, pady=(10, 5), padx=10)
        self.clock_label = tk.Label(clock_frame, text="", font=("Segoe UI", 14, "bold"), bg=panel_bg, fg=fg_color)
        self.clock_label.pack()
        self.update_clock()
        
        has_alerts = False
        
        # Calculate stats
        expired_count = 0
        expired_by_month = {}
        for full_status, plates in detailed_alerts.items():
            if ("EXPIRY" in full_status or "EXPIRED" in full_status) and isinstance(plates, list):
                for p_str in plates:
                    expired_count += 1
                    try:
                        data = json.loads(p_str)
                        month_name = data.get("sheet", "Unknown")
                    except:
                        month_name = "Unknown Date"
                    
                    expired_by_month[month_name] = expired_by_month.get(month_name, 0) + 1
                    
        # Add Stats Label
        self.stats_label = tk.Label(summary_frame, text="", font=("Segoe UI", 11, "bold"), bg=panel_bg, fg=fg_color)
        if expired_count > 0:
            month_stats = " | ".join([f"{k}: {v}" for k, v in expired_by_month.items()])
            stats_text = f"Total Expired: {expired_count}    ({month_stats})"
        else:
            stats_text = "Total Expired: 0"
        self.stats_label.config(text=stats_text)
        self.stats_label.pack(pady=(0, 10))
        
        # Setup Treeview Table
        columns = ("office", "plate", "engine", "chassis", "brand", "year", "date", "cost", "acq", "owner", "status", "alert", "month", "sheet")
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
        tree.heading("status", text="STATUS", anchor=tk.W)
        tree.heading("alert", text="ALERT", anchor=tk.W)
        tree.heading("month", text="MONTH", anchor=tk.W)
        tree.heading("sheet", text="Sheet", anchor=tk.W) 
        
        tree.column("office", width=70, minwidth=60, stretch=tk.NO)
        tree.column("plate", width=110, minwidth=100, stretch=tk.NO)
        tree.column("engine", width=120, minwidth=100, stretch=tk.YES)
        tree.column("chassis", width=120, minwidth=100, stretch=tk.YES)
        tree.column("brand", width=120, minwidth=100, stretch=tk.YES)
        tree.column("year", width=60, minwidth=50, stretch=tk.NO)
        tree.column("date", width=110, minwidth=100, stretch=tk.NO)
        tree.column("cost", width=90, minwidth=70, stretch=tk.NO)
        tree.column("acq", width=110, minwidth=90, stretch=tk.NO)
        tree.column("owner", width=140, minwidth=110, stretch=tk.YES)
        tree.column("status", width=90, minwidth=80, stretch=tk.NO)
        tree.column("alert", width=160, minwidth=120, stretch=tk.YES)
        tree.column("month", width=90, minwidth=80, stretch=tk.NO)
        tree.column("sheet", width=0, minwidth=0, stretch=tk.NO)
        
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(summary_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Setup clean row styling (remove striping backgrounds)
        tree.tag_configure('evenrow', background=bg_color)
        tree.tag_configure('oddrow', background=bg_color)
        
        # Hover effect styling
        hover_color = '#35363A' if actual_theme == 'Dark' else '#E8EAED'
        tree.tag_configure('hover', background=hover_color)
        
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
                        data = json.loads(p_str)
                        d_str = data.get("date", "N/A")
                        return datetime.strptime(d_str, '%Y-%m-%d')
                    except Exception:
                        return datetime.max
                        
                matching_plates.sort(key=extract_date)
                
                for p_str in matching_plates:
                    try:
                        data = json.loads(p_str)
                    except:
                        data = {}
                        
                    office = data.get("office", "")
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
                    tree.insert("", tk.END, values=(office, plate, engine, chassis, brand, year, date_val, cost, acq_date, owner, phys_status, alert_val, sheet_name, sheet_name), tags=(status_key, stripe_tag))
                    row_count += 1
                    has_alerts = True
                    
        if has_alerts:
            last_click_time = [0.0]
            def on_row_click(event):
                # Only open if they clicked on an actual item
                current_time = time.time()
                if current_time - last_click_time[0] < 2.0:
                    return # Debounce multiple clicks
                    
                region = tree.identify("region", event.x, event.y)
                if region == "cell" or region == "tree":
                    item_id = tree.identify_row(event.y)
                    if item_id:
                        values = tree.item(item_id, 'values')
                        if len(values) >= 14: # Get the sheet name from the hidden column
                            sheet_to_open = values[13]
                            last_click_time[0] = current_time
                            def open_excel_threaded():
                                try:
                                    import win32com.client
                                    import pythoncom
                                    pythoncom.CoInitialize() # required for threads
                                    
                                    abs_path = os.path.abspath(EXCEL_FILE)
                                    excel = None
                                    wb = None
                                    
                                    # Try to link to an already open instance of Excel First
                                    try:
                                        excel = win32com.client.GetActiveObject("Excel.Application")
                                        for w in excel.Workbooks:
                                            if w.FullName.lower() == abs_path.lower():
                                                wb = w
                                                break
                                    except:
                                        pass
                                        
                                    if not wb:
                                        # It's not open. Open it normally so it registers with Windows properly.
                                        os.startfile(abs_path)
                                        # Give it a moment to load so COM can grab it
                                        time.sleep(2.5) 
                                        try:
                                            excel = win32com.client.GetActiveObject("Excel.Application")
                                            for w in excel.Workbooks:
                                                if w.FullName.lower() == abs_path.lower():
                                                    wb = w
                                                    break
                                        except:
                                            pass
                                            
                                    if wb:
                                        try:
                                            # Found the workbook, jump to the exact sheet!
                                            if sheet_to_open and sheet_to_open != "Unknown":
                                                for sh in wb.Sheets:
                                                    if sh.Name == sheet_to_open:
                                                        sh.Activate()
                                                        break
                                        except Exception as e:
                                            print(f"Failed to activate sheet {sheet_to_open}: {e}")
                                            
                                        # Force Excel to the front
                                        try:
                                            excel.Visible = True
                                            import win32gui
                                            import win32con
                                            hwnd = excel.Hwnd
                                            if hwnd:
                                                if win32gui.IsIconic(hwnd):
                                                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                                                win32gui.SetForegroundWindow(hwnd)
                                        except Exception as e:
                                            print(f"Failed to bring to front: {e}")
                                            
                                    pythoncom.CoUninitialize()
                                except Exception as e:
                                    print(f"Major COM failure: {e}")
                                    try:
                                        os.startfile(os.path.abspath(EXCEL_FILE))
                                    except:
                                        pass
                                    
                            threading.Thread(target=open_excel_threaded, daemon=True).start()
                        
            # Hover fading light effect
            self.last_hovered_item = None
            def on_tree_motion(event):
                item = tree.identify_row(event.y)
                if item != self.last_hovered_item:
                    # Clear previous hover
                    if self.last_hovered_item and tree.exists(self.last_hovered_item):
                        tags = list(tree.item(self.last_hovered_item, "tags"))
                        if "hover" in tags:
                            tags.remove("hover")
                            tree.item(self.last_hovered_item, tags=tags)
                            
                    # Set new hover
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
             lbl = tk.Label(summary_frame, text="All vehicles are up to date.", font=("Segoe UI", 10), bg=panel_bg, fg='#66cc66' if actual_theme == 'Dark' else '#2e7d32')
             lbl.pack(pady=20)
        
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

def send_notification(detailed_alerts, title="⚠ Vehicle Update Detected", is_auto=False):
    if not detailed_alerts:
        return
    gui_queue.put({'type': 'show', 'alerts': detailed_alerts, 'title': title, 'is_auto': is_auto})

def format_plate_with_data(plate, exp_date, sheet_name="Unknown", owner="Unknown", office="", engine="", chassis="", brand="", year="", cost="", acq_date="", phys_status="", alert=""):
    if pd.isna(exp_date) or str(exp_date).strip() == '':
        dt_str = "N/A"
    else:
        try:
            if not hasattr(exp_date, 'strftime'):
                exp_date_str = str(exp_date).replace('\\', '/')
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
    })

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
        # Implement retry logic for file reading to avoid crashes when Excel is saving
        file_buffer = None
        for attempt in range(4):
            try:
                with open(filepath, 'rb') as f:
                    file_buffer = io.BytesIO(f.read())
                break
            except PermissionError as pe:
                if attempt < 3:
                    time.sleep(1)
                else:
                    raise pe

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
        
        owner_col_candidates = [c for c in df_sheet.columns if 'NAME' in str(c).upper() or 'OWNER' in str(c).upper() or 'CUSTOMER' in str(c).upper() or 'ACCOUNTABLE' in str(c).upper() or 'PERSON' in str(c).upper()]
        owner_col = owner_col_candidates[0] if owner_col_candidates else None
        
        exp_col_candidates = [c for c in df_sheet.columns if 'REMINDER' in str(c).upper() or 'EXPIRATION' in str(c).upper() or 'EXPIRY' in str(c).upper() or ('DATE' in str(c).upper() and 'ACQUISITION' not in str(c).upper())]
        exp_col = exp_col_candidates[0] if exp_col_candidates else 'REMINDER'
        
        status_col_keys = [c for c in df_sheet.columns if 'REGISTERED' in str(c).upper()]
        status_col = status_col_keys[0] if status_col_keys else None
        
        phys_status_keys = [c for c in df_sheet.columns if 'STATUS' in str(c).upper() and 'NOT' not in str(c).upper()]
        phys_status_col = phys_status_keys[0] if phys_status_keys else None
        
        alert_col_candidates = [c for c in df_sheet.columns if 'ALERT' in str(c).upper() and 'SYSTEM' not in str(c).upper()]
        alert_col = alert_col_candidates[0] if alert_col_candidates else None

        office_c = [c for c in df_sheet.columns if 'OFFICE' in str(c).upper()]
        engine_c = [c for c in df_sheet.columns if 'ENGINE' in str(c).upper()]
        chassis_c = [c for c in df_sheet.columns if 'CHASSIS' in str(c).upper()]
        brand_c = [c for c in df_sheet.columns if 'BRAND' in str(c).upper() or 'BODY TYPE' in str(c).upper()]
        year_c = [c for c in df_sheet.columns if 'YEAR' in str(c).upper()]
        cost_c = [c for c in df_sheet.columns if 'COST' in str(c).upper()]
        acq_date_c = [c for c in df_sheet.columns if 'ACQUISITION DATE' in str(c).upper()]
        
        office_col = office_c[0] if office_c else None
        engine_col = engine_c[0] if engine_c else None
        chassis_col = chassis_c[0] if chassis_c else None
        brand_col = brand_c[0] if brand_c else None
        year_col = year_c[0] if year_c else None
        cost_col = cost_c[0] if cost_c else None
        acq_date_col = acq_date_c[0] if acq_date_c else None

        if plate_col not in df_sheet.columns:
            continue
            
        current_state = {}
        changed_records = []
        
        for index, row in df_sheet.iterrows():
            plate = row[plate_col]
            owner = str(row[owner_col]).strip() if owner_col and pd.notna(row[owner_col]) else "Unknown"
            
            val_office = str(row[office_col]).strip() if office_col and pd.notna(row[office_col]) else ""
            val_engine = str(row[engine_col]).strip() if engine_col and pd.notna(row[engine_col]) else ""
            val_chassis = str(row[chassis_col]).strip() if chassis_col and pd.notna(row[chassis_col]) else ""
            val_brand = str(row[brand_col]).strip() if brand_col and pd.notna(row[brand_col]) else ""
            val_year = str(row[year_col]).strip() if year_col and pd.notna(row[year_col]) else ""
            if val_year and val_year.endswith(".0"): val_year = val_year[:-2]
            val_cost = str(row[cost_col]).strip() if cost_col and pd.notna(row[cost_col]) else ""
            
            acq_d = row[acq_date_col] if acq_date_col and pd.notna(row[acq_date_col]) else ""
            val_acq_date = ""
            if acq_d != "":
                try:
                    if hasattr(acq_d, 'strftime'): val_acq_date = acq_d.strftime('%Y-%m-%d')
                    else: val_acq_date = str(acq_d).split(" ")[0]
                except:
                   val_acq_date = str(acq_d)
                   
            val_phys_status = str(row[phys_status_col]).strip() if phys_status_col and pd.notna(row[phys_status_col]) else ""
            
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
                if 'EXPIRED' in val or 'LESS THAN' in val:
                    status = 'EXPIRED (RED)'
                elif '1 WEEK' in val or '1-WEEK' in val or ('WEEK' in val and '1' in val) or '1 TO 7' in val or '1-7' in val:
                    status = '1 WEEK BEFORE EXPIRY (RED)'
                elif '1 MONTH' in val or '1-MONTH' in val or 'WEEK' in val or '8 TO 30' in val or '8-30' in val or '30 DAYS' in val:
                    status = '1 MONTH BEFORE EXPIRY (ORANGE)'
                elif '2 MONTH' in val or '2-MONTH' in val or '60 DAYS' in val or '31 TO 60' in val or '31-60' in val:
                    status = '2 MONTHS BEFORE EXPIRY (YELLOW)'
                elif 'SUFFICIENT' in val or 'MORE' in val:
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
                
            current_state[plate] = (status, exp_date, sheet_name, owner, val_office, val_engine, val_chassis, val_brand, val_year, val_cost, val_acq_date, val_phys_status)
            
            if not first_run or manual_sheet_target is not None:
                old_state = previous_state.get(plate, None)
                if old_state is not None:
                    old_status = old_state[0]
                    old_exp = old_state[1] if len(old_state) > 1 else None
                    old_sheet = old_state[2] if len(old_state) > 2 else "Unknown"
                        
                    if old_status != status or old_exp != exp_date or old_sheet != sheet_name:
                        changed_records.append({
                            'plate': plate,
                            'owner': owner,
                            'old_status': old_status,
                            'new_status': status,
                            'sheet': sheet_name,
                            'exp_date': exp_date
                        })
                elif old_state is None and ('EXPIRED' in status or 'DAYS BEFORE' in status or '2-WEEK' in status or '1-WEEK' in status):
                     changed_records.append({
                        'plate': plate,
                        'owner': owner,
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
            status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
            owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
            office = state_tuple[4] if len(state_tuple) > 4 else ""
            engine = state_tuple[5] if len(state_tuple) > 5 else ""
            chassis = state_tuple[6] if len(state_tuple) > 6 else ""
            brand = state_tuple[7] if len(state_tuple) > 7 else ""
            year = state_tuple[8] if len(state_tuple) > 8 else ""
            cost = state_tuple[9] if len(state_tuple) > 9 else ""
            acq_date = state_tuple[10] if len(state_tuple) > 10 else ""
            phys_status = state_tuple[11] if len(state_tuple) > 11 else ""
            print_status(f"[{plate}] {status}", status)
            if status not in initial_alerts:
                initial_alerts[status] = []
            initial_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status))
        
        print(f"{Fore.CYAN}--- End Initial Scan ---{Style.RESET_ALL}")
        
        if initial_alerts:
             send_notification(initial_alerts, title="⚠ Initial Scan Results", is_auto=True)
        else:
             send_notification({"SUFFICIENT TIME": ["All Plates inside Excel File"]}, title="⚠ Initial Scan Results", is_auto=True)
        
    elif combined_changed_records or is_manual_scan:
        # User requested a specific sheet or requested "Scan All"
        if is_manual_scan:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Manual Scan Triggered{Style.RESET_ALL}")
             title_text = f"⚠ Manual Scan: {manual_sheet_target if manual_sheet_target else 'All Sheets'}"
             
             manual_alerts = {}
             # Just pull from the results of what we read!
             for plate, state_tuple in combined_current_state.items():
                 status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
                 owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
                 office = state_tuple[4] if len(state_tuple) > 4 else ""
                 engine = state_tuple[5] if len(state_tuple) > 5 else ""
                 chassis = state_tuple[6] if len(state_tuple) > 6 else ""
                 brand = state_tuple[7] if len(state_tuple) > 7 else ""
                 year = state_tuple[8] if len(state_tuple) > 8 else ""
                 cost = state_tuple[9] if len(state_tuple) > 9 else ""
                 acq_date = state_tuple[10] if len(state_tuple) > 10 else ""
                 phys_status = state_tuple[11] if len(state_tuple) > 11 else ""
                 if status not in manual_alerts:
                     manual_alerts[status] = []
                 manual_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status))
             
             if manual_alerts:
                 send_notification(manual_alerts, title=title_text, is_auto=False)
             else:
                 send_notification({"SUFFICIENT TIME": [f"All vehicles checked are valid."]}, title=title_text, is_auto=False)
             return True

        if not is_manual_scan:
             print(f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Background Change Detected!{Style.RESET_ALL}")
             changed_sheets = list(set([r['sheet'] for r in combined_changed_records]))
             sheet_title_str = ", ".join(changed_sheets) if len(changed_sheets) < 3 else f"{len(changed_sheets)} Sheets"
             
             for record in combined_changed_records:
                 plate = record['plate']
                 owner = record.get('owner', 'Unknown')
                 old = record['old_status']
                 new = record['new_status']
                 sheet = record['sheet']
                 print_status(f"Real-time Update ({sheet}): [{plate}] ({owner}) {old} -> {new}", new)
                 
             # Send comprehensive updated state so UI refreshes real-time
             full_alerts = {}
             for plate, state_tuple in combined_current_state.items():
                 status, exp_date, sheet_name = state_tuple[0], state_tuple[1], state_tuple[2]
                 owner = state_tuple[3] if len(state_tuple) > 3 else "Unknown"
                 office = state_tuple[4] if len(state_tuple) > 4 else ""
                 engine = state_tuple[5] if len(state_tuple) > 5 else ""
                 chassis = state_tuple[6] if len(state_tuple) > 6 else ""
                 brand = state_tuple[7] if len(state_tuple) > 7 else ""
                 year = state_tuple[8] if len(state_tuple) > 8 else ""
                 cost = state_tuple[9] if len(state_tuple) > 9 else ""
                 acq_date = state_tuple[10] if len(state_tuple) > 10 else ""
                 phys_status = state_tuple[11] if len(state_tuple) > 11 else ""
                 if status not in full_alerts:
                     full_alerts[status] = []
                 full_alerts[status].append(format_plate_with_data(plate, exp_date, sheet_name, owner, office, engine, chassis, brand, year, cost, acq_date, phys_status, status))
                     
             if full_alerts:
                 send_notification(full_alerts, title=f"⚠ Real-time File Update: {sheet_title_str}", is_auto=True)
             else:
                 send_notification({"SUFFICIENT TIME": ["All Vehicles clear in latest update!"]}, title=f"⚠ Real-time File Update: {sheet_title_str}", is_auto=True)
            
    if manual_sheet_target is None:
        previous_state = combined_current_state
        first_run = False
        
    return True

def background_monitor():
    global monitor_active
    last_mtime = 0
    last_checked_date = datetime.now().date()
    
    while monitor_active:
        try:
            current_date = datetime.now().date()
            if current_date != last_checked_date:
                # Force rescan automatically at midnight/new day
                last_mtime = 0 
                last_checked_date = current_date
                
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
    print("Manually Scanning Excel...")
    threading.Thread(target=process_excel, args=(EXCEL_FILE,), kwargs={'is_manual_scan': True}, daemon=True).start()
    
def make_scan_sheet_callback(sheet_name):
    def callback(icon, item):
        print(f"Manually Scanning: {sheet_name}")
        threading.Thread(target=process_excel, args=(EXCEL_FILE,), kwargs={'manual_sheet_target': sheet_name, 'is_manual_scan': True}, daemon=True).start()
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
        items = [pystray.MenuItem('Scan Excel', on_scan_all), pystray.Menu.SEPARATOR]
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
    # --- Single Instance Lock ---
    global lock_socket
    lock_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        lock_socket.bind(('127.0.0.1', 47123))
        
        def listen_for_triggers():
            while monitor_active:
                try:
                    lock_socket.settimeout(1.0)
                    data, _ = lock_socket.recvfrom(1024)
                    if data == b'trigger':
                        print("Received trigger from another instance!")
                        threading.Thread(target=process_excel, args=(EXCEL_FILE,), kwargs={'is_manual_scan': True}, daemon=True).start()
                except socket.timeout:
                    continue
                except Exception as e:
                    break
                    
        threading.Thread(target=listen_for_triggers, daemon=True).start()
    except socket.error:
        print("Vehicle Monitor is already running. Pinging the active instance...")
        try:
            client_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            client_sock.sendto(b'trigger', ('127.0.0.1', 47123))
        except:
            pass
        sys.exit(0)
    # --- End Single Instance Lock ---
    
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
