from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QComboBox, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QFrame)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont, QColor, QBrush, QPixmap
import threading
import time
import queue
import json
import os
import sys
import winsound
from datetime import datetime
import winreg

# We assume get_system_theme, clean_currency, app_settings, gui_queue, save_settings, EXCEL_FILE, process_excel, current_sheets exist in global scope of the final merged file.

class AlertWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("⚠ Vehicle Expiration Alert")
        self.last_alerts = {}
        self.last_title = ""
        self.first_popup_sound_played = False
        
        # In the context of vehicle_monitor.py, app_settings is global.
        self.current_theme = app_settings.get("theme", "Nature")
        
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, 'excel_scan_v3_final.ico')
            else:
                icon_path = os.path.abspath('excel_scan_v3_final.ico')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
            
            import ctypes
            myappid = 'localgov.gso.vehiclemonitor.1' 
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        self.resize(1300, 750)
        
        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # Top Bar
        self.top_bar = QFrame()
        self.top_bar_layout = QHBoxLayout(self.top_bar)
        self.top_bar_layout.setContentsMargins(50, 15, 50, 15)
        
        self.logo_l = QLabel()
        self.logo_r = QLabel()
        self.header_label = QLabel("Republic of the Philippines\nLocal Government Unit of Manolo Fortich\nGENERAL SERVICE OFFICE\nVEHICULAR RECORDS")
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Load logos if they exist
        logo_left_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_left.png")
        logo_right_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_right.png")
        
        if os.path.exists(logo_left_path):
            pix_l = QPixmap(logo_left_path).scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.logo_l.setPixmap(pix_l)
        if os.path.exists(logo_right_path):
            pix_r = QPixmap(logo_right_path).scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.logo_r.setPixmap(pix_r)
            
        self.top_bar_layout.addWidget(self.logo_l, 0, Qt.AlignmentFlag.AlignLeft)
        self.top_bar_layout.addWidget(self.header_label, 1, Qt.AlignmentFlag.AlignCenter)
        self.top_bar_layout.addWidget(self.logo_r, 0, Qt.AlignmentFlag.AlignRight)
        
        self.main_layout.addWidget(self.top_bar)
        
        # Content panel
        self.content_panel = QFrame()
        self.content_layout = QVBoxLayout(self.content_panel)
        self.content_layout.setContentsMargins(30, 20, 30, 10)
        self.main_layout.addWidget(self.content_panel, 1) # Expandable
        
        self.clock_label = QLabel()
        self.clock_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.content_layout.addWidget(self.clock_label)
        
        self.title_label = QLabel()
        self.content_layout.addWidget(self.title_label)
        
        self.stats_label = QLabel()
        self.content_layout.addWidget(self.stats_label)
        
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.content_layout.addWidget(self.table, 1)
        
        self.status_lbl = QLabel()
        self.main_layout.addWidget(self.status_lbl, 0, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft)
        
        # Bottom controls
        self.bottom_bar = QFrame()
        self.bottom_layout = QHBoxLayout(self.bottom_bar)
        self.bottom_layout.setContentsMargins(30, 5, 30, 15)
        
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Grain", "Nature", "Light", "Dark", "System"])
        idx = self.theme_combo.findText(self.current_theme)
        if idx >= 0: self.theme_combo.setCurrentIndex(idx)
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        
        self.scan_all_btn = QPushButton("Scan All")
        self.scan_all_btn.clicked.connect(self.do_scan_all)
        
        self.month_combo = QComboBox()
        self.month_combo.addItem("Select Month...")
        global current_sheets
        if current_sheets:
            self.month_combo.addItems(current_sheets)
        else:
            self.month_combo.addItem("No Sheets Found")
        
        self.month_combo.currentTextChanged.connect(lambda t: self.do_scan_month(t) if t not in ["Select Month...", "No Sheets Found", ""] else None)
        
        # Setup table click to open excel
        self.table.itemClicked.connect(self.on_row_click)
        
        bottom_left = QHBoxLayout()
        bottom_left.addWidget(QLabel("Theme:"))
        bottom_left.addWidget(self.theme_combo)
        
        bottom_right = QHBoxLayout()
        bottom_right.addWidget(QLabel("Run Manual Scan:"))
        bottom_right.addWidget(self.scan_all_btn)
        bottom_right.addWidget(self.month_combo)
        
        self.bottom_layout.addLayout(bottom_left)
        self.bottom_layout.addStretch()
        self.bottom_layout.addLayout(bottom_right)
        
        self.main_layout.addWidget(self.bottom_bar)
        
        # Setup clock timer
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_clock)
        self.clock_timer.start(1000)
        self.update_clock()
        
        # Setup queue checker
        self.queue_timer = QTimer(self)
        self.queue_timer.timeout.connect(self.check_queue)
        self.queue_timer.start(200)

        # Style UI immediately
        self.apply_stylesheet()
        
        # Start hidden
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.hide()

    def update_clock(self):
        now = datetime.now()
        time_str = now.strftime("%I:%M:%S %p")
        date_str = now.strftime("%m/%d/%Y")
        self.clock_label.setText(f"{date_str}   |   {time_str}")

    def closeEvent(self, event):
        # Prevent actually killing the app when user clicks X
        event.ignore()
        self.hide()

    def check_queue(self):
        try:
            while True:
                msg = gui_queue.get_nowait()
                if msg['type'] == 'show':
                    if msg.get('is_auto', False) and not self.first_popup_sound_played:
                        self.first_popup_sound_played = True
                        def play_alert():
                            try:
                                winsound.Beep(1200, 300) 
                                winsound.Beep(800, 200) 
                            except: pass
                        threading.Thread(target=play_alert, daemon=True).start()
                    
                    self.build_ui(msg['alerts'], msg['title'])
                    self.showNormal()
                    self.raise_()
                    self.activateWindow()
                elif msg['type'] == 'exit':
                    QApplication.quit()
                    return
        except queue.Empty:
            pass

    def change_theme(self, selection):
        self.current_theme = selection
        app_settings["theme"] = selection
        save_settings(app_settings)
        if self.last_alerts:
            self.build_ui(self.last_alerts, self.last_title)

    def apply_stylesheet(self):
        actual_theme = get_system_theme() if self.current_theme == "System" else self.current_theme
        
        if actual_theme == "Dark":
            bg = "#121212"; fg = "#E0E0E0"; panel = "#1E1E1E"; accent = "#04395E"; header_c = "#FFFFFF"
            top_bg = "#0F0F0F"; row1 = "#262626"; row2 = "#2A2A2A"; btn_bg = "#2D2D2D"; hover = "#333333"
        elif actual_theme == "Nature":
            bg = "#0D1410"; fg = "#E8F6F3"; panel = "#14221A"; accent = "#117A65"; header_c = "#A3E4D7"
            top_bg = "#090F0C"; row1 = "#192B21"; row2 = "#1E3529"; btn_bg = "#1A3B2F"; hover = "#234433"
        elif actual_theme == "Grain":
            bg = "#110F14"; fg = "#F4ECF7"; panel = "#1A1821"; accent = "#6C3483"; header_c = "#D7BDE2"
            top_bg = "#08070A"; row1 = "#211F2B"; row2 = "#282536"; btn_bg = "#3A1F4C"; hover = "#2C273D"
        else:
            bg = "#F5F5F5"; fg = "#202124"; panel = "#FFFFFF"; accent = "#E3F2FD"; header_c = "#202124"
            top_bg = "#FFFFFF"; row1 = "#FFFFFF"; row2 = "#FAFAFA"; btn_bg = "#F1F3F4"; hover = "#F8F9FA"

        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background-color: {bg};
                color: {fg};
                font-family: 'Segoe UI';
                font-size: 13px;
            }}
            QFrame#TopBar {{
                background-color: {top_bg};
            }}
            QFrame#ContentPanel {{
                background-color: {panel};
            }}
            QLabel#HeaderTitle {{
                color: {header_c};
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }}
            QLabel#Clock, QLabel#SectionTitle {{
                background-color: transparent;
                font-size: 18px;
                font-weight: bold;
            }}
            QLabel#Stats {{
                background-color: transparent;
                font-size: 14px;
                color: #A0A0A0;
            }}
            QTableWidget {{
                background-color: {panel};
                alternate-background-color: {row2};
                color: {fg};
                border: none;
                gridline-color: transparent;
                selection-background-color: {accent};
                selection-color: {fg};
            }}
            QTableWidget::item {{
                padding: 5px;
            }}
            QHeaderView::section {{
                background-color: {top_bg};
                color: {header_c};
                padding: 5px;
                border: none;
                font-weight: bold;
            }}
            QPushButton, QComboBox {{
                background-color: {btn_bg};
                color: {fg};
                border: 1px solid {hover};
                padding: 5px 15px;
                border-radius: 4px;
            }}
            QPushButton:hover, QComboBox:hover {{
                background-color: {hover};
            }}
            QComboBox::drop-down {{
                border: none;
            }}
            QComboBox QAbstractItemView {{
                background-color: {panel};
                color: {fg};
                selection-background-color: {accent};
            }}
            QScrollBar:vertical {{
                border: none;
                background: {bg};
                width: 12px;
                margin: 0px 0px 0px 0px;
            }}
            QScrollBar::handle:vertical {{
                background: {hover};
                min-height: 20px;
                border-radius: 6px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        
        self.top_bar.setObjectName("TopBar")
        self.content_panel.setObjectName("ContentPanel")
        self.header_label.setObjectName("HeaderTitle")
        self.clock_label.setObjectName("Clock")
        self.title_label.setObjectName("SectionTitle")
        self.stats_label.setObjectName("Stats")
        
        # Store for use in build_ui
        self.actual_theme = actual_theme

    def build_ui(self, detailed_alerts, window_title):
        self.last_alerts = detailed_alerts
        self.last_title = window_title
        self.apply_stylesheet()
        
        display_title = window_title.replace("⚠ ", "").replace("??? ", "").upper()
        self.title_label.setText(display_title)
        
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
                    
        if expired_count > 0:
            month_stats = " | ".join([f"{k}: {v}" for k, v in expired_by_month.items()])
            stats_text = f"Total Expired: {expired_count}   ({month_stats})"
        else:
            stats_text = "Total Expired: 0"
        self.stats_label.setText(stats_text)
        
        # Color palettes per theme
        colors_map = {}
        if self.actual_theme == "Dark":
            colors_map = {'1 WEEK BEFORE EXPIRY': '#EF5350', '1 MONTH BEFORE EXPIRY': '#FFA726', '2 MONTHS BEFORE EXPIRY': '#FFEE58', 'EXPIRED': '#EF5350', 'DAYS BEFORE EXPIRY': '#FFA726', 'DAYS BEFORE 2 WEEK NOTICE': '#FFEE58', 'SUFFICIENT TIME': '#66BB6A', 'PLEASE INPUT LAST REG': '#9E9E9E', 'REGISTERED': '#4FC3F7'}
        elif self.actual_theme == "Nature":
            colors_map = {'1 WEEK BEFORE EXPIRY': '#FF6B6B', '1 MONTH BEFORE EXPIRY': '#F4D03F', '2 MONTHS BEFORE EXPIRY': '#F7DC6F', 'EXPIRED': '#FF6B6B', 'DAYS BEFORE EXPIRY': '#F4D03F', 'DAYS BEFORE 2 WEEK NOTICE': '#F7DC6F', 'SUFFICIENT TIME': '#82E0AA', 'PLEASE INPUT LAST REG': '#839192', 'REGISTERED': '#85C1E9'}
        elif self.actual_theme == "Grain":
            colors_map = {'1 WEEK BEFORE EXPIRY': '#F1948A', '1 MONTH BEFORE EXPIRY': '#F8C471', '2 MONTHS BEFORE EXPIRY': '#FAD7A1', 'EXPIRED': '#F1948A', 'DAYS BEFORE EXPIRY': '#F8C471', 'DAYS BEFORE 2 WEEK NOTICE': '#FAD7A1', 'SUFFICIENT TIME': '#A9DFBF', 'PLEASE INPUT LAST REG': '#ABB2B9', 'REGISTERED': '#AED6F1'}
        else:
            colors_map = {'1 WEEK BEFORE EXPIRY': '#D32F2F', '1 MONTH BEFORE EXPIRY': '#F57C00', '2 MONTHS BEFORE EXPIRY': '#FBC02D', 'EXPIRED': '#D32F2F', 'DAYS BEFORE EXPIRY': '#F57C00', 'DAYS BEFORE 2 WEEK NOTICE': '#FBC02D', 'SUFFICIENT TIME': '#388E3C', 'PLEASE INPUT LAST REG': '#757575', 'REGISTERED': '#1976D2'}
            
        columns = ["STATUS (YES/NO)", "OFFICE", "PLATE #", "MAKE", "TYPE", "EMISSION", "GSIS", "LTO", "LAST REG.", "REMINDER", "ALERT", "INSURANCE (₱)", "DRIVER", "ACQ. COST (₱)", "DATE ACQUIRED", "MONTH", "Sheet_Hidden"]
        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        self.table.setColumnHidden(len(columns)-1, True) # Hide sheet column
        
        # Fetching out all proper plates
        all_rows = []
        # Re-use importance ordering logic implicitly mapping from Tkinter script
        importance_order = [
            '1 WEEK BEFORE EXPIRY', '1 MONTH BEFORE EXPIRY', '2 MONTHS BEFORE EXPIRY', 
            'EXPIRED', 'DAYS BEFORE EXPIRY', 'DAYS BEFORE 2 WEEK NOTICE', 
            'SUFFICIENT TIME', 'PLEASE INPUT LAST REG', 'REGISTERED'
        ]
        
        for status_key in importance_order:
            matching_plates = []
            for full_status, plates in detailed_alerts.items():
                if status_key in full_status:
                    if isinstance(plates, list):
                        matching_plates.extend(plates)
            
            if matching_plates:
                def extract_date(p_str):
                    try:
                        data = json.loads(p_str)
                        return datetime.strptime(data.get("date", "N/A"), '%Y-%m-%d')
                    except:
                        return datetime.max
                matching_plates.sort(key=extract_date)
                
                fg_hex = colors_map.get(status_key, "#FFFFFF" if self.actual_theme != "Light" else "#000000")
                fg_color = QColor(fg_hex)
                
                for p_str in matching_plates:
                    try: data = json.loads(p_str)
                    except: data = {}
                    
                    row_data = [
                        data.get("status", ""), data.get("office", ""), data.get("plate", "Unknown"),
                        data.get("make", ""), data.get("type", ""), data.get("emission", ""),
                        data.get("gsis", ""), data.get("lto", ""), data.get("last_reg", ""),
                        data.get("date", "N/A"), data.get("alert", status_key), data.get("insurance", ""),
                        data.get("driver", "Unknown"), data.get("cost", ""), data.get("acq_date", ""),
                        data.get("sheet", "Unknown"), data.get("sheet", "Unknown")
                    ]
                    all_rows.append((row_data, fg_color))
                    
        self.table.setRowCount(len(all_rows))
        for row_idx, (row_data, fg_color) in enumerate(all_rows):
            for col_idx, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                item.setForeground(QBrush(fg_color))
                # Text alignment
                if columns[col_idx] in ["INSURANCE (₱)", "ACQ. COST (₱)"]:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                elif columns[col_idx] in ["STATUS (YES/NO)", "OFFICE", "LAST REG.", "REMINDER", "DATE ACQUIRED", "MONTH"]:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)
                
        self.table.resizeRowsToContents()
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        # Approximate column widths
        self.table.setColumnWidth(0, 120); self.table.setColumnWidth(1, 80); self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 120); self.table.setColumnWidth(4, 120); self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 100); self.table.setColumnWidth(7, 100); self.table.setColumnWidth(8, 100)
        self.table.setColumnWidth(9, 100); self.table.setColumnWidth(10, 160); self.table.setColumnWidth(11, 120)
        self.table.setColumnWidth(12, 120); self.table.setColumnWidth(13, 120); self.table.setColumnWidth(14, 100)
        self.table.setColumnWidth(15, 100)

    def do_scan_all(self):
        global current_viewed_sheet
        current_viewed_sheet = None
        self.status_lbl.setText("Scanning all sheets in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE, None, True), daemon=True).start()

    def do_scan_month(self, selection):
        global current_viewed_sheet
        current_viewed_sheet = selection
        self.status_lbl.setText(f"Scanning {selection} in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE, selection, True), daemon=True).start()

    def on_row_click(self, item):
        row = item.row()
        sheet_to_open = self.table.item(row, self.table.columnCount()-1).text()
        
        # Debounce multiple clicks
        current_time = time.time()
        if hasattr(self, 'last_click_time') and current_time - self.last_click_time < 2.0: return
        self.last_click_time = current_time
        
        def open_excel_threaded():
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                abs_path = os.path.abspath(EXCEL_FILE)
                excel = None; wb = None
                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                    for w in excel.Workbooks:
                        if w.FullName.lower() == abs_path.lower(): wb = w; break
                except: pass
                    
                if not wb:
                    os.startfile(abs_path)
                    time.sleep(2.5) 
                    try:
                        excel = win32com.client.GetActiveObject("Excel.Application")
                        for w in excel.Workbooks:
                            if w.FullName.lower() == abs_path.lower(): wb = w; break
                    except: pass
                        
                if wb:
                    try:
                        if sheet_to_open and sheet_to_open != "Unknown":
                            for sh in wb.Sheets:
                                if sh.Name == sheet_to_open: sh.Activate(); break
                    except: pass
                    try:
                        excel.Visible = True
                        import win32gui, win32con
                        hwnd = excel.Hwnd
                        if hwnd:
                            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                    except: pass
                pythoncom.CoUninitialize()
            except Exception as e:
                try: os.startfile(os.path.abspath(EXCEL_FILE))
                except: pass
        threading.Thread(target=open_excel_threaded, daemon=True).start()
