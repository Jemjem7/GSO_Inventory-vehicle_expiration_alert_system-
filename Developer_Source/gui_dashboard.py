from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QComboBox, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QFrame,
                             QStackedWidget, QScrollArea, QSizePolicy, QGridLayout)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QSize, QRectF
from PyQt6.QtGui import QIcon, QFont, QColor, QBrush, QPixmap, QPainter, QTransform, QCursor
import threading
import time
import queue
import json
import os
import sys
import winsound
from datetime import datetime
import winreg

# We assume get_system_theme, clean_currency, app_settings, gui_queue, save_settings, EXCEL_FILE, process_excel, current_sheets, sort_sheets_chronologically exist in global scope

class ClickableLabel(QLabel):
    clicked = pyqtSignal()
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)

class AlertWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("⚠ Vehicle Expiration Alert Dashboard")
        self.last_alerts = {}
        self.last_title = ""
        self.first_popup_sound_played = False
        
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
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        self.stacked_widget = QStackedWidget()
        self.main_layout.addWidget(self.stacked_widget)
        
        # --- PAGE 1: DASHBOARD ---
        self.page_dashboard = QWidget()
        self.dashboard_layout = QHBoxLayout(self.page_dashboard)
        self.dashboard_layout.setContentsMargins(0, 0, 0, 0)
        self.dashboard_layout.setSpacing(0)
        
        # Left Sidebar (White strip)
        self.left_sidebar = QFrame()
        self.left_sidebar.setObjectName("LeftBar")
        self.left_sidebar.setFixedWidth(60)
        self.left_layout = QVBoxLayout(self.left_sidebar)
        self.left_layout.setContentsMargins(2, 20, 2, 20)
        self.left_layout.setSpacing(20)
        
        self.theme_lbl = QLabel("THEME")
        self.theme_lbl.setObjectName("SidebarLbl")
        self.theme_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Grain", "Nature", "Light", "Dark", "System"])
        idx = self.theme_combo.findText(self.current_theme)
        if idx >= 0: self.theme_combo.setCurrentIndex(idx)
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        
        self.scan_lbl = QLabel("MANUAL")
        self.scan_lbl.setObjectName("SidebarLbl")
        self.scan_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scan_all_btn = QPushButton("SCAN\nALL")
        self.scan_all_btn.setObjectName("SidebarBtn")
        self.scan_all_btn.clicked.connect(self.do_scan_all)
        
        self.left_layout.addWidget(self.theme_lbl)
        self.left_layout.addWidget(self.theme_combo)
        self.left_layout.addStretch()
        self.left_layout.addWidget(self.scan_lbl)
        self.left_layout.addWidget(self.scan_all_btn)
        self.left_layout.addStretch()
        
        # Middle Menu (Months)
        self.mid_panel = QFrame()
        self.mid_panel.setObjectName("MidPanel")
        self.mid_layout = QVBoxLayout(self.mid_panel)
        self.mid_layout.setContentsMargins(40, 20, 20, 20)
        
        self.time_lbl = QLabel("TIME")
        self.time_lbl.setObjectName("TimeHeader")
        self.mid_layout.addWidget(self.time_lbl, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.month_scroll = QScrollArea()
        self.month_scroll.setWidgetResizable(True)
        self.month_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.month_scroll.setStyleSheet("background: transparent; border: none;")
        self.month_scroll_content = QWidget()
        self.month_scroll_content.setStyleSheet("background: transparent;")
        self.month_list_layout = QVBoxLayout(self.month_scroll_content)
        self.month_list_layout.setContentsMargins(50, 20, 50, 20)
        self.month_list_layout.setSpacing(15)
        self.month_list_layout.addStretch() # Push to top
        self.month_scroll.setWidget(self.month_scroll_content)
        self.mid_layout.addWidget(self.month_scroll, 1)
        
        # Right Stats Structure
        self.right_panel = QFrame()
        self.right_panel.setObjectName("RightPanel")
        self.right_layout = QVBoxLayout(self.right_panel)
        self.right_layout.setContentsMargins(20, 20, 40, 20)
        
        self.mmddyy_lbl = QLabel("MM/DAY/YEAR")
        self.mmddyy_lbl.setObjectName("TimeHeader")
        self.right_layout.addWidget(self.mmddyy_lbl, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        self.stats_container = QWidget()
        self.stats_grid = QVBoxLayout(self.stats_container)
        self.stats_grid.setContentsMargins(0, 0, 0, 0)
        self.stats_grid.setSpacing(10)
        
        self.stat_labels = {}
        self.stat_categories = [
            ("TOTAL EXPIRED", "LESS THAN 0 DAYS = EXPIRED (RED)", "EXPIRED"),
            ("DAYS BEFORE EXPIRY", "1 TO 14 DAYS = DAYS BEFORE EXPIRY (ORANGE)", "DAYS BEFORE EXPIRY"),
            ("DAYS BEFORE 2 WEEK NOTICE", "15 TO 29 DAYS = DAYS BEFORE 2 WEEK NOTICE (YELLOW)", "2 WEEK NOTICE"),
            ("SUFFICIENT TIME", "30 DAYS AND MORE = SUFFICIENT TIME (GREEN)", "SUFFICIENT TIME"),
            ("PLEASE INPUT LAST REG.", "PLEASE INPUT LAST REG (GRAY)", "PLEASE INPUT LAST REG"),
            ("REGISTERED", "REGISTERED (BLUE)", "REGISTERED")
        ]
        
        for disp_name, parse_name, _ in self.stat_categories:
            group = QWidget()
            glay = QVBoxLayout(group)
            glay.setContentsMargins(0,0,0,0)
            glay.setSpacing(0)
            
            num_lbl = ClickableLabel("00")
            num_lbl.setObjectName("StatNum")
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_lbl.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            
            txt_lbl = ClickableLabel(disp_name)
            txt_lbl.setObjectName("StatTxt")
            txt_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            txt_lbl.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            
            glay.addWidget(num_lbl)
            glay.addWidget(txt_lbl)
            
            # Click events
            callback = (lambda cat=parse_name: lambda: self.show_table_view(filter_status=cat))()
            num_lbl.clicked.connect(callback)
            txt_lbl.clicked.connect(callback)
            
            self.stat_labels[parse_name] = num_lbl
            self.stats_grid.addWidget(group)
            self.stats_grid.addStretch()

        self.right_layout.addWidget(self.stats_container, 1)
        
        self.dashboard_layout.addWidget(self.left_sidebar, 0)
        self.dashboard_layout.addWidget(self.mid_panel, 1)
        self.dashboard_layout.addWidget(self.right_panel, 1)
        
        # --- PAGE 2: TABLE VIEW ---
        self.page_table = QWidget()
        self.page_table.setObjectName("TablePage")
        self.table_main_layout = QVBoxLayout(self.page_table)
        self.table_main_layout.setContentsMargins(0,0,0,0)
        self.table_main_layout.setSpacing(0)
        
        self.table_top_bar = QFrame()
        self.table_top_bar.setObjectName("TableTopBar")
        self.table_top_layout = QHBoxLayout(self.table_top_bar)
        self.table_top_layout.setContentsMargins(20, 15, 20, 15)
        self.back_btn = QPushButton("◄ BACK TO DASHBOARD")
        self.back_btn.setObjectName("BackBtn")
        self.back_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.back_btn.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(0))
        
        self.table_title_lbl = QLabel("DETAIL VIEW")
        self.table_title_lbl.setObjectName("TableTitle")
        
        self.table_top_layout.addWidget(self.back_btn)
        self.table_top_layout.addStretch()
        self.table_top_layout.addWidget(self.table_title_lbl)
        self.table_top_layout.addStretch()
        
        self.table_main_layout.addWidget(self.table_top_bar)
        
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.itemClicked.connect(self.on_row_click)
        self.table_main_layout.addWidget(self.table, 1)

        self.stacked_widget.addWidget(self.page_dashboard)
        self.stacked_widget.addWidget(self.page_table)
        self.stacked_widget.setCurrentIndex(0) # Start Dashboard
        
        # Setup clock timer
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_clock)
        self.clock_timer.start(1000)
        self.update_clock()
        
        # Setup queue checker
        self.queue_timer = QTimer(self)
        self.queue_timer.timeout.connect(self.check_queue)
        self.queue_timer.start(200)

        self.apply_stylesheet()
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.hide()

    def update_clock(self):
        now = datetime.now()
        time_str = now.strftime("%I:%M %p")
        date_str = now.strftime("%b %d, %Y").upper()
        self.time_lbl.setText(time_str)
        self.mmddyy_lbl.setText(date_str)

    def closeEvent(self, event):
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
                    self.stacked_widget.setCurrentIndex(0) # force popup to dashboard
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
        self.apply_stylesheet()

    def apply_stylesheet(self):
        # Dashboard uses global taupe mockup style, Table view uses the old theme styles!
        
        actual_theme = get_system_theme() if self.current_theme == "System" else self.current_theme
        if actual_theme == "Dark":
            t_bg = "#121212"; t_fg = "#E0E0E0"; t_panel = "#1E1E1E"; t_accent = "#04395E"; t_header_c = "#FFFFFF"
            t_top_bg = "#0F0F0F"; t_row1 = "#262626"; t_row2 = "#2A2A2A"
        else:
            t_bg = "#F5F5F5"; t_fg = "#202124"; t_panel = "#FFFFFF"; t_accent = "#E3F2FD"; t_header_c = "#202124"
            t_top_bg = "#FFFFFF"; t_row1 = "#FFFFFF"; t_row2 = "#FAFAFA"

        self.setStyleSheet(f"""
            QWidget {{
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            
            QFrame#LeftBar {{
                background-color: #F8F9FA;
                border-right: 1px solid #D0D0D0;
            }}
            
            QFrame#MidPanel, QFrame#RightPanel, QWidget#page_dashboard {{
                background-color: #9A958E; /* Taupe dashboard background */
            }}
            
            QWidget#TablePage {{
                background-color: {t_bg};
            }}
            QFrame#TableTopBar {{
                background-color: {t_top_bg};
            }}
            
            QLabel#SidebarLbl {{
                font-size: 10px;
                color: #555555;
                font-weight: bold;
            }}
            QPushButton#SidebarBtn {{
                background-color: transparent;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                padding: 10px 5px;
                font-size: 10px;
                font-weight: bold;
                color: #333;
            }}
            QPushButton#SidebarBtn:hover {{
                background-color: #E0E0E0;
            }}
            QComboBox {{
                font-size: 10px;
                padding: 2px;
                border: 1px solid #CCC;
                max-width: 45px;
            }}
            
            QLabel#TimeHeader {{
                font-size: 16px;
                font-weight: 800;
                color: #1A1A1A;
            }}
            
            /* Month Button Pills */
            QPushButton.MonthPill {{
                background-color: #EAE6DF;
                border-radius: 6px;
                padding: 12px 20px;
                font-size: 15px;
                font-weight: bold;
                color: #FFFFFF;
                text-align: center;
                min-width: 140px;
                border: 2px solid transparent; /* default */
            }}
            QPushButton.MonthPill:hover {{
                background-color: #F5F1EA;
            }}
            QPushButton.MonthPill[has_current="true"] {{
                border-bottom: 2px solid #1A1A1A; /* Imitating the black line under active in mockup? */
            }}
            
            /* Stat Text */
            QLabel#StatNum {{
                font-size: 55px;
                font-weight: 900;
                font-style: italic;
                color: #FFFFFF;
            }}
            QLabel#StatTxt {{
                font-size: 14px;
                font-weight: 900;
                font-style: italic;
                color: #FFFFFF;
            }}
            
            /* Table Styling */
            QTableWidget {{
                background-color: {t_panel};
                alternate-background-color: {t_row2};
                color: {t_fg};
                border: none;
                gridline-color: transparent;
                selection-background-color: {t_accent};
                selection-color: {t_fg};
                font-size: 13px;
            }}
            QTableWidget::item {{
                padding: 5px;
            }}
            QHeaderView::section {{
                background-color: {t_top_bg};
                color: {t_header_c};
                padding: 10px;
                border: 1px solid {t_row1};
                font-size: 13px;
                font-weight: bold;
            }}
            QLabel#TableTitle {{
                font-size: 18px;
                font-weight: bold;
                color: {t_header_c};
            }}
            QPushButton#BackBtn {{
                background-color: transparent;
                color: {t_header_c};
                font-size: 14px;
                font-weight: bold;
                border: none;
            }}
            QPushButton#BackBtn:hover {{
                color: #1976D2;
            }}
        """)
        self.actual_theme = actual_theme

    def show_table_view(self, filter_status=None, filter_sheet=None):
        # Refresh the table with filter
        self.populate_table(self.last_alerts, filter_status, filter_sheet)
        self.stacked_widget.setCurrentIndex(1)
        title_str = "DETAIL VIEW"
        if filter_month: title_str += f" : {filter_month}"
        if filter_status: title_str += f" : {filter_status.split('=')[-1].strip()}"
        self.table_title_lbl.setText(title_str)

    def build_ui(self, detailed_alerts, window_title):
        self.last_alerts = detailed_alerts
        self.last_title = window_title
        
        # 1. Update the Left Panel Counters
        # Initialize counts
        counts = {c[1]: 0 for c in self.stat_categories}
        
        for full_status, plates in detailed_alerts.items():
            if isinstance(plates, list):
                # We map the specific status string keys accurately
                for disp, parse, _ in self.stat_categories:
                    if parse in full_status:
                        counts[parse] += len(plates)
        
        for disp, parse, _ in self.stat_categories:
            str_val = str(counts[parse]).zfill(2)
            if parse in self.stat_labels:
                self.stat_labels[parse].setText(str_val)

        # 2. Update the Month Pills
        # Clear existing
        for i in reversed(range(self.month_list_layout.count())): 
            widget = self.month_list_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
                
        global current_sheets
        if current_sheets:
            sorted_sheets = sort_sheets_chronologically(current_sheets)
            for sh in sorted_sheets:
                btn = QPushButton(str(sh).upper())
                btn.setProperty("class", "MonthPill")
                btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
                # Add shadow style text effect directly
                btn.clicked.connect((lambda s=sh: lambda: self.show_table_view(filter_sheet=s))())
                self.month_list_layout.addWidget(btn)
        self.month_list_layout.addStretch()
        
        # We don't populate the table immediately until requested to save performance unless they switch
        self.populate_table(detailed_alerts, None, None)

    def populate_table(self, detailed_alerts, filter_status=None, filter_sheet=None):
        colors_map = {}
        if self.actual_theme == "Dark":
            colors_map = {'EXPIRED': '#EF5350', 'DAYS BEFORE EXPIRY': '#FFA726', 'DAYS BEFORE 2 WEEK NOTICE': '#FFEE58', 'SUFFICIENT TIME': '#66BB6A', 'PLEASE INPUT LAST REG': '#9E9E9E', 'REGISTERED': '#4FC3F7'}
        else:
            colors_map = {'EXPIRED': '#D32F2F', 'DAYS BEFORE EXPIRY': '#F57C00', 'DAYS BEFORE 2 WEEK NOTICE': '#FBC02D', 'SUFFICIENT TIME': '#388E3C', 'PLEASE INPUT LAST REG': '#757575', 'REGISTERED': '#1976D2'}
            
        columns = ["STATUS (YES/NO)", "OFFICE", "PLATE #", "MAKE", "TYPE", "EMISSION", "GSIS", "LTO", "LAST REG.", "REMINDER", "ALERT", "INSURANCE (₱)", "DRIVER", "ACQ. COST (₱)", "DATE ACQUIRED", "MONTH", "Sheet_Hidden"]
        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        self.table.setColumnHidden(len(columns)-1, True)
        
        all_rows = []
        importance_order = [
            'EXPIRED', 'DAYS BEFORE EXPIRY', 'DAYS BEFORE 2 WEEK NOTICE', 
            'SUFFICIENT TIME', 'PLEASE INPUT LAST REG', 'REGISTERED'
        ]
        
        for status_key in importance_order:
            if filter_status and filter_status not in status_key and status_key not in filter_status:
                continue
                
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
                    
                    if filter_sheet and data.get("sheet", "") != filter_sheet:
                        continue
                        
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
        header.setStretchLastSection(True)
        self.table.setColumnWidth(0, 140); self.table.setColumnWidth(1, 100); self.table.setColumnWidth(2, 120)
        self.table.setColumnWidth(3, 140); self.table.setColumnWidth(4, 140); self.table.setColumnWidth(5, 120)
        self.table.setColumnWidth(6, 120); self.table.setColumnWidth(7, 120); self.table.setColumnWidth(8, 110)
        self.table.setColumnWidth(9, 110); self.table.setColumnWidth(10, 200); self.table.setColumnWidth(11, 140)
        self.table.setColumnWidth(12, 140); self.table.setColumnWidth(13, 140); self.table.setColumnWidth(14, 120)
        self.table.setColumnWidth(15, 120)

    def do_scan_all(self):
        global current_viewed_sheet
        current_viewed_sheet = None
        self.table_title_lbl.setText("Scanning all sheets in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE, None, True), daemon=True).start()

    def do_scan_month(self, selection):
        global current_viewed_sheet
        current_viewed_sheet = selection
        self.table_title_lbl.setText(f"Scanning {selection} in background...")
        threading.Thread(target=process_excel, args=(EXCEL_FILE, selection, True), daemon=True).start()

    def on_row_click(self, item):
        row = item.row()
        sheet_to_open = self.table.item(row, self.table.columnCount()-1).text()
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
