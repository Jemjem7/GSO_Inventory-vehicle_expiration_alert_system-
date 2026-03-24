import io
import json
import os
import queue
import socket
import sys
import threading
import time
import traceback
import winreg
import winsound
from datetime import datetime

import colorama
import pandas as pd
import pystray
import win32com.client
from colorama import Fore, Style
from PIL import Image, ImageDraw
from PyQt6.QtCore import (
    QEasingCurve,
    QPoint,
    QPropertyAnimation,
    QRectF,
    QSize,
    Qt,
    QTimer,
    pyqtSignal,
)
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QCursor,
    QFont,
    QIcon,
    QPainter,
    QPixmap,
    QTransform,
)
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFrame,
    QGraphicsDropShadowEffect,
    QGraphicsOpacityEffect,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

class SafeStream:
    def __init__(self, original_stream):
        self._original = original_stream

    def write(self, s):
        try:
            if self._original:
                self._original.write(s)
        except Exception:
            pass

    def flush(self):
        try:
            if self._original:
                self._original.flush()
        except Exception:
            pass


sys.stdout = SafeStream(sys.stdout)
sys.stderr = SafeStream(sys.stderr)

colorama.init(autoreset=True)

# Configuration
EXCEL_FILE = "VehicleMonitoring.xlsx"
CHECK_INTERVAL_SECONDS = 5


def get_system_theme():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(
            registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
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
app_settings["theme"] = app_settings.get("theme", "Monokai")

# State
previous_state = {}
first_run = True
current_sheets = []
monitor_active = True
tray_icon = None
current_viewed_sheet = None


def sort_sheets_chronologically(sheet_list):
    def sort_key(s):
        try:
            # Handle cases like 'APRIL 2024' or 'JAN 1' (which might mean Jan 2024 or Jan 2025)
            # Default to current year if year not specified
            parts = str(s).strip().split()
            if not parts:
                return (9999, 99)  # fallback

            month_str = parts[0][:3].lower()  # match 'apr'
            months = [
                "jan",
                "feb",
                "mar",
                "apr",
                "may",
                "jun",
                "jul",
                "aug",
                "sep",
                "oct",
                "nov",
                "dec",
            ]
            m_idx = months.index(month_str) if month_str in months else 99

            y = datetime.now().year
            if len(parts) > 1:
                try:
                    # check if the second part is a year (e.g. 2024). if it's '1', maybe it just means day 1, or 202x?
                    year_cand = int(parts[1])
                    if year_cand > 1900:
                        y = year_cand
                    # if it's small e.g. 'JAN 1', we assume current year.
                except:
                    pass
            return (y, m_idx)
        except:
            return (9999, 99)  # put unknowns at the end

    return sorted(sheet_list, key=sort_key)
def backup_excel(filepath):
    try:
        if not os.path.exists(filepath):
            return
        backup_dir = "backup_excel_files"
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            
        import shutil
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)
        backup_path = os.path.join(backup_dir, f"{name}_{timestamp}{ext}")
        
        shutil.copy2(filepath, backup_path)
        
        # Cleanup old backups (keep last 10)
        backups = sorted([os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.startswith(name)], key=os.path.getmtime)
        while len(backups) > 10:
            try: os.remove(backups.pop(0))
            except: pass
            
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Backup Created: {backup_path}")
    except Exception as e:
        print(f"[Backup Error] {e}")


gui_queue = queue.Queue()



def create_image():
    width = 64
    height = 64
    image = Image.new("RGB", (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        (width // 4, height // 4, width * 3 // 4, height * 3 // 4), fill=(0, 120, 215)
    )  # Blue square
    dc.text((width // 3 + 2, height // 3 + 5), "V", fill=(255, 255, 255))
    return image


def print_status(message, status_col=""):
    color = Style.RESET_ALL
    if "EXPIRED" in status_col:
        color = Fore.RED
    elif "DAYS BEFORE EXPIRY" in status_col:
        color = Fore.YELLOW
    elif "2 WEEK NOTICE" in status_col:
        color = Fore.LIGHTYELLOW_EX
    elif "SUFFICIENT TIME" in status_col:
        color = Fore.GREEN
    elif "PLEASE INPUT LAST REG" in status_col:
        color = Fore.LIGHTBLACK_EX
    elif "REGISTERED" in status_col:
        color = Fore.CYAN

    print(f"{color}{message}{Style.RESET_ALL}")
    
    try:
        gui_queue.put({"type": "log", "message": message, "status": status_col})
    except:
        pass



def get_expiration_status(exp_date, status_override):
    # Fallback status generator if the user's Excel sheet doesn't calculate the ALERT column
    if pd.notna(status_override) and str(status_override).strip().upper() in [
        "YES",
        "REGISTERED",
    ]:
        return "REGISTERED"
    if pd.isna(exp_date) or str(exp_date).strip() == "":
        return "PLEASE INPUT LAST REG"
    try:
        if isinstance(exp_date, pd.Timestamp) or isinstance(exp_date, datetime):
            target_date = exp_date.date()
        else:
            exp_date_str = str(exp_date).replace("\\", "/")
            target_date = pd.to_datetime(exp_date_str, dayfirst=False).date()

        today = datetime.now().date()
        delta_days = (target_date - today).days

        if delta_days < 0:
            return "LESS THAN 0 DAYS = EXPIRED"
        elif 0 <= delta_days <= 14:
            return "1 TO 14 DAYS = DAYS BEFORE EXPIRY"
        elif 15 <= delta_days <= 29:
            return "15 TO 29 DAYS = DAYS BEFORE 2 WEEK NOTICE"
        else:
            return "30 DAYS AND MORE = SUFFICIENT TIME"
    except Exception as e:
        return "PLEASE INPUT LAST REG"


def clean_currency(val):
    import re

    if not val or pd.isna(val) or str(val).strip() == "":
        return ""

    val_str = str(val).strip()
    lines = val_str.split("\n")
    out_lines = []

    for line in lines:
        if not line.strip():
            continue
        clean_num = line.replace("₱", "").replace("P", "").replace(",", "").strip()
        try:
            numeric_val = float(clean_num)
            out_lines.append(f"{numeric_val:,.2f}  ")
        except ValueError:
            cln = re.sub(r"\s+", " ", line).replace("₱", "").replace("P", "").strip()
            out_lines.append(cln + "  ")

    return "\n".join(out_lines)


class ClickableLabel(QLabel):
    clicked = pyqtSignal()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)


class MonthButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setObjectName("MonthPill")
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setMinimumHeight(45)




class AlertWindow(QMainWindow):


    def __init__(self):
        super().__init__()
        self.setWindowTitle("⚠ Vehicle Expiration Alert Dashboard")
        self.last_alerts = {}
        self.last_title = ""
        self.first_popup_sound_played = False
        self.current_filter_status = None
        self.current_filter_sheet = None

        self.current_theme = app_settings.get("theme", "Monokai")

        try:
            if hasattr(sys, "_MEIPASS"):
                icon_path = os.path.join(sys._MEIPASS, "excel_scan_v3_final.ico")
            else:
                icon_path = os.path.abspath("excel_scan_v3_final.ico")
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
            import ctypes

            myappid = "localgov.gso.vehiclemonitor.1"
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
        self.page_dashboard_layout = QVBoxLayout(self.page_dashboard)
        self.page_dashboard_layout.setContentsMargins(0, 0, 0, 0)
        self.page_dashboard_layout.setSpacing(0)

        self.dashboard_top_content = QWidget()
        self.dashboard_layout = QHBoxLayout(self.dashboard_top_content)
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
        self.theme_combo.addItems(["Monokai", "Grain", "Nature", "Light", "Dark", "OLED Black", "System"])

        idx = self.theme_combo.findText(self.current_theme)
        if idx >= 0:
            self.theme_combo.setCurrentIndex(idx)
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
        self.mid_layout.setContentsMargins(40, 20, 0, 20)


        self.time_lbl = QLabel("TIME")
        self.time_lbl.setObjectName("TimeHeader")
        self.mid_layout.addWidget(
            self.time_lbl, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )

        self.month_scroll = QScrollArea()
        self.month_scroll.setWidgetResizable(True)
        self.month_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.month_scroll.setStyleSheet("background: transparent; border: none;")
        self.month_scroll_content = QWidget()
        self.month_scroll_content.setStyleSheet("background: transparent;")
        self.month_list_layout = QVBoxLayout(self.month_scroll_content)
        self.month_list_layout.setContentsMargins(50, 20, 50, 20)
        self.month_list_layout.setSpacing(15)
        self.month_list_layout.addStretch()  # Push to top
        self.month_scroll.setWidget(self.month_scroll_content)
        self.mid_layout.addWidget(self.month_scroll, 1)

        # Right Stats Structure
        self.right_panel = QFrame()
        self.right_panel.setObjectName("RightPanel")
        self.right_layout = QVBoxLayout(self.right_panel)
        self.right_layout.setContentsMargins(20, 20, 40, 20)

        self.mmddyy_lbl = QLabel("MM/DAY/YEAR")
        self.mmddyy_lbl.setObjectName("TimeHeader")
        self.right_layout.addWidget(
            self.mmddyy_lbl, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight
        )

        self.stats_container = QWidget()
        self.stats_grid = QVBoxLayout(self.stats_container)
        self.stats_grid.setContentsMargins(0, 0, 0, 0)
        self.stats_grid.setSpacing(10)

        self.stat_labels = {}
        self.stat_categories = [
            ("TOTAL EXPIRED", "LESS THAN 0 DAYS = EXPIRED", "EXPIRED"),
            (
                "DAYS BEFORE EXPIRY",
                "1 TO 14 DAYS = DAYS BEFORE EXPIRY",
                "DAYS BEFORE EXPIRY",
            ),
            (
                "DAYS BEFORE 2 WEEK NOTICE",
                "15 TO 29 DAYS = DAYS BEFORE 2 WEEK NOTICE",
                "2 WEEK NOTICE",
            ),
            (
                "SUFFICIENT TIME",
                "30 DAYS AND MORE = SUFFICIENT TIME",
                "SUFFICIENT TIME",
            ),
            (
                "PLEASE INPUT LAST REG.",
                "PLEASE INPUT LAST REG",
                "PLEASE INPUT LAST REG",
            ),
            ("REGISTERED", "REGISTERED", "REGISTERED"),
        ]

        for disp_name, parse_name, short_key in self.stat_categories:
            group = QWidget()
            group.setMinimumHeight(65)
            
            cat_map = {
                "EXPIRED": "expired",
                "DAYS BEFORE EXPIRY": "days_before_expiry",
                "2 WEEK NOTICE": "notice_2week",
                "SUFFICIENT TIME": "sufficient_time",
                "PLEASE INPUT LAST REG": "input_reg",
                "REGISTERED": "registered"
            }
            group.setProperty("stat_cat", cat_map.get(short_key, "unknown"))

            glay = QVBoxLayout(group)
            glay.setContentsMargins(0, 4, 0, 4)
            glay.setSpacing(8)

            num_lbl = ClickableLabel("--")
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
            callback = (
                lambda cat=parse_name: lambda: self.show_table_view(filter_status=cat)
            )()
            num_lbl.clicked.connect(callback)
            txt_lbl.clicked.connect(callback)

            self.stat_labels[parse_name] = num_lbl
            self.stats_grid.addWidget(group)

        # Single stretch at the bottom of the stats grid so stats pack together at the top
        self.stats_grid.addStretch()

        self.right_layout.addWidget(self.stats_container)

        # Bottom stretch to anchor items to top and keep layout locked upon parent scaling
        self.right_layout.addStretch()




        self.dashboard_layout.addWidget(self.left_sidebar, 0)
        self.dashboard_layout.addWidget(self.mid_panel, 1)
        self.dashboard_layout.addWidget(self.right_panel, 1)


        # Bottom Log Frame
        from PyQt6.QtWidgets import QListWidget, QAbstractItemView
        self.log_widget = QFrame()
        self.log_widget.setObjectName("LogWidget")
        self.log_widget.setFixedHeight(120)
        self.log_layout = QVBoxLayout(self.log_widget)
        self.log_layout.setContentsMargins(15, 5, 15, 5)
        self.log_layout.setSpacing(3)

        self.log_header = QWidget()
        self.log_header_layout = QHBoxLayout(self.log_header)
        self.log_header_layout.setContentsMargins(0, 0, 0, 0)

        self.log_title = QLabel("SYSTEM ACTIVITY LOGS")
        self.log_title.setStyleSheet("font-size: 11px; font-weight: bold; color: #777777;")
        
        self.log_toggle_btn = QPushButton("HIDE LOGS")
        self.log_toggle_btn.setFixedWidth(85)
        self.log_toggle_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.log_toggle_btn.setStyleSheet("""
            QPushButton {
                font-size: 9px; font-weight: bold; color: #666666;
                background: #FFFFFF; border: 1px solid #CCCCCC;
                border-radius: 4px; padding: 2px 6px;
            }
            QPushButton:hover { background: #EEEEEE; }
        """)
        self.log_toggle_btn.clicked.connect(self.toggle_logs)

        self.log_header_layout.addWidget(self.log_title)
        self.log_header_layout.addStretch()
        self.log_header_layout.addWidget(self.log_toggle_btn)

        self.log_list = QListWidget()
        self.log_list.setObjectName("LogList")
        self.log_list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.log_list.setFrameShape(QFrame.Shape.NoFrame)
        self.log_list.setStyleSheet("background: transparent; font-size: 11px; color: #666666;")
        
        self.log_layout.addWidget(self.log_header)
        self.log_layout.addWidget(self.log_list)


        self.page_dashboard_layout.addWidget(self.dashboard_top_content, 1)
        self.page_dashboard_layout.addWidget(self.log_widget, 0)


        # --- PAGE 2: TABLE VIEW ---
        self.page_table = QWidget()
        self.page_table.setObjectName("TablePage")
        self.table_main_layout = QVBoxLayout(self.page_table)
        self.table_main_layout.setContentsMargins(0, 0, 0, 0)
        self.table_main_layout.setSpacing(0)

        self.table_top_bar = QFrame()
        self.table_top_bar.setObjectName("TableTopBar")
        self.table_top_layout = QHBoxLayout(self.table_top_bar)
        self.table_top_layout.setContentsMargins(20, 15, 20, 15)
        self.back_btn = QPushButton("◄ BACK TO DASHBOARD")
        self.back_btn.setObjectName("BackBtn")
        self.back_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.back_btn.clicked.connect(lambda: self.animate_to_page(0))

        self.table_title_lbl = QLabel("LOADING EXCEL DATA... PLEASE WAIT")
        self.table_title_lbl.setObjectName("TableTitle")

        from PyQt6.QtWidgets import QLineEdit
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("🔍 Search table (Plate, Office, Driver...)")
        self.search_bar.setFixedWidth(260)
        self.search_bar.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                border-radius: 12px;
                padding: 4px 10px;
                font-size: 11px;
                color: #333333;
            }
            QLineEdit:focus {
                border: 1px solid #1A1A1A;
            }
        """)
        self.search_bar.textChanged.connect(self.filter_table_by_search)

        self.table_top_layout.addWidget(self.back_btn)
        self.table_top_layout.addStretch()
        self.table_top_layout.addWidget(self.table_title_lbl)
        self.table_top_layout.addStretch()
        self.table_top_layout.addWidget(self.search_bar)

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
        self.stacked_widget.setCurrentIndex(0)  # Start Dashboard

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
        
        self.table_title_lbl.setText("LOADING EXCEL DATA... PLEASE WAIT...")
        
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
        )
        self.showNormal()
        QTimer.singleShot(100, self.do_initial_startup)

    def do_initial_startup(self):
        try:
            backup_excel(EXCEL_FILE)
            process_excel(EXCEL_FILE)
        except Exception as e:
            print(e)
        self.table_title_lbl.setText("DETAIL VIEW")
        
        monitor_thread = threading.Thread(target=background_monitor, daemon=True)
        monitor_thread.start()


    def toggle_logs(self):
        if self.log_list.isVisible():
            self.log_list.hide()
            self.log_widget.setFixedHeight(35)
            self.log_toggle_btn.setText("SHOW LOGS")
        else:
            self.log_list.show()
            self.log_widget.setFixedHeight(120)
            self.log_toggle_btn.setText("HIDE LOGS")


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
                if msg["type"] == "show":
                    if msg.get("is_auto", False) and not self.first_popup_sound_played:
                        self.first_popup_sound_played = True

                        def play_alert():
                            try:
                                winsound.Beep(1200, 300)
                                winsound.Beep(800, 200)
                            except:
                                pass

                        threading.Thread(target=play_alert, daemon=True).start()

                    was_visible = self.isVisible()
                    was_minimized = self.isMinimized()
                    self.build_ui(msg["alerts"], msg["title"])

                    if not was_visible or was_minimized:
                        if not was_visible:
                            # Popup size (not full screen)
                            self.resize(500, 700)
                            # Position in bottom-right corner of screen (near system tray)
                            screen = QApplication.primaryScreen()
                            sr = screen.availableGeometry()
                            w, h = self.width(), self.height()
                            x = max(sr.left(), sr.right() - w - 20)
                            y = max(sr.top(), sr.bottom() - h - 20)
                            self.setGeometry(x, y, w, h)
                            self.showNormal()
                            self.stacked_widget.setCurrentIndex(0)
                        elif was_minimized:
                            self.showNormal()
                        
                        self.raise_()
                        self.activateWindow()
                elif msg["type"] == "show_request":
                    self.showNormal()
                    self.raise_()
                    self.activateWindow()
                elif msg["type"] == "log":
                    from PyQt6.QtWidgets import QListWidgetItem
                    itm = QListWidgetItem(f"[{datetime.now().strftime('%I:%M:%S %p')}] {msg['message']}")
                    colors = {
                        "EXPIRED": "#EF5350",
                        "DAYS BEFORE EXPIRY": "#FF8C00",
                        "2 WEEK NOTICE": "#FFEE58",
                        "SUFFICIENT TIME": "#66BB6A",
                        "PLEASE INPUT LAST REG": "#9E9E9E",
                        "REGISTERED": "#4FC3F7",
                    }

                    if msg.get("status") in colors:
                        itm.setForeground(QColor(colors[msg["status"]]))
                    self.log_list.insertItem(0, itm)
                    if self.log_list.count() > 30:
                        self.log_list.takeItem(30)
                elif msg["type"] == "exit":

                    QApplication.quit()
                    return
        except queue.Empty:
            pass

    def animate_to_page(self, index, callback=None):
        if self.stacked_widget.currentIndex() == index:
            if callback:
                callback()
            return

        # 1. Capture current widget as a pixmap for the "sliding out" effect
        current_widget = self.stacked_widget.currentWidget()
        pixmap = current_widget.grab()
        
        # Create an overlay to show the old screen sliding away
        self.animation_overlay = QLabel(self)
        self.animation_overlay.setPixmap(pixmap)
        self.animation_overlay.setGeometry(self.stacked_widget.geometry())
        self.animation_overlay.show()
        self.animation_overlay.raise_()

        # 2. Prepare the NEXT widget
        is_forward = index > self.stacked_widget.currentIndex()
        if callback:
            callback()
        
        self.stacked_widget.setCurrentIndex(index)
        next_widget = self.stacked_widget.currentWidget()
        
        # 3. Setup Animations
        offset = self.width() if is_forward else -self.width()
        
        # Old screen slides OUT
        self.anim_old = QPropertyAnimation(self.animation_overlay, b"pos")
        self.anim_old.setDuration(450)
        self.anim_old.setStartValue(self.animation_overlay.pos())
        self.anim_old.setEndValue(QPoint(self.animation_overlay.x() - offset, self.animation_overlay.y()))
        self.anim_old.setEasingCurve(QEasingCurve.Type.OutCubic)
        
        # New screen slides IN
        self.anim_new = QPropertyAnimation(next_widget, b"pos")
        self.anim_new.setDuration(450)
        self.anim_new.setStartValue(QPoint(offset, 0))
        self.anim_new.setEndValue(QPoint(0, 0))
        self.anim_new.setEasingCurve(QEasingCurve.Type.OutCubic)
        
        # Cleanup after animation
        self.anim_old.finished.connect(self.animation_overlay.deleteLater)
        
        self.anim_old.start()
        self.anim_new.start()

    def change_theme(self, selection):
        self.current_theme = selection
        app_settings["theme"] = selection
        save_settings(app_settings)
        self.apply_stylesheet()
        
        # Force refresh components that use fixed colors or need re-calculation
        if hasattr(self, "last_alerts") and self.last_alerts:
            self.build_ui(self.last_alerts, self.last_title)
            # If we are in the table view, re-populate it too
            if self.stacked_widget.currentIndex() == 1:
                # We need to preserve filters if possible or just reset
                self.populate_table(self.last_alerts)

    def apply_stylesheet(self):
        # Map dynamic dashboard and table colors
        theme_map = {
            "OLED Black": {
                "dash_bg": "#000000",
                "pill_bg": "#111111",
                "pill_hover": "#222222",
                "pill_fg": "#FFFFFF",
                "stat_fg": "#FFFFFF",
                "header_fg": "#FFFFFF",
                "table_bg": "#000000",
                "table_panel": "#111111",
                "table_fg": "#E0E0E0",
                "accent": "#1A1A1A"
            },

            "Monokai": {
                "dash_bg": "#1B1D1E",
                "pill_bg": "#333333",
                "pill_hover": "#444444",
                "pill_fg": "#D6D6D6",
                "stat_fg": "#E6DB74",
                "header_fg": "#A6E22E",
                "table_bg": "#1B1D1E",
                "table_panel": "#232526",
                "table_fg": "#D6D6D6",
                "accent": "#464646"
            },
            "Nature": {
                "dash_bg": "#9A958E",
                "pill_bg": "#EAE6DF",
                "pill_hover": "#F5F1EA",
                "pill_fg": "#FFFFFF",
                "stat_fg": "#FFFFFF",
                "header_fg": "#1A1A1A",
                "table_bg": "#F5F5F5",
                "table_panel": "#FFFFFF",
                "table_fg": "#202124",
                "accent": "#E3F2FD"
            },
            "Grain": {
                "dash_bg": "#C2B280",
                "pill_bg": "#D2C29D",
                "pill_hover": "#E2D2BD",
                "pill_fg": "#FFFFFF",
                "stat_fg": "#FFFFFF",
                "header_fg": "#1A1A1A",
                "table_bg": "#FAF9F6",
                "table_panel": "#FFFFFF",
                "table_fg": "#3E2723",
                "accent": "#D7CCC8"
            },
            "Dark": {
                "dash_bg": "#1A1A1A",
                "pill_bg": "#333333",
                "pill_hover": "#444444",
                "pill_fg": "#E0E0E0",
                "stat_fg": "#FFFFFF",
                "header_fg": "#FFFFFF",
                "table_bg": "#121212",
                "table_panel": "#1E1E1E",
                "table_fg": "#E0E0E0",
                "accent": "#04395E"
            },
            "Light": {
                "dash_bg": "#FFFFFF",
                "pill_bg": "#F0F0F0",
                "pill_hover": "#E0E0E0",
                "pill_fg": "#1A1A1A",
                "stat_fg": "#1A1A1A",
                "header_fg": "#1A1A1A",
                "table_bg": "#F8F9FA",
                "table_panel": "#FFFFFF",
                "table_fg": "#202124",
                "accent": "#E3F2FD"
            }
        }

        actual_theme = self.current_theme
        if actual_theme == "System":
            actual_theme = get_system_theme()
        
        # Fallback to Nature if theme not found
        t_cfg = theme_map.get(actual_theme, theme_map["Nature"])
        
        # For legacy compatibility with table populating
        t_bg = t_cfg["table_bg"]
        t_fg = t_cfg["table_fg"]
        t_panel = t_cfg["table_panel"]
        t_accent = t_cfg["accent"]
        t_top_bg = t_cfg["table_panel"]
        t_header_c = t_cfg["table_fg"]
        t_row1 = t_cfg["table_panel"]
        t_row2 = "#FAFAFA" if actual_theme != "Dark" else "#2A2A2A"

        self.setStyleSheet(f"""
            QWidget {{
                font-family: 'Segoe UI', Arial, sans-serif;
            }}

            QFrame#LeftBar {{
                background-color: #F8F9FA;
                border-right: 1px solid #D0D0D0;
            }}

            QFrame#MidPanel, QFrame#RightPanel, QWidget#page_dashboard {{
                background-color: {t_cfg['dash_bg']};
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
                color: {t_cfg['header_fg']};
            }}
            QPushButton#MonthPill {{
                background-color: {t_cfg['pill_bg']};
                border-radius: 6px;
                padding: 12px 20px;
                font-size: 15px;
                font-weight: bold;
                color: {t_cfg['pill_fg']};
                text-align: center;
                min-width: 140px;
                border: 2px solid transparent; /* default */
            }}
            QPushButton#MonthPill:hover {{
                background-color: {t_cfg['pill_hover']};
                border: 2px solid #1A1A1A;
            }}
            QPushButton#MonthPill[has_current="true"] {{
                border-bottom: 2px solid #1A1A1A; /* Imitating the black line under active in mockup? */
            }}

            /* Stat Text */
            QLabel#StatNum {{
                font-size: 55px;
                font-weight: 900;
                font-style: italic;
                color: {t_cfg['stat_fg']};
            }}
            QLabel#StatTxt {{
                font-size: 14px;
                font-weight: 900;
                font-style: italic;
                color: {t_cfg['stat_fg']};
            }}

            QWidget[stat_cat="expired"]:hover QLabel {{ color: #DC3545; }}
            QWidget[stat_cat="days_before_expiry"]:hover QLabel {{ color: #FD7E14; }}
            QWidget[stat_cat="notice_2week"]:hover QLabel {{ color: #FFC107; }}
            QWidget[stat_cat="sufficient_time"]:hover QLabel {{ color: #28A745; }}
            QWidget[stat_cat="input_reg"]:hover QLabel {{ color: #AAAAAA; }}
            QWidget[stat_cat="registered"]:hover QLabel {{ color: #007BFF; }}

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

            /* Scrollbar Styling */
            QScrollBar:vertical {{
                border: none;
                background: rgba(0, 0, 0, 0.05);
                width: 8px;
                margin: 0px;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background: #6D6963;
                min-height: 30px;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #1A1A1A;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
            }}
        """)
        self.actual_theme = actual_theme

    def show_table_view(self, filter_status=None, filter_sheet=None):
        self.current_filter_status = filter_status
        self.current_filter_sheet = filter_sheet
        
        def update_ui():
            self.populate_table(self.last_alerts, filter_status, filter_sheet)
            title_str = "DETAIL VIEW"
            if filter_sheet:
                title_str += f" : {filter_sheet}"
            if filter_status:
                title_str += f" : {filter_status.split('=')[-1].strip()}"
            self.table_title_lbl.setText(title_str)

        self.animate_to_page(1, update_ui)

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
        while self.month_list_layout.count():
            item = self.month_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        global current_sheets
        if current_sheets:
            sorted_sheets = sort_sheets_chronologically(current_sheets)
            for sh in sorted_sheets:
                btn = MonthButton(str(sh).upper())
                btn.clicked.connect(
                    (lambda s=sh: lambda: self.show_table_view(filter_sheet=s))()
                )
                self.month_list_layout.addWidget(btn)
        self.month_list_layout.addStretch()

        # Update table using current filters to preserve view
        self.populate_table(detailed_alerts, self.current_filter_status, self.current_filter_sheet)
        self.table_title_lbl.setText("DETAIL VIEW" if not self.current_filter_sheet else f"DETAIL VIEW : {self.current_filter_sheet}")
        self.scan_all_btn.setText("SCAN\nALL")
        self.scan_all_btn.setEnabled(True)

    def populate_table(self, detailed_alerts, filter_status=None, filter_sheet=None):
        colors_map = {}
        actual_theme = getattr(self, "actual_theme", "Light")
        if actual_theme == "Dark":
            colors_map = {
                "EXPIRED": "#EF5350",
                "DAYS BEFORE EXPIRY": "#FF8C00",
                "DAYS BEFORE 2 WEEK NOTICE": "#FFEE58",
                "SUFFICIENT TIME": "#66BB6A",
                "PLEASE INPUT LAST REG": "#9E9E9E",
                "REGISTERED": "#4FC3F7",
            }
        else:
            colors_map = {
                "EXPIRED": "#D32F2F",
                "DAYS BEFORE EXPIRY": "#F57C00",
                "DAYS BEFORE 2 WEEK NOTICE": "#FBC02D",
                "SUFFICIENT TIME": "#388E3C",
                "PLEASE INPUT LAST REG": "#757575",
                "REGISTERED": "#1976D2",
            }

        columns = [
            "STATUS (YES/NO)",
            "OFFICE",
            "PLATE #",
            "MAKE",
            "TYPE",
            "EMISSION",
            "GSIS",
            "LTO",
            "LAST REG.",
            "REMINDER",
            "ALERT",
            "INSURANCE (₱)",
            "DRIVER",
            "ACQ. COST (₱)",
            "DATE ACQUIRED",
            "MONTH",
            "Sheet_Hidden",
        ]
        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        self.table.setColumnHidden(len(columns) - 1, True)

        all_rows = []
        importance_order = [
            "EXPIRED",
            "DAYS BEFORE EXPIRY",
            "DAYS BEFORE 2 WEEK NOTICE",
            "SUFFICIENT TIME",
            "PLEASE INPUT LAST REG",
            "REGISTERED",
        ]

        for status_key in importance_order:
            if (
                filter_status
                and filter_status not in status_key
                and status_key not in filter_status
            ):
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
                        return datetime.strptime(data.get("date", "N/A"), "%Y-%m-%d")
                    except:
                        return datetime.max

                matching_plates.sort(key=extract_date)

                fg_hex = colors_map.get(
                    status_key,  "#FFFFFF" if actual_theme != "Light" else "#000000"
                )
                fg_color = QColor(fg_hex)

                for p_str in matching_plates:
                    try:
                        data = json.loads(p_str)
                    except:
                        data = {}

                    if filter_sheet and data.get("sheet", "") != filter_sheet:
                        continue

                    row_data = [
                        data.get("status", ""),
                        data.get("office", ""),
                        data.get("plate", "Unknown"),
                        data.get("make", ""),
                        data.get("type", ""),
                        data.get("emission", ""),
                        data.get("gsis", ""),
                        data.get("lto", ""),
                        data.get("last_reg", ""),
                        data.get("date", "N/A"),
                        data.get("alert", status_key),
                        data.get("insurance", ""),
                        data.get("driver", "Unknown"),
                        data.get("cost", ""),
                        data.get("acq_date", ""),
                        data.get("sheet", "Unknown"),
                        data.get("sheet", "Unknown"),
                    ]
                    all_rows.append((row_data, fg_color))

        self.table.setRowCount(len(all_rows))
        for row_idx, (row_data, fg_color) in enumerate(all_rows):
            for col_idx, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                item.setForeground(QBrush(fg_color))
                if columns[col_idx] in ["INSURANCE (₱)", "ACQ. COST (₱)"]:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                elif columns[col_idx] in [
                    "STATUS (YES/NO)",
                    "OFFICE",
                    "LAST REG.",
                    "REMINDER",
                    "DATE ACQUIRED",
                    "MONTH",
                ]:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                    )
                self.table.setItem(row_idx, col_idx, item)

        self.table.resizeRowsToContents()
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        self.table.setColumnWidth(0, 140)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 120)
        self.table.setColumnWidth(3, 140)
        self.table.setColumnWidth(4, 140)
        self.table.setColumnWidth(5, 120)
        self.table.setColumnWidth(6, 120)
        self.table.setColumnWidth(7, 120)
        self.table.setColumnWidth(8, 110)
        self.table.setColumnWidth(9, 110)
        self.table.setColumnWidth(10, 200)
        self.table.setColumnWidth(11, 140)
        self.table.setColumnWidth(12, 140)
        self.table.setColumnWidth(13, 140)
        self.table.setColumnWidth(14, 120)
        self.table.setColumnWidth(15, 120)

    def filter_table_by_search(self):
        search_text = self.search_bar.text().lower().strip()
        for row in range(self.table.rowCount()):
            match_found = False
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and search_text in item.text().lower():
                    match_found = True
                    break
            self.table.setRowHidden(row, not match_found)

    def do_scan_all(self):

        global current_viewed_sheet
        current_viewed_sheet = None
        self.current_filter_status = None
        self.current_filter_sheet = None
        self.scan_all_btn.setText("SCANNING...")
        self.scan_all_btn.setEnabled(False)
        self.table_title_lbl.setText("Scanning all sheets in background...")
        
        # Clear UI for visual refresh
        self.table.clearContents()
        self.table.setRowCount(0)
        while self.month_list_layout.count():
            item = self.month_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
                
        for disp, parse, _ in self.stat_categories:
            if parse in self.stat_labels:
                self.stat_labels[parse].setText("00")
                
        global previous_state
        previous_state.clear()
        
        def run_scan():
            try:
                process_excel(EXCEL_FILE, None, True)
            finally:
                # Reset button in main thread
                QTimer.singleShot(0, lambda: self.scan_all_btn.setText("SCAN\nALL"))
                QTimer.singleShot(0, lambda: self.scan_all_btn.setEnabled(True))
                
        threading.Thread(target=run_scan, daemon=True).start()

    def do_scan_month(self, selection):
        global current_viewed_sheet
        current_viewed_sheet = selection
        self.table_title_lbl.setText(f"Scanning {selection} in background...")
        
        global previous_state
        previous_state.clear()
        
        threading.Thread(
            target=process_excel, args=(EXCEL_FILE, selection, True), daemon=True
        ).start()

    def open_excel_at_sheet(self, sheet_to_open):
        current_time = time.time()
        if (
            hasattr(self, "last_click_time")
            and current_time - self.last_click_time < 2.0
        ):
            return
        self.last_click_time = current_time

        def open_excel_threaded():
            try:
                import pythoncom
                import win32com.client

                pythoncom.CoInitialize()
                abs_path = os.path.abspath(EXCEL_FILE)
                excel = None
                wb = None
                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                    for w in excel.Workbooks:
                        if w.FullName.lower() == abs_path.lower():
                            wb = w
                            break
                except:
                    pass
                if not wb:
                    os.startfile(abs_path)
                    for _ in range(12):  # Wait up to 6s
                        time.sleep(0.5)
                        try:
                            excel = win32com.client.GetActiveObject("Excel.Application")
                            for w in excel.Workbooks:
                                if w.FullName.lower() == abs_path.lower():
                                    wb = w
                                    break
                            if wb:
                                break
                        except:
                            pass
                if wb:
                    try:
                        if sheet_to_open and sheet_to_open != "Unknown":
                            print(f"Attempting to activate Excel sheet: '{sheet_to_open}'")
                            for sh in wb.Sheets:
                                if str(sh.Name).strip().upper() == str(sheet_to_open).strip().upper():
                                    sh.Activate()
                                    break
                    except:
                        pass
                    try:
                        excel.Visible = True
                        import win32con
                        import win32gui

                        hwnd = excel.Hwnd
                        if hwnd:
                            if win32gui.IsIconic(hwnd):
                                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                    except:
                        pass
                pythoncom.CoUninitialize()
            except Exception as e:
                try:
                    os.startfile(os.path.abspath(EXCEL_FILE))
                except:
                    pass

        threading.Thread(target=open_excel_threaded, daemon=True).start()

    def on_row_click(self, item):
        row = item.row()
        sheet_to_open = self.table.item(row, self.table.columnCount() - 1).text()
        self.open_excel_at_sheet(sheet_to_open)


def send_notification(
    detailed_alerts, title="⚠ Vehicle Update Detected", is_auto=False
):
    if not detailed_alerts:
        return
    gui_queue.put(
        {"type": "show", "alerts": detailed_alerts, "title": title, "is_auto": is_auto}
    )

    # Native Desktop Toast Alert for Auto/Background Changes
    if is_auto:
        def show_toast_threaded():
            try:
                import sys
                import os
                icon_p = "excel_scan_v3_final.ico"
                if hasattr(sys, "_MEIPASS"):
                    icon_p = os.path.join(sys._MEIPASS, "excel_scan_v3_final.ico")
                else:
                    icon_p = os.path.abspath("excel_scan_v3_final.ico")
                if not os.path.exists(icon_p):
                    icon_p = None # Fallback to default python or windows icon

                from win10toast import ToastNotifier
                toaster = ToastNotifier()
                total_count = sum(len(v) for v in detailed_alerts.values() if isinstance(v, list))
                
                # Check for critical counts (EXPIRED)
                expired_count = len(detailed_alerts.get("LESS THAN 0 DAYS = EXPIRED", []))
                msg = f"{total_count} vehicles flag thresholds."
                if expired_count > 0:
                    msg = f"CRITICAL: {expired_count} fully EXPIRED vehicles found!"

                toaster.show_toast(
                    title, 
                    msg, 
                    icon_path=icon_p, 
                    duration=6, 
                    threaded=False # inside threading.Thread already
                )
            except:
                pass
        threading.Thread(target=show_toast_threaded, daemon=True).start()



def format_plate_with_data(
    plate,
    exp_date,
    sheet_name="Unknown",
    owner="Unknown",
    office="",
    make="",
    type_val="",
    emission="",
    gsis="",
    lto="",
    last_reg="",
    cost="",
    acq_date="",
    phys_status="",
    alert="",
    insurance="",
    driver="",
):
    if pd.isna(exp_date) or str(exp_date).strip() == "":
        dt_str = "N/A"
    else:
        try:
            if not hasattr(exp_date, "strftime"):
                exp_date_str = str(exp_date).replace("\\", "/")
                exp_date = pd.to_datetime(exp_date_str, dayfirst=False)
            dt_str = exp_date.strftime("%Y-%m-%d")
        except:
            dt_str = str(exp_date)

    return json.dumps(
        {
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
            "alert": alert,
        }
    )


def find_header_row(excel_file_obj, sheet_name):
    """
    Scans the first 15 rows looking for "PLATE".
    Returns the integer index of the row to use as the header.
    """
    try:
        df_test = pd.read_excel(
            excel_file_obj, nrows=15, header=None, sheet_name=sheet_name
        )
        for i, row in df_test.iterrows():
            if any(isinstance(v, str) and "PLATE" in v.upper() for v in row.values):
                return i
    except:
        pass
    return 3  # fallback default


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
                with open(filepath, "rb") as f:
                    file_buffer = io.BytesIO(f.read())
                break
            except PermissionError as pe:
                if attempt < 3:
                    time.sleep(1)
                else:
                    raise pe

        # Load specific sheet or all sheets
        with pd.ExcelFile(file_buffer, engine="openpyxl") as xl:
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
                current_sheets = sort_sheets_chronologically(list(dfs.keys()))

    except Exception as e:
        if is_manual_scan:
            print(f"{Fore.RED}Error loading Excel: {e}{Style.RESET_ALL}")
        return False

    all_data = []

    for sheet_name, df_sheet in dfs.items():
        if df_sheet.empty:
            continue

        df_sheet.columns = (
            df_sheet.columns.astype(str).str.strip().str.replace("\n", " ")
        )

        # Find the dynamic columns
        plate_col_candidates = [
            c for c in df_sheet.columns if "PLATE" in str(c).upper()
        ]
        plate_col = plate_col_candidates[0] if plate_col_candidates else "PLATE #"

        owner_col_candidates = [
            c
            for c in df_sheet.columns
            if "NAME" in str(c).upper()
            or "OWNER" in str(c).upper()
            or "CUSTOMER" in str(c).upper()
            or "ACCOUNTABLE" in str(c).upper()
            or "PERSON" in str(c).upper()
        ]
        owner_col = owner_col_candidates[0] if owner_col_candidates else None

        exp_col_candidates = [
            c
            for c in df_sheet.columns
            if "REMINDER" in str(c).upper()
            or "EXPIRATION" in str(c).upper()
            or "EXPIRY" in str(c).upper()
            or ("DATE" in str(c).upper() and "ACQUISITION" not in str(c).upper())
        ]
        exp_col = exp_col_candidates[0] if exp_col_candidates else "REMINDER"

        status_col_keys = [
            c for c in df_sheet.columns if "REGISTERED" in str(c).upper()
        ]
        status_col = status_col_keys[0] if status_col_keys else None

        phys_status_keys = [
            c
            for c in df_sheet.columns
            if "YES" in str(c).upper() and "NOT" not in str(c).upper()
        ]
        phys_status_col = phys_status_keys[0] if phys_status_keys else None

        alert_col_candidates = [
            c
            for c in df_sheet.columns
            if "ALERT" in str(c).upper() and "SYSTEM" not in str(c).upper()
        ]
        alert_col = alert_col_candidates[0] if alert_col_candidates else None

        office_c = [c for c in df_sheet.columns if "OFFICE" in str(c).upper()]
        make_c = [c for c in df_sheet.columns if "MAKE" in str(c).upper()]
        type_c = [
            c
            for c in df_sheet.columns
            if "TYPE" in str(c).upper() and "BODY" not in str(c).upper()
        ]
        emission_c = [c for c in df_sheet.columns if "EMISSION" in str(c).upper()]
        gsis_c = [c for c in df_sheet.columns if "GSIS" in str(c).upper()]
        lto_c = [c for c in df_sheet.columns if "LTO" in str(c).upper()]
        last_reg_c = [c for c in df_sheet.columns if "LAST REG" in str(c).upper()]
        insurance_c = [c for c in df_sheet.columns if "INSURANCE" in str(c).upper()]
        cost_c = [c for c in df_sheet.columns if "COST" in str(c).upper()]
        acq_date_c = [
            c
            for c in df_sheet.columns
            if "ACQUIRED" in str(c).upper() or "ACQUISITION DATE" in str(c).upper()
        ]
        driver_c = [c for c in df_sheet.columns if "DRIVER" in str(c).upper()]

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
            owner = (
                str(row[owner_col]).strip()
                if owner_col and pd.notna(row[owner_col])
                else "Unknown"
            )
            val_driver = (
                str(row[driver_col]).strip()
                if driver_col and pd.notna(row[driver_col])
                else ""
            )

            val_office = (
                str(row[office_col]).strip()
                if office_col and pd.notna(row[office_col])
                else ""
            )
            val_make = (
                str(row[make_col]).strip()
                if make_col and pd.notna(row[make_col])
                else ""
            )
            val_type = (
                str(row[type_col]).strip()
                if type_col and pd.notna(row[type_col])
                else ""
            )
            val_emission = (
                str(row[emission_col]).strip()
                if emission_col and pd.notna(row[emission_col])
                else ""
            )
            val_gsis = (
                str(row[gsis_col]).strip()
                if gsis_col and pd.notna(row[gsis_col])
                else ""
            )
            val_lto = (
                str(row[lto_col]).strip() if lto_col and pd.notna(row[lto_col]) else ""
            )
            val_last_reg = (
                str(row[last_reg_col]).strip()
                if last_reg_col and pd.notna(row[last_reg_col])
                else ""
            )
            val_insurance = (
                clean_currency(row[insurance_col])
                if insurance_col and pd.notna(row[insurance_col])
                else ""
            )
            val_cost = (
                clean_currency(row[cost_col])
                if cost_col and pd.notna(row[cost_col])
                else ""
            )

            acq_d = (
                row[acq_date_col]
                if acq_date_col and pd.notna(row[acq_date_col])
                else ""
            )
            val_acq_date = ""
            if acq_d != "":
                try:
                    if hasattr(acq_d, "strftime"):
                        val_acq_date = acq_d.strftime("%Y-%m-%d")
                    else:
                        val_acq_date = str(acq_d).split(" ")[0]
                except:
                    val_acq_date = str(acq_d)

            val_phys_status = (
                str(row[phys_status_col]).strip()
                if phys_status_col and pd.notna(row[phys_status_col])
                else ""
            )

            if (
                pd.isna(plate)
                or str(plate).strip() == ""
                or str(plate).upper() == "CRITERIA"
            ):
                # Avoid breaking fully if there is just an empty row, unless it explicitly says CRITERIA
                if str(plate).upper() == "CRITERIA":
                    break
                continue

            plate = str(plate).strip()
            exp_date = (
                row[exp_col]
                if exp_col in df_sheet.columns and pd.notna(row[exp_col])
                else None
            )

            status = None
            # NATIVE EXCEL ALERT READING
            if (
                alert_col
                and pd.notna(row[alert_col])
                and str(row[alert_col]).strip() != ""
            ):
                val = str(row[alert_col]).strip().upper()
                if "EXPIRED" in val or "LESS THAN" in val:
                    status = "LESS THAN 0 DAYS = EXPIRED"
                elif "EXPIRY" in val or "1 TO 14" in val or "1-14" in val:
                    status = "1 TO 14 DAYS = DAYS BEFORE EXPIRY"
                elif "NOTICE" in val or "15 TO 29" in val or "15-29" in val:
                    status = "15 TO 29 DAYS = DAYS BEFORE 2 WEEK NOTICE"
                elif "SUFFICIENT" in val or "MORE" in val or "30 DAYS" in val:
                    status = "30 DAYS AND MORE = SUFFICIENT TIME"
                elif "INPUT" in val:
                    status = "PLEASE INPUT LAST REG"
                elif "REGISTERED" in val or "YES" in val:
                    status = "REGISTERED"

            # Fallback if no alert mapped from Excel
            if not status:
                status_override = None
                if status_col and pd.notna(row[status_col]):
                    val = str(row[status_col]).strip().upper()
                    if val in ["YES", "REGISTERED"]:
                        status_override = "REGISTERED"
                status = get_expiration_status(exp_date, status_override)

            current_state[plate] = (
                status,
                exp_date,
                sheet_name,
                owner,
                val_office,
                val_make,
                val_type,
                val_emission,
                val_gsis,
                val_lto,
                val_last_reg,
                val_insurance,
                val_cost,
                val_acq_date,
                val_phys_status,
                val_driver,
            )

            if not first_run or manual_sheet_target is not None:
                old_state = previous_state.get(plate, None)
                if old_state is not None:
                    old_status = old_state[0]
                    old_exp = old_state[1] if len(old_state) > 1 else None
                    old_sheet = old_state[2] if len(old_state) > 2 else "Unknown"

                    if (
                        old_status != status
                        or old_exp != exp_date
                        or old_sheet != sheet_name
                    ):
                        changed_records.append(
                            {
                                "plate": plate,
                                "owner": owner,
                                "old_status": old_status,
                                "new_status": status,
                                "sheet": sheet_name,
                                "exp_date": exp_date,
                            }
                        )
                elif old_state is None and (
                    "EXPIRED" in status
                    or "BEFORE EXPIRY" in status
                    or "NOTICE" in status
                ):
                    changed_records.append(
                        {
                            "plate": plate,
                            "owner": owner,
                            "old_status": "NEW RECORD",
                            "new_status": status,
                            "sheet": sheet_name,
                            "exp_date": exp_date,
                        }
                    )

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
        print(
            f"{Fore.CYAN}--- Initial Scan Results ({len(dfs)} sheets checked) ---{Style.RESET_ALL}"
        )
        initial_alerts = {}
        for plate, state_tuple in combined_current_state.items():
            status, exp_date, sheet_name = (
                state_tuple[0],
                state_tuple[1],
                state_tuple[2],
            )
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
            driver = state_tuple[15] if len(state_tuple) > 15 else ""
            print_status(f"[{plate}] {status}", status)
            if status not in initial_alerts:
                initial_alerts[status] = []
            initial_alerts[status].append(
                format_plate_with_data(
                    plate,
                    exp_date,
                    sheet_name,
                    owner,
                    office,
                    make,
                    type_val,
                    emission,
                    gsis,
                    lto,
                    last_reg,
                    cost,
                    acq_date,
                    phys_status,
                    status,
                    insurance,
                    driver,
                )
            )

        print(f"{Fore.CYAN}--- End Initial Scan ---{Style.RESET_ALL}")

        if initial_alerts:
            send_notification(
                initial_alerts, title="⚠ Initial Scan Results", is_auto=True
            )
        else:
            send_notification(
                {"SUFFICIENT TIME": ["All Plates inside Excel File"]},
                title="⚠ Initial Scan Results",
                is_auto=True,
            )

    elif combined_changed_records or is_manual_scan:
        # User requested a specific sheet or requested "Scan All"
        if is_manual_scan:
            print(
                f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Manual Scan Triggered{Style.RESET_ALL}"
            )
            title_text = f"⚠ Manual Scan: {manual_sheet_target if manual_sheet_target else 'All Sheets'}"

            manual_alerts = {}
            # Just pull from the results of what we read!
            for plate, state_tuple in combined_current_state.items():
                status, exp_date, sheet_name = (
                    state_tuple[0],
                    state_tuple[1],
                    state_tuple[2],
                )
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
                driver = state_tuple[15] if len(state_tuple) > 15 else ""
                if status not in manual_alerts:
                    manual_alerts[status] = []
                manual_alerts[status].append(
                    format_plate_with_data(
                        plate,
                        exp_date,
                        sheet_name,
                        owner,
                        office,
                        make,
                        type_val,
                        emission,
                        gsis,
                        lto,
                        last_reg,
                        cost,
                        acq_date,
                        phys_status,
                        status,
                        insurance,
                        driver,
                    )
                )

            if manual_alerts:
                send_notification(manual_alerts, title=title_text, is_auto=False)
            else:
                send_notification(
                    {"SUFFICIENT TIME": [f"All vehicles checked are valid."]},
                    title=title_text,
                    is_auto=False,
                )
            return True

        if not is_manual_scan:
            print(
                f"\n{Fore.CYAN}[{datetime.now().strftime('%H:%M:%S')}] Background Change Detected!{Style.RESET_ALL}"
            )
            changed_sheets = list(set([r["sheet"] for r in combined_changed_records]))
            sheet_title_str = (
                ", ".join(changed_sheets)
                if len(changed_sheets) < 3
                else f"{len(changed_sheets)} Sheets"
            )

            for record in combined_changed_records:
                plate = record["plate"]
                owner = record.get("owner", "Unknown")
                old = record["old_status"]
                new = record["new_status"]
                sheet = record["sheet"]
                print_status(
                    f"Real-time Update ({sheet}): [{plate}] ({owner}) {old} -> {new}",
                    new,
                )

            # Send comprehensive updated state so UI refreshes real-time
            full_alerts = {}
            for plate, state_tuple in combined_current_state.items():
                status, exp_date, sheet_name = (
                    state_tuple[0],
                    state_tuple[1],
                    state_tuple[2],
                )
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
                driver = state_tuple[15] if len(state_tuple) > 15 else ""
                if status not in full_alerts:
                    full_alerts[status] = []
                full_alerts[status].append(
                    format_plate_with_data(
                        plate,
                        exp_date,
                        sheet_name,
                        owner,
                        office,
                        make,
                        type_val,
                        emission,
                        gsis,
                        lto,
                        last_reg,
                        cost,
                        acq_date,
                        phys_status,
                        status,
                        insurance,
                        driver,
                    )
                )

            if full_alerts:
                send_notification(
                    full_alerts,
                    title=f"⚠ Real-time File Update: {sheet_title_str}",
                    is_auto=True,
                )
            else:
                send_notification(
                    {"SUFFICIENT TIME": ["All Vehicles clear in latest update!"]},
                    title=f"⚠ Real-time File Update: {sheet_title_str}",
                    is_auto=True,
                )

    if manual_sheet_target is None:
        previous_state = combined_current_state
        first_run = False

    return True


tracked_mtimes = {}


def background_monitor():
    global monitor_active, tracked_mtimes
    last_checked_date = datetime.now().date()

    while monitor_active:
        try:
            current_date = datetime.now().date()
            if current_date != last_checked_date:
                tracked_mtimes = {}
                last_checked_date = current_date

            if os.path.exists(EXCEL_FILE):
                current_mtime = os.path.getmtime(EXCEL_FILE)
                if tracked_mtimes.get(EXCEL_FILE) != current_mtime:
                    time.sleep(2)
                    process_excel(EXCEL_FILE)
                    try:
                        tracked_mtimes[EXCEL_FILE] = os.path.getmtime(EXCEL_FILE)
                    except WindowsError:
                        pass
            time.sleep(CHECK_INTERVAL_SECONDS)
        except Exception as e:
            time.sleep(CHECK_INTERVAL_SECONDS)


# Manual Scan All via Tray (Sends entire overview)
def on_scan_all(icon, item):
    global current_viewed_sheet
    current_viewed_sheet = None
    print("Manually Scanning Excel...")
    threading.Thread(
        target=process_excel,
        args=(EXCEL_FILE,),
        kwargs={"is_manual_scan": True},
        daemon=True,
    ).start()


def make_scan_sheet_callback(sheet_name):
    def callback(icon, item):
        global current_viewed_sheet
        current_viewed_sheet = sheet_name
        print(f"Manually Scanning: {sheet_name}")
        threading.Thread(
            target=process_excel,
            args=(EXCEL_FILE,),
            kwargs={"manual_sheet_target": sheet_name, "is_manual_scan": True},
            daemon=True,
        ).start()

    return callback


def on_exit(icon, item):
    global monitor_active
    monitor_active = False
    icon.stop()
    print("Exiting...")
    gui_queue.put({"type": "exit"})


def pystray_runner():
    global tray_icon
    image = create_image()
    tray_icon = pystray.Icon("VehicleMonitor", image, "Vehicle Alert System")

    def setup_menu():
        items = [pystray.MenuItem("Scan Excel", on_scan_all), pystray.Menu.SEPARATOR]
        if current_sheets:
            sheet_menus = []
            for sheet in current_sheets:
                sheet_menus.append(
                    pystray.MenuItem(f"Scan {sheet}", make_scan_sheet_callback(sheet))
                )
            items.append(pystray.MenuItem("Scan Month...", pystray.Menu(*sheet_menus)))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Exit", on_exit))
        return items

    tray_icon.menu = pystray.Menu(setup_menu)
    tray_icon.run()


def main():
    # Force working directory to the script's directory for safety
    if hasattr(sys, 'frozen'):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(app_dir)

    # Smart Detect Excel Name
    if not os.path.exists(EXCEL_FILE):
        xlsx_files = [f for f in os.listdir(".") if f.lower().endswith(".xlsx") and not f.startswith("~")]
        if len(xlsx_files) == 1:
            try:
                os.rename(xlsx_files[0], EXCEL_FILE)
            except Exception as e:
                pass

    # --- Single Instance Lock ---
    global lock_socket
    lock_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        lock_socket.bind(("127.0.0.1", 47123))

        def listen_for_triggers():
            while monitor_active:
                try:
                    lock_socket.settimeout(1.0)
                    data, _ = lock_socket.recvfrom(1024)
                    if data == b"trigger":
                        with open("startup_log.txt", "a") as f:
                            f.write("Received trigger at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                        # Request UI to show window
                        gui_queue.put({"type": "show_request"})
                        threading.Thread(
                            target=process_excel,
                            args=(EXCEL_FILE,),
                            kwargs={"is_manual_scan": True},
                            daemon=True,
                        ).start()
                except socket.timeout:
                    continue
                except Exception as e:
                    break

        threading.Thread(target=listen_for_triggers, daemon=True).start()
    except socket.error:
        print("Vehicle Monitor is already running. Pinging the active instance...")
        try:
            client_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            client_sock.sendto(b"trigger", ("127.0.0.1", 47123))
        except:
            pass
        sys.exit(0)
    # --- End Single Instance Lock ---

    print(f"{Fore.GREEN}Starting Vehicle Monitor Dashboard...{Style.RESET_ALL}")

    tray_thread = threading.Thread(target=pystray_runner, daemon=True)
    tray_thread.start()

    # PyQt6 Main Window must be in main thread
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    window = AlertWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
