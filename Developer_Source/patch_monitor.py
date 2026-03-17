import os
import re

def patch_file():
    with open('vehicle_monitor_v1.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update EXCEL_FILE to active_file globally and background monitor
    # Wait, instead of global EXCEL_FILE, the script looks for all xlsx files.
    # Let's replace background_monitor block
    bg_monitor_old = """def background_monitor():
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
            time.sleep(CHECK_INTERVAL_SECONDS)"""

    bg_monitor_new = """tracked_mtimes = {}
def background_monitor():
    global monitor_active, tracked_mtimes
    last_checked_date = datetime.now().date()
    
    while monitor_active:
        try:
            current_date = datetime.now().date()
            if current_date != last_checked_date:
                tracked_mtimes = {}
                last_checked_date = current_date
                
            xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
            for f in xlsx_files:
                current_mtime = os.path.getmtime(f)
                if tracked_mtimes.get(f) != current_mtime:
                    time.sleep(2)
                    process_excel(f)
                    try:
                        tracked_mtimes[f] = os.path.getmtime(f)
                    except WindowsError:
                        pass
            time.sleep(CHECK_INTERVAL_SECONDS)
        except Exception as e:
            time.sleep(CHECK_INTERVAL_SECONDS)"""
    
    content = content.replace(bg_monitor_old, bg_monitor_new)

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
if __name__ == '__main__':
    patch_file()
