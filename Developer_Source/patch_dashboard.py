import re

with open("vehicle_monitor.py", "r", encoding="utf-8") as f:
    original = f.read()

with open("gui_dashboard.py", "r", encoding="utf-8") as f:
    gui_code = f.read()

# Extract PyQt imports and the classes
pyqt_imports = []
class_def_lines = []
in_class = False

for line in gui_code.split("\n"):
    if line.startswith("from PyQt6"):
        pyqt_imports.append(line)
    elif line.startswith("class ClickableLabel"):
        in_class = True
        class_def_lines.append(line)
    elif in_class:
        class_def_lines.append(line)

new_class_str = "\n".join(class_def_lines)

start_idx = original.find("class AlertWindow(QMainWindow):")

# Find the def do_scan_month to skip over the old class body
end_idx_search = original.find("def do_scan_month(self, selection):", start_idx)

# Find the end of on_row_click function body which is the end of the class
thread_start_idx = original.find("threading.Thread(target=open_excel_threaded", end_idx_search)
end_idx = original.find("\n", thread_start_idx) + 1 # Add 1 to actually include the newline

send_notif_idx = original.find("def send_notification", start_idx)

new_original = original[:start_idx] + new_class_str + "\n\n" + original[send_notif_idx:]

with open("vehicle_monitor.py", "w", encoding="utf-8") as f:
    f.write(new_original)

print("Dashboard Patching complete!")
