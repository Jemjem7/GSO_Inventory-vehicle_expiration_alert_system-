import re

with open("vehicle_monitor.py", "r", encoding="utf-8") as f:
    original = f.read()

with open("gui_pyqt.py", "r", encoding="utf-8") as f:
    gui_code = f.read()

# Extract PyQt imports and the class
pyqt_imports = []
class_def_lines = []
in_class = False

for line in gui_code.split("\n"):
    if line.startswith("from PyQt6"):
        pyqt_imports.append(line)
    elif line.startswith("class AlertWindow"):
        in_class = True
        class_def_lines.append(line)
    elif in_class:
        class_def_lines.append(line)

new_class_str = "\n".join(class_def_lines)

# Find where to replace in original
# Remove tkinter imports
original = original.replace("import tkinter as tk\n", "")
original = original.replace("from tkinter import ttk\n", "\n" + "\n".join(pyqt_imports) + "\n")

# Replace class AlertWindow(tk.Tk): ... till end of do_scan_month
start_idx = original.find("class AlertWindow(tk.Tk):")

# Find the def do_scan_month
end_idx_search = original.find("def do_scan_month(self, selection):", start_idx)

# Find the end of do_scan_month function body
# It ends with threading.Thread(...)
thread_start_idx = original.find("threading.Thread(target=process_excel", end_idx_search)

# Find the first newline after thread_start_idx
end_idx = original.find("\\n", thread_start_idx)
if end_idx == -1:
    end_idx = original.find("\n", thread_start_idx) + 1

# Replace!
first_part = original[:start_idx]
second_part = original[end_idx:]

# The end_idx might leave some lines, so let's be careful and use regex or just standard find
# Actually, the string after thread_start_idx is:
#  args=(EXCEL_FILE, selection, True), daemon=True).start()\n\n
# Let's find def send_notification which comes right after.
send_notif_idx = original.find("def send_notification", start_idx)
new_original = original[:start_idx] + new_class_str + "\n\n" + original[send_notif_idx:]

# Update the main() block
main_old = '''    # TKinter Main Window must be in main thread
    window = AlertWindow()
    window.mainloop()'''

main_new = '''    # PyQt6 Main Window must be in main thread
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    window = AlertWindow()
    sys.exit(app.exec())'''

new_original = new_original.replace(main_old, main_new)

with open("vehicle_monitor.py", "w", encoding="utf-8") as f:
    f.write(new_original)

print("Patching complete!")
