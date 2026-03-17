import os
import shutil

def cleanup_and_enhance():
    # Priority 1: Move extra Excel files to a backup directory so they don't hijack the application
    backup_dir = "backup_excel_files"
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
        
    allowed_file = "VehicleMonitoring.xlsx"
    for f in os.listdir('.'):
        if f.endswith('.xlsx') and not f.startswith('~') and f.lower() != allowed_file.lower():
            try:
                shutil.move(f, os.path.join(backup_dir, f))
                print(f"Moved {f} to backup.")
            except Exception as e:
                print(f"Failed to move {f}: {e}")

    # Priority 2: Add the selected month title above the Treeview in vehicle_monitor.py
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()
        
    old_clock = """        self.clock_label.pack()
        self.update_clock()"""
        
    new_clock = """        self.clock_label.pack()
        self.update_clock()
        
        # Inject the selected Month / View title right above the Treeview
        display_title = window_title.replace("⚠ ", "").replace("??? ", "").upper()
        title_label = tk.Label(summary_frame, text=display_title, font=("Segoe UI", 14, "bold"), bg=panel_bg, fg=fg_color)
        title_label.pack(fill=tk.X, pady=(0, 5))"""
        
    code = code.replace(old_clock, new_clock)
    
    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    cleanup_and_enhance()
