import os

def fix_syntax():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()
        
    # Fix the global EXCEL_FILE string that was inadvertently butchered during the naive text replacement:
    old_broken_line = "list([f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')])[0] = 'VehicleMonitoring.xlsx'"
    new_fixed_line = "EXCEL_FILE = 'VehicleMonitoring.xlsx' # Default fallback if no dynamic files found"
    
    code = code.replace(old_broken_line, new_fixed_line)
    
    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    fix_syntax()
