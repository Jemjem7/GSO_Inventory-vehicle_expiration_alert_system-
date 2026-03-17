import os
import re

def execute_debug_patch():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # The AI extraction patch might have a bug inside the Treeview instantiation because `columns_list` might contain weird characters, or `tree.heading()` is failing on an empty value. Let's add prints.
    
    old_method_head = "def build_ui(self, detailed_alerts, columns_list, window_title):"
    new_method_head = """def build_ui(self, detailed_alerts, columns_list, window_title):
        print(f"DEBUG: Entering build_ui. Alerts keys: {list(detailed_alerts.keys())}")
        print(f"DEBUG: Columns list: {columns_list}")"""

    code = code.replace(old_method_head, new_method_head)

    # Put a debug BEFORE column insertion
    old_tree = "tree = ttk.Treeview(summary_frame, columns=columns, show=\"headings\", style=\"Custom.Treeview\", height=15)"
    new_tree = """print(f"DEBUG: Initializing Treeview with columns: {columns}")
        try:
            tree = ttk.Treeview(summary_frame, columns=columns, show="headings", style="Custom.Treeview", height=15)
        except Exception as e:
            print(f"DEBUG: CRASH ON TREEVIEW INIT: {e}")"""
            
    code = code.replace(old_tree, new_tree)
    
    old_col_loop = """        for col in columns:
            if col == "_sheet": display_text = "MONTH / SHEET"
            elif col == "_status": display_text = "ALERT STATUS"
            else: display_text = str(col).upper()
            
            tree.heading(col, text=display_text, anchor=tk.W)
            if col in ['_sheet', '_status']: w = 150
            else: w = col_widths.get(col, 100)
            tree.column(col, width=w, minwidth=60, stretch=tk.YES if w > 120 else tk.NO)"""
            
    new_col_loop = """        print("DEBUG: Inserting Columns Headers...")
        for col in columns:
            try:
                if col == "_sheet": display_text = "MONTH / SHEET"
                elif col == "_status": display_text = "ALERT STATUS"
                else: display_text = str(col).upper()
                
                tree.heading(col, text=display_text, anchor=tk.W)
                if col in ['_sheet', '_status']: w = 150
                else: w = col_widths.get(col, 100)
                tree.column(col, width=w, minwidth=60, stretch=tk.YES if w > 120 else tk.NO)
            except Exception as e:
                print(f"DEBUG: CRASH INJECTING COLUMN {col}: {e}")"""
                
    code = code.replace(old_col_loop, new_col_loop)
    
    # Put a debug AFTER row insertions
    old_row_loop = "if has_alerts:"
    new_row_loop = "print(f'DEBUG: Reached end of row insertions. has_alerts={has_alerts}')\n        if has_alerts:"
    code = code.replace(old_row_loop, new_row_loop)

    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    execute_debug_patch()
