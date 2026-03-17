import os
import re

def execute_syntax_repair():
    with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # The previous script did a naive `code.replace('try:', '')`, which deleted it globally. 
    # Let's manually restore the 3 crucial try blocks at the top of the script.
    
    old_theme = """def get_system_theme():
    
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)"""
        
    new_theme = """def get_system_theme():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)"""
        
    code = code.replace(old_theme, new_theme)
    
    old_load = """def load_settings():
    
        if os.path.exists("settings.json"):"""
        
    new_load = """def load_settings():
    try:
        if os.path.exists("settings.json"):"""
        
    code = code.replace(old_load, new_load)
    
    old_save = """def save_settings(settings):
    
        with open("settings.json", "w") as f:"""
        
    new_save = """def save_settings(settings):
    try:
        with open("settings.json", "w") as f:"""
        
    code = code.replace(old_save, new_save)

    # Let's ensure the socket block was also not damaged
    old_sock = """        lock_socket.bind(('127.0.0.1', 47123))"""
    if "try:\n        lock_socket.bind" not in code:
        code = code.replace(old_sock, "try:\n        lock_socket.bind(('127.0.0.1', 47123))")
        
    with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
        f.write(code)

if __name__ == '__main__':
    execute_syntax_repair()
