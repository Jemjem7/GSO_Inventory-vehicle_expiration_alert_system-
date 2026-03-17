import os

with open('vehicle_monitor.py', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Update theme block
old_theme_block = """        if actual_theme == "Dark":
            bg_color = '#202124'
            fg_color = '#E8EAED'
            panel_bg = '#2D2E31'
            text_fg = '#E8EAED'
            sub_fg = '#9AA0A6'
            stripe_1 = '#2D2E31'
            stripe_2 = '#35363A'
            importance_order = ["""

new_theme_block = """        top_bar_bg = "black"
        header_fg = "white"
        tree_head_bg = ""
        tree_head_fg = ""
        tree_head_active_bg = ""
        tree_sel = ""
        btn_bg = ""
        btn_fg = ""
        btn_border = ""
        btn_active = ""
        menu_bg = ""
        menu_fg = ""
        menu_sel = ""
        hover_color = ""
        success_fg = ""
        
        if actual_theme == "Dark":
            bg_color = '#202124'
            fg_color = '#E8EAED'
            panel_bg = '#2D2E31'
            text_fg = '#E8EAED'
            sub_fg = '#9AA0A6'
            stripe_1 = '#2D2E31'
            stripe_2 = '#35363A'
            
            top_bar_bg = '#000000'
            header_fg = '#FFFFFF'
            tree_head_bg = '#202124'
            tree_head_fg = '#E8EAED'
            tree_head_active_bg = '#3C4043'
            tree_sel = '#5F6368'
            btn_bg = '#3C4043'
            btn_fg = '#E8EAED'
            btn_border = '#5F6368'
            btn_active = '#5F6368'
            menu_bg = '#2D2E31'
            menu_fg = '#E8EAED'
            menu_sel = '#5F6368'
            hover_color = '#35363A'
            success_fg = '#66cc66'
            importance_order = ["""

if old_theme_block in text:
    text = text.replace(old_theme_block, new_theme_block)
else:
    print("Failed to find old_theme_block")


old_light_block = """        else:
            bg_color = '#F1F3F4'
            fg_color = '#202124'
            panel_bg = '#FFFFFF'
            text_fg = '#202124'
            sub_fg = '#5F6368'
            stripe_1 = '#FFFFFF'
            stripe_2 = '#F8F9FA'
            importance_order = ["""

new_light_block = """        elif actual_theme == "Nature":
            bg_color = '#213722' # Dark Green
            fg_color = '#F8C662' # Saffron
            panel_bg = '#41644A' # Hunter Green
            text_fg = '#FFFFFF'
            sub_fg = '#F8C662'
            stripe_1 = '#41644A'
            stripe_2 = '#2C263F'
            
            top_bar_bg = '#2C263F' # Dark Purple
            header_fg = '#F8C662' # Saffron
            tree_head_bg = '#595082' # Ultra Violet
            tree_head_fg = '#FFFFFF'
            tree_head_active_bg = '#2C263F'
            tree_sel = '#2C263F'
            btn_bg = '#595082'
            btn_fg = '#FFFFFF'
            btn_border = '#2C263F'
            btn_active = '#2C263F'
            menu_bg = '#41644A'
            menu_fg = '#FFFFFF'
            menu_sel = '#213722'
            hover_color = '#595082'
            success_fg = '#F8C662'
            
            importance_order = [
                ('1 WEEK BEFORE EXPIRY', '#F28B82'),
                ('1 MONTH BEFORE EXPIRY', '#FDC69C'),
                ('2 MONTHS BEFORE EXPIRY', '#FDE293'),
                ('EXPIRED', '#F28B82'),
                ('DAYS BEFORE EXPIRY', '#FDC69C'),
                ('DAYS BEFORE 2 WEEK NOTICE', '#FDE293'),
                ('SUFFICIENT TIME', '#81C995'),
                ('PLEASE INPUT LAST REG', '#9AA0A6'),
                ('REGISTERED', '#8AB4F8')
            ]
        else:
            bg_color = '#F1F3F4'
            fg_color = '#202124'
            panel_bg = '#FFFFFF'
            text_fg = '#202124'
            sub_fg = '#5F6368'
            stripe_1 = '#FFFFFF'
            stripe_2 = '#F8F9FA'
            
            top_bar_bg = '#000000'
            header_fg = '#FFFFFF'
            tree_head_bg = '#F1F3F4'
            tree_head_fg = '#202124'
            tree_head_active_bg = '#E8EAED'
            tree_sel = '#DADCE0'
            btn_bg = '#E8EAED'
            btn_fg = '#202124'
            btn_border = '#DADCE0'
            btn_active = '#DADCE0'
            menu_bg = '#FFFFFF'
            menu_fg = '#202124'
            menu_sel = '#E8EAED'
            hover_color = '#E8EAED'
            success_fg = '#2e7d32'
            importance_order = ["""

if old_light_block in text:
    text = text.replace(old_light_block, new_light_block)
else:
    print("Failed to find old_light_block")

# Replace top_bar bg="black" to bg=top_bar_bg
text = text.replace('top_bar = tk.Frame(self.main_container, bg="black")', 'top_bar = tk.Frame(self.main_container, bg=top_bar_bg)')
text = text.replace('ll_label = tk.Label(top_bar, image=self.logo_l_photo, bg="black")', 'll_label = tk.Label(top_bar, image=self.logo_l_photo, bg=top_bar_bg)')
text = text.replace('rr_label = tk.Label(top_bar, image=self.logo_r_photo, bg="black")', 'rr_label = tk.Label(top_bar, image=self.logo_r_photo, bg=top_bar_bg)')
full_header_old = 'header = tk.Label(top_bar, text=header_text, font=("Segoe UI", 16, "bold"), bg="black", fg="white", justify="center")'
full_header_new = 'header = tk.Label(top_bar, text=header_text, font=("Segoe UI", 16, "bold"), bg=top_bar_bg, fg=header_fg, justify="center")'
if full_header_old in text:
    text = text.replace(full_header_old, full_header_new)

# Replace hover color logic
hover_old = "hover_color = '#35363A' if actual_theme == 'Dark' else '#E8EAED'\n        tree.tag_configure('hover', background=hover_color)"
hover_new = "tree.tag_configure('hover', background=hover_color)"
text = text.replace(hover_old, hover_new)

# Replace success status foreground
succ_old = "fg='#66cc66' if actual_theme == 'Dark' else '#2e7d32')"
succ_new = "fg=success_fg)"
text = text.replace(succ_old, succ_new)

# Fix ttk buttons setup 
old_style_block = """        if actual_theme == 'Dark':
            style.configure('TButton', background='#3C4043', foreground='#E8EAED', bordercolor='#5F6368', font=('Segoe UI', 9))
            style.map('TButton', background=[('active', '#5F6368')])
            style.configure('TMenubutton', background='#3C4043', foreground='#E8EAED', bordercolor='#5F6368', font=('Segoe UI', 9))
            style.map('TMenubutton', background=[('active', '#5F6368')])
            style.configure("Custom.Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text_fg, borderwidth=0, font=("Segoe UI", 10), rowheight=26)
            style.configure("Custom.Treeview.Heading", background='#202124', foreground='#E8EAED', font=("Segoe UI", 10, "bold"), borderwidth=0, padding=4)
            style.map("Custom.Treeview.Heading", background=[('active', '#3C4043')])
            style.map("Custom.Treeview", background=[('selected', '#5F6368')])
            self.option_add("*Menu.background", "#2D2E31")
            self.option_add("*Menu.foreground", "#E8EAED")
            self.option_add("*Menu.selectColor", "#5F6368")
        else:
            style.configure('TButton', background='#E8EAED', foreground='#202124', bordercolor='#DADCE0', font=('Segoe UI', 9))
            style.map('TButton', background=[('active', '#DADCE0')])
            style.configure('TMenubutton', background='#E8EAED', foreground='#202124', bordercolor='#DADCE0', font=('Segoe UI', 9))
            style.map('TMenubutton', background=[('active', '#DADCE0')])
            
            style.configure("Custom.Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text_fg, borderwidth=0, font=("Segoe UI", 10), rowheight=26)
            style.configure("Custom.Treeview.Heading", background='#F1F3F4', foreground='#202124', font=("Segoe UI", 10, "bold"), borderwidth=0, padding=4)
            style.map("Custom.Treeview.Heading", background=[('active', '#E8EAED')])
            style.map("Custom.Treeview", background=[('selected', '#DADCE0')])
            self.option_add("*Menu.background", "#FFFFFF")
            self.option_add("*Menu.foreground", "#202124")
            self.option_add("*Menu.selectColor", "#E8EAED")"""

new_style_block = """        style.configure('TButton', background=btn_bg, foreground=btn_fg, bordercolor=btn_border, font=('Segoe UI', 9))
        style.map('TButton', background=[('active', btn_active)])
        style.configure('TMenubutton', background=btn_bg, foreground=btn_fg, bordercolor=btn_border, font=('Segoe UI', 9))
        style.map('TMenubutton', background=[('active', btn_active)])
        
        style.configure("Custom.Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text_fg, borderwidth=0, font=("Segoe UI", 10), rowheight=26)
        style.configure("Custom.Treeview.Heading", background=tree_head_bg, foreground=tree_head_fg, font=("Segoe UI", 10, "bold"), borderwidth=0, padding=4)
        style.map("Custom.Treeview.Heading", background=[('active', tree_head_active_bg)])
        style.map("Custom.Treeview", background=[('selected', tree_sel)])
        
        self.option_add("*Menu.background", menu_bg)
        self.option_add("*Menu.foreground", menu_fg)
        self.option_add("*Menu.selectColor", menu_sel)"""

if old_style_block in text:
    text = text.replace(old_style_block, new_style_block)
else:
    print("Failed to replace old_style_block")

# Add Nature to theme dropdown array
drp_old = '"Light", "Dark", "System", command=self.change_theme'
drp_new = '"Nature", "Light", "Dark", "System", command=self.change_theme'
text = text.replace(drp_old, drp_new)

t_drp_old = "theme_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')"
t_drp_new = "theme_dropdown['menu'].configure(bg=menu_bg, fg=menu_fg)"
text = text.replace(t_drp_old, t_drp_new)

s_drp_old = "sheet_dropdown['menu'].configure(bg='#2d2d2d' if actual_theme == 'Dark' else '#f0f0f0', fg='#ffffff' if actual_theme == 'Dark' else '#000000')"
s_drp_new = "sheet_dropdown['menu'].configure(bg=menu_bg, fg=menu_fg)"
text = text.replace(s_drp_old, s_drp_new)

# Force default theme to "Nature" if it was "System" so they see it immediately
def_old = 'app_settings = load_settings()'
def_new = 'app_settings = load_settings()\napp_settings["theme"] = app_settings.get("theme", "Nature")'
text = text.replace(def_old, def_new)

# One more place for default
init_th_old = 'self.current_theme = app_settings.get("theme", "System")'
init_th_new = 'self.current_theme = app_settings.get("theme", "Nature")'
text = text.replace(init_th_old, init_th_new)

with open('vehicle_monitor.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Patch complete.")
