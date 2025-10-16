from py_stealth import *

import sys
import os
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import StringVar
import threading
import re
import subprocess
import platform

from modules.connection import get_char_name
import modules.connection as connection
import modules.common_utils as utils
from bs_config import veritePackList, agapitePackList, goldPackList
from xlsGenerator import export_to_excel

AUTOMATION_RUNNING = False
SELECTED_PACK_NAME = None
EXPORT_PATH = None
STATUS_TEXT = None

char_name = connection.get_char_name()

bodPackOptions = {
    "Verite": veritePackList,
    "Agapite": agapitePackList,
    "Gold": goldPackList
}

def PropertyTxt(objid=None, name=None):
    tooltip = GetTooltip(objid)   
    properties = tooltip.upper().split("|") 
    for i, property in enumerate(properties):
        if name.upper() in property:
            if ': ' in property:
                prop = property.split(': ')
                propname = prop[0]
                propvalue = str(prop[1]) 
                return propvalue
            else:
                utils.debug('[Property] Error: Invalid', property)

def sanitize_sheet_name(name):
    # Remove caracteres inválidos e limita a 31 caracteres
    return re.sub(r'[:\\/?*\[\]]', '', name)[:31]

def Property(objid=None, name=None):
    tooltip = GetTooltip(objid)   
    properties = tooltip.upper().split("|") 
    for i, property in enumerate(properties):
        if name.upper() in property:
            if ': ' in property:
                prop = property.split(': ')
                propname = prop[0]
                propvalue = int(prop[1]) 
                return propvalue
            else:
                print('[Property] Error: Invalid', property)

def countBods(packList, collected_data):
    SetFindVertical(70)
    SetFindDistance(50)
    if FindTypeEx(0xe75, 0xffff, Ground(), False):
        bagsDone = GetFindedList()
        for bagDone in bagsDone:
            for k, v in packList.items():
                # print(f"Item Bag Done {hex(bagDone)}")
                # print(f"Item Packet {k}")
                # print(f"Item packet {hex(v)}")
                if hex(bagDone) == hex(v):
                    bag_id = v
                    name = k
                    bagBodStone = PropertyTxt(bag_id, 'Contents:')
                    count = int(bagBodStone.split("/")[0])
                    if count >= 0:
                        if "verite" in name:
                            ore_type = "verite"
                        elif "valorite" in name:
                            ore_type = "valorite"
                        elif "agapite" in name:
                            ore_type = "agapite"
                        elif "gold" in name:
                            ore_type = "gold"
                        elif "bronze" in name:
                            ore_type = "bronze"
                        elif "cooper" in name:
                            ore_type = "cooper"
                        else:
                            ore_type = "outro"

                        collected_data.append({
                            "Bod": name,
                            "Type": ore_type,
                            "Amount": count
                        })

def update_status(message):
    global STATUS_TEXT
    if STATUS_TEXT:
        STATUS_TEXT.set(message)
        root.update()

def open_export_folder():
    global EXPORT_PATH
    if EXPORT_PATH and os.path.exists(EXPORT_PATH):
        if platform.system() == "Windows":
            os.startfile(EXPORT_PATH)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", EXPORT_PATH])
        else:  # Linux
            subprocess.run(["xdg-open", EXPORT_PATH])
    else:
        messagebox.showwarning("Warning", "Export folder not found or not configured!")

def browse_folder():
    global EXPORT_PATH
    folder = filedialog.askdirectory(title="Select Export Folder")
    if folder:
        EXPORT_PATH = folder
        export_path_var.set(folder)
        # Update bs_config
        import bs_config
        bs_config.export_path = folder

def automation_loop():
    global AUTOMATION_RUNNING, SELECTED_PACK_NAME, EXPORT_PATH
    SHOULD_RESTART = False
    
    while AUTOMATION_RUNNING:
        if SHOULD_RESTART:
            SHOULD_RESTART = False
            continue
       
        if not EXPORT_PATH:
            update_status("❌ Please select an export folder first!")
            messagebox.showerror("Error", "Please select an export folder before starting!")
            break
            
        selected_list = bodPackOptions.get(SELECTED_PACK_NAME.get(), None)
        if selected_list:
            update_status("🔍 Searching for BOD bags...")
            bod_data = []
            countBods(selected_list, bod_data)
            
            if len(bod_data) > 0:
                update_status("📊 Generating Excel report...")
                sheet_name = sanitize_sheet_name(SELECTED_PACK_NAME.get())
                success = export_to_excel(bod_data, sheet_name=sheet_name)
                if success:
                    update_status("✅ Report generated successfully!")
                    messagebox.showinfo("Success", f"BOD report generated successfully!\n\nFound {len(bod_data)} BOD entries.\n\nFile saved to:\n{EXPORT_PATH}")
                else:
                    update_status("❌ Failed to generate report!")
            else:
                update_status("⚠️ No BOD data found!")
                messagebox.showwarning("Warning", "No BOD data was collected!\n\nMake sure:\n• Bags are on the ground near your character\n• Bag IDs in bs_config.py are correct\n• Your character is close enough to the bags")
            break
        else:
            update_status("❌ BOD list not found!")
            utils.debug(f"[automation_loop] Lista '{SELECTED_PACK_NAME.get()}' não encontrada.")

def update_player(name):
    global root
    root.current_playerName.set(f"Name : {name}")
                
# Thread-safe start
def start_automation():
    global AUTOMATION_RUNNING
    if not AUTOMATION_RUNNING:
        AUTOMATION_RUNNING = True
        threading.Thread(target=automation_loop, daemon=True).start()

def pause_automation():
    global AUTOMATION_RUNNING
    AUTOMATION_RUNNING = False

# Interface
def build_ui():
    global root, SELECTED_PACK_NAME, EXPORT_PATH, STATUS_TEXT, export_path_var

    root = tk.Tk()
    root.title("UO BOD Automator XLS Counter Report")
    root.geometry("500x400")
    root.configure(bg="#1e1e1e")
    root.resizable(False, False)

    # Initialize variables
    SELECTED_PACK_NAME = tk.StringVar(value="Verite")
    STATUS_TEXT = tk.StringVar(value="Ready to start...")
    export_path_var = tk.StringVar()
    
    # Set default export path
    import bs_config
    EXPORT_PATH = bs_config.export_path
    export_path_var.set(EXPORT_PATH)

    # Dark theme styling
    style = ttk.Style(root)
    style.theme_use("clam")
    
    # Configure dark theme colors
    style.configure("TLabel", background="#1e1e1e", foreground="#ffffff")
    style.configure("TFrame", background="#1e1e1e")
    style.configure("TButton", 
                    background="#2d2d2d", 
                    foreground="#ffffff",
                    font=("Segoe UI", 10, "bold"),
                    padding=8,
                    borderwidth=0)
    style.map("TButton",
              background=[("active", "#404040"), ("pressed", "#1a1a1a")],
              foreground=[("active", "#ffffff")])
    
    style.configure("TCombobox", 
                    background="#2d2d2d", 
                    foreground="#ffffff",
                    fieldbackground="#2d2d2d",
                    borderwidth=1)
    
    style.configure("TEntry",
                    background="#2d2d2d",
                    foreground="#ffffff",
                    fieldbackground="#2d2d2d",
                    borderwidth=1)

    # Header frame with logo area
    header_frame = ttk.Frame(root)
    header_frame.pack(pady=10, fill='x')
    
    # Logo placeholder (you can replace with actual image)
    logo_label = tk.Label(header_frame, 
                         text="📊 BOD COUNTER", 
                         font=("Segoe UI", 18, "bold"),
                         bg="#1e1e1e", 
                         fg="#4CAF50")
    logo_label.pack()
    
    subtitle_label = tk.Label(header_frame,
                             text="Ultima Online BOD Report Generator",
                             font=("Segoe UI", 10),
                             bg="#1e1e1e",
                             fg="#888888")
    subtitle_label.pack(pady=(2, 0))

    # Player info frame
    player_frame = ttk.Frame(root)
    player_frame.pack(pady=5, fill='x')
    
    current_playerName = tk.StringVar(value="Collector Name: -")
    player_label = ttk.Label(player_frame, textvariable=current_playerName, font=("Segoe UI", 12, "bold"))
    player_label.pack()

    # Export path frame
    export_frame = ttk.Frame(root)
    export_frame.pack(pady=8, padx=15, fill='x')
    
    ttk.Label(export_frame, text="Export Directory:", font=("Segoe UI", 10, "bold")).pack(anchor='w')
    
    path_frame = ttk.Frame(export_frame)
    path_frame.pack(fill='x', pady=(3, 0))
    
    path_entry = ttk.Entry(path_frame, textvariable=export_path_var, width=40, state="readonly")
    path_entry.pack(side='left', fill='x', expand=True, padx=(0, 8))
    
    browse_btn = ttk.Button(path_frame, text="Browse", command=browse_folder, width=8)
    browse_btn.pack(side='right')

    # BOD selection frame
    bod_frame = ttk.Frame(root)
    bod_frame.pack(pady=8, padx=15, fill='x')
    
    ttk.Label(bod_frame, text="BOD Collection:", font=("Segoe UI", 10, "bold")).pack(anchor='w')
    
    list_selector = ttk.Combobox(bod_frame, 
                                textvariable=SELECTED_PACK_NAME,
                                values=list(bodPackOptions.keys()), 
                                state="readonly",
                                width=18)
    list_selector.pack(anchor='w', pady=(3, 0))

    # Status frame
    status_frame = ttk.Frame(root)
    status_frame.pack(pady=8, padx=15, fill='x')
    
    ttk.Label(status_frame, text="Status:", font=("Segoe UI", 10, "bold")).pack(anchor='w')
    
    status_label = ttk.Label(status_frame, 
                            textvariable=STATUS_TEXT,
                            font=("Segoe UI", 9),
                            foreground="#4CAF50")
    status_label.pack(anchor='w', pady=(3, 0))

    # Control buttons frame
    control_frame = ttk.Frame(root)
    control_frame.pack(pady=15, fill='x')
    
    # Start button
    start_btn = ttk.Button(control_frame, 
                          text="🚀 Start", 
                          command=start_automation,
                          width=12)
    start_btn.pack(side='left', padx=(15, 8))
    
    # Pause button
    pause_btn = ttk.Button(control_frame, 
                          text="⏸️ Pause", 
                          command=pause_automation,
                          width=10)
    pause_btn.pack(side='left', padx=8)
    
    # Open folder button
    open_btn = ttk.Button(control_frame, 
                         text="📁 Open", 
                         command=open_export_folder,
                         width=10)
    open_btn.pack(side='right', padx=(8, 15))

    # Footer
    footer_frame = ttk.Frame(root)
    footer_frame.pack(side='bottom', fill='x', pady=5)
    
    footer_label = tk.Label(footer_frame,
                           text="Made for Astraroth",
                           font=("Segoe UI", 8),
                           bg="#1e1e1e",
                           fg="#666666")
    footer_label.pack()

    # Store references
    root.current_playerName = current_playerName
    update_player(char_name)

    root.mainloop()
    return root

if __name__ == "__main__":
    root = build_ui()
