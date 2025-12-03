import os
import json
import pythoncom
import subprocess
import threading
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageDraw
import pystray
import win32com.client
import win32api, win32ui
from pathlib import Path

CONFIG_FILE = "config.json"
KEY_LABELS = [f"F{i}" for i in range(13, 25)]

# =============================================
# UTILIDADES
# =============================================
def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {
            "assignments": {},
            "appearance_mode": "dark",
            "start_with_windows": False,
            "start_minimized": False
        }
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def save_config(data):
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)

def get_programs():
    paths = [
        os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs"),
        os.path.join(os.environ["PROGRAMDATA"], r"Microsoft\Windows\Start Menu\Programs"),
        "C:\\Program Files",
        "C:\\Program Files (x86)"
    ]
    programs = []
    for base in paths:
        if not os.path.exists(base):
            continue
        for root, _, files in os.walk(base):
            for file in files:
                if file.lower().endswith((".exe", ".lnk")):
                    programs.append(os.path.join(root, file))
    return sorted(programs)

def extract_icon(path):
    try:
        if path.lower().endswith(".lnk"):
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)
            path = shortcut.Targetpath
        if not os.path.exists(path):
            return None
        large, small = win32api.ExtractIconEx(path, 0)
        hicon = small[0] if small else (large[0] if large else None)
        if not hicon:
            return None
        dc = win32ui.CreateDCFromHandle(win32api.GetDC(0))
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(dc, 64, 64)
        memdc = dc.CreateCompatibleDC()
        memdc.SelectObject(bmp)
        memdc.DrawIcon((0, 0), hicon)
        bmp.SaveBitmapFile(memdc, "tmp_icon.bmp")
        img = Image.open("tmp_icon.bmp").resize((48, 48))
        os.remove("tmp_icon.bmp")
        return ctk.CTkImage(img, img)
    except:
        return None

# =============================================
# FUNCIONALIDAD INICIO AUTOMÁTICO
# =============================================
def set_startup(enable=True):
    try:
        pythoncom.CoInitialize()
        startup_dir = Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs\Startup"
        exe_path = Path(os.path.abspath("APPVIAL.exe"))  # Ruta del exe final
        shortcut_path = startup_dir / "APPVIAL.lnk"

        shell = win32com.client.Dispatch("WScript.Shell")
        if enable:
            shortcut = shell.CreateShortCut(str(shortcut_path))
            shortcut.Targetpath = str(exe_path)
            shortcut.WorkingDirectory = str(exe_path.parent)
            shortcut.save()
        else:
            if shortcut_path.exists():
                shortcut_path.unlink()
    except Exception as e:
        print("Error set_startup:", e)

# =============================================
# LANZADOR CUADRÍCULA
# =============================================
class LaunchGridPage(ctk.CTkFrame):
    def __init__(self, master, program_assignments, open_program_window, unassign_callback):
        super().__init__(master)
        self.program_assignments = program_assignments
        self.open_program_window = open_program_window
        self.unassign_callback = unassign_callback
        self.icon_images = {}
        self.buttons = {}
        rows, cols = 2, 6
        for i in range(rows):
            self.grid_rowconfigure(i, weight=1)
        for j in range(cols):
            self.grid_columnconfigure(j, weight=1)
        index = 0
        for r in range(rows):
            for c in range(cols):
                key = KEY_LABELS[index]
                self.create_button(r, c, key)
                index += 1

    def create_button(self, row, col, key):
        frame = ctk.CTkFrame(self)
        frame.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
        program_path = self.program_assignments.get(key)
        icon_img = None
        if program_path:
            icon_img = self.icon_images.get(program_path)
            if not icon_img:
                icon_img = extract_icon(program_path)
                if icon_img:
                    self.icon_images[program_path] = icon_img
        if icon_img:
            icon_label = ctk.CTkLabel(frame, image=icon_img, text="")
            icon_label.pack(pady=(6,2))
        label = ctk.CTkLabel(frame, text=key, font=("Segoe UI", 14, "bold"))
        label.pack(pady=2)
        btn = ctk.CTkButton(
            frame,
            text=os.path.basename(program_path).split(".")[0] if program_path else "",
            width=70,
            height=70,
            fg_color="#1e1e1e",
            command=lambda k=key: self.open_program_window(k)
        )
        btn.bind("<Button-3>", lambda e, k=key: self.unassign_callback(k))
        btn.pack(pady=6)
        self.buttons[key] = btn

    def update_button_name(self, key, path):
        btn = self.buttons.get(key)
        if btn:
            btn.configure(text=os.path.basename(path).split(".")[0])

# =============================================
# VENTANA SELECCIÓN PROGRAMAS
# =============================================
class ProgramSelectionWindow(ctk.CTkToplevel):
    def __init__(self, master, assign_callback):
        super().__init__(master)
        self.title("Selecciona Programa")
        self.geometry("500x600")
        self.assign_callback = assign_callback
        self.search_var = tk.StringVar()
        self.all_programs = get_programs()

        label = ctk.CTkLabel(self, text="Selecciona programa", font=("Segoe UI", 14, "bold"))
        label.pack(pady=6)
        entry = ctk.CTkEntry(self, placeholder_text="Buscar...", textvariable=self.search_var)
        entry.pack(fill="x", padx=6)
        entry.bind("<KeyRelease>", self.update_list)

        self.listbox = tk.Listbox(self)
        self.listbox.pack(fill="both", expand=True, padx=6, pady=6)
        self.listbox.bind("<Double-Button-1>", self.on_double_click)

        self.update_list()

    def update_list(self, event=None):
        query = self.search_var.get().lower()
        filtered = [p for p in self.all_programs if query in os.path.basename(p).lower()]
        self.listbox.delete(0, "end")
        for p in filtered:
            self.listbox.insert("end", os.path.basename(p))
        self.program_paths = filtered

    def on_double_click(self, event=None):
        idx = self.listbox.curselection()
        if not idx:
            return
        selected_path = self.program_paths[idx[0]]
        self.assign_callback(selected_path)
        self.destroy()

# =============================================
# PESTAÑA CONFIGURACIÓN
# =============================================
class ConfigPage(ctk.CTkFrame):
    def __init__(self, master, config, update_callback):
        super().__init__(master)
        self.config = config
        self.update_callback = update_callback
        self.create_ui()

    def create_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        title = ctk.CTkLabel(self, text="Opciones de Configuración", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.start_var = ctk.BooleanVar(value=self.config.get("start_with_windows", False))
        self.start_check = ctk.CTkCheckBox(
            self, text="Iniciar con Windows", variable=self.start_var, command=self.update_settings
        )
        self.start_check.grid(row=1, column=0, padx=20, pady=10, sticky="w")

        self.minimized_var = ctk.BooleanVar(value=self.config.get("start_minimized", False))
        self.minimized_check = ctk.CTkCheckBox(
            self, text="Iniciar minimizado en bandeja", variable=self.minimized_var, command=self.update_settings
        )
        self.minimized_check.grid(row=2, column=0, padx=20, pady=10, sticky="w")

        ctk.CTkLabel(self, text="Modo de Apariencia:", font=("Segoe UI", 12, "bold")).grid(
            row=3, column=0, padx=20, pady=(20,5), sticky="w"
        )
        self.appearance_option = ctk.CTkOptionMenu(
            self, values=["Light","Dark","System"], command=self.change_appearance
        )
        self.appearance_option.set(self.config.get("appearance_mode","Dark").capitalize())
        self.appearance_option.grid(row=4, column=0, padx=20, pady=(5,20), sticky="w")

    def change_appearance(self, value):
        ctk.set_appearance_mode(value)
        self.config["appearance_mode"] = value.lower()
        self.update_settings()

    def update_settings(self):
        self.config["start_with_windows"] = self.start_var.get()
        self.config["start_minimized"] = self.minimized_var.get()
        self.update_callback(self.config)
        try:
            set_startup(self.start_var.get())
        except Exception as e:
            messagebox.showerror("Error Startup", f"No se pudo cambiar inicio automático.\n{e}")

# =============================================
# APLICACIÓN PRINCIPAL
# =============================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Macropad F13-F24")
        self.config_data = load_config()
        self.program_assignments = self.config_data.get("assignments", {})
        self.geometry("1000x500")
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        # Botones de navegación
        self.nav_frame = ctk.CTkFrame(self)
        self.nav_frame.pack(side="top", fill="x")
        self.launch_btn = ctk.CTkButton(self.nav_frame, text="Lanzador", command=self.show_launcher)
        self.launch_btn.pack(side="left", padx=6, pady=6)
        self.config_btn = ctk.CTkButton(self.nav_frame, text="Configuración", command=self.show_config)
        self.config_btn.pack(side="left", padx=6, pady=6)

        info_label = ctk.CTkLabel(self, text="click derecho para desconfigurar", font=("Segoe UI", 12), text_color="gray")
        info_label.pack(pady=(4,10))

        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)

        self.launcher_page = LaunchGridPage(self.container, self.program_assignments, self.open_program_window, self.unassign_program)
        self.launcher_page.pack(fill="both", expand=True)

        self.config_page = ConfigPage(self.container, self.config_data, self.save_config)
        self.selected_key = None
        ctk.set_appearance_mode(self.config_data.get("appearance_mode","dark"))

        # Mostrar minimizado si corresponde
        if self.config_data.get("start_minimized", False):
            self.after(100, self.hide_to_tray)

    # Mostrar páginas
    def show_launcher(self):
        self.config_page.pack_forget()
        self.launcher_page.pack(fill="both", expand=True)

    def show_config(self):
        self.launcher_page.pack_forget()
        self.config_page.pack(fill="both", expand=True)

    # Programas
    def open_program_window(self, key):
        self.selected_key = key
        ProgramSelectionWindow(self, self.assign_program_to_button)

    def assign_program_to_button(self, path):
        if self.selected_key:
            self.program_assignments[self.selected_key] = path
            self.launcher_page.update_button_name(self.selected_key, path)
            self.selected_key = None
            self.save_config()

    def unassign_program(self, key):
        if key in self.program_assignments:
            del self.program_assignments[key]
            self.refresh()

    def refresh(self):
        self.launcher_page.pack_forget()
        self.launcher_page = LaunchGridPage(self.container, self.program_assignments, self.open_program_window, self.unassign_program)
        self.launcher_page.pack(fill="both", expand=True)

    def save_config(self, new_config=None):
        if new_config:
            self.config_data.update(new_config)
        self.config_data["assignments"] = self.program_assignments
        save_config(self.config_data)

    # =============================================
    # BANDEJA DEL SISTEMA
    # =============================================
    def hide_to_tray(self):
        self.withdraw()
        try:
            image = Image.open("app_icon.png")
        except FileNotFoundError:
            image = Image.new("RGB", (64,64), "blue")
            draw = ImageDraw.Draw(image)
            draw.rectangle((0,0,64,64), fill="blue")

        menu = pystray.Menu(
            pystray.MenuItem("Abrir", lambda icon, item: self.show_window(icon)),
            pystray.MenuItem("Salir", lambda icon, item: self.quit_app(icon))
        )
        self.tray_icon = pystray.Icon("APPVIAL", image, "APPVIAL", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self, icon):
        icon.stop()
        self.deiconify()

    def quit_app(self, icon):
        icon.stop()
        self.destroy()

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
