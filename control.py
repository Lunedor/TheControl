import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import subprocess
import os
import sys
import psutil
import ctypes
from ctypes import wintypes
from PIL import Image
import pystray
import threading
import json
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Dict
from collections import defaultdict
import time

# Windows registry for startup
try:
    import winreg
except Exception:
    winreg = None

# ==============================================================================
# --- 0. DATA STRUCTURES & PERSISTENCE ---
# ==============================================================================
APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "wallch_settings.json"
STATUS_FILE = APP_DIR / "wallch.status"
CMD_FILE = APP_DIR / "wallch.cmd"
APPS_CONFIG_FILE = APP_DIR / "apps.json"

# NEW: state/profiles/logs
STATE_FILE = APP_DIR / "control.state.json"
PROFILES_FILE = APP_DIR / "profiles.json"
LOG_DIR = APP_DIR / "logs"

@dataclass
class AppConfig:
    name: str
    process_name: str
    type: str = "standard"
    cwd: Optional[str] = None
    command: Optional[str] = None
    path: Optional[str] = None
    script: Optional[str] = None

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items()}

DEFAULT_WALLCH_SETTINGS = {
    "folder": "", "interval": 300, "style": "fill",
    "shuffle": True, "recursive": False, "once": False
}

def _ensure_dirs():
    try:
        LOG_DIR.mkdir(exist_ok=True)
    except Exception:
        pass

def _kill_other_control_instances():
    """Terminate any other running control.py instances (python/pythonw) except this PID."""
    my_pid = os.getpid()
    targets = []
    for p in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if p.pid == my_pid:
                continue
            name = (p.info.get('name') or '').lower()
            if name not in ('python.exe', 'pythonw.exe'):
                continue
            cmd = [c.lower() for c in (p.info.get('cmdline') or []) if c]
            if any('control.py' in part or 'control.pyw' in part for part in cmd):
                targets.append(p)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    for p in targets:
        try: p.terminate()
        except Exception: pass
    gone, alive = psutil.wait_procs(targets, timeout=2.5)
    for p in alive:
        try: p.kill()
        except Exception: pass

# --- basic IO ---
def load_wallch_settings():
    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            data.setdefault("folder", "")
            return data
        except Exception:
            pass
    return DEFAULT_WALLCH_SETTINGS.copy()

def save_wallch_settings(data: dict):
    try:
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"Failed to save settings: {e}")

def load_apps_from_json() -> List[AppConfig]:
    if not APPS_CONFIG_FILE.exists():
        return []
    try:
        data = json.loads(APPS_CONFIG_FILE.read_text(encoding="utf-8"))
        for app_data in data:
            if app_data.get('script') == 'wallch.py':
                settings = load_wallch_settings()
                app_data['command'] = build_wallch_command(settings)
                if app_data.get('cwd') == '.':
                    app_data['cwd'] = str(APP_DIR)
        return [AppConfig(**item) for item in data]
    except (json.JSONDecodeError, TypeError) as e:
        messagebox.showerror("Config Error", f"Failed to load 'apps.json':\n{e}")
        return []

def save_apps_to_json(apps: List[AppConfig]):
    try:
        app_list_to_save = []
        for app in apps:
            app_dict = app.to_dict()
            if app_dict.get('script') == 'wallch.py':
                app_dict['command'] = ""
                if app_dict.get('cwd') == str(APP_DIR):
                    app_dict['cwd'] = "."
            app_list_to_save.append(app_dict)
        APPS_CONFIG_FILE.write_text(json.dumps(app_list_to_save, indent=2), encoding="utf-8")
    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save to 'apps.json':\n{e}")

def build_wallch_command(settings: dict) -> str:
    folder, interval, style = settings["folder"], int(settings["interval"]), settings["style"]
    args = [f'"{folder}"', f"--interval {max(1, interval)}", f"--style {style}"]
    if settings["shuffle"]:   args.append("--shuffle")
    if settings["recursive"]: args.append("--recursive")
    if settings["once"]:      args.append("--once")
    return f'start "" pythonw.exe wallch.py ' + " ".join(args)

def read_wallch_status() -> str:
    try:
        return STATUS_FILE.read_text(encoding="utf-8").strip() if STATUS_FILE.exists() else "Unknown"
    except Exception:
        return "Unknown"

def send_wallch_command(text: str):
    try:
        CMD_FILE.write_text(text.strip() + "\n", encoding="utf-8")
    except Exception as e:
        print(f"Failed to send command '{text}': {e}")

# ==============================================================================
# --- 1. PROCESS MANAGEMENT + LOGGING + AUTORESTART
# ==============================================================================
def _aggregate_proc_dict():
    """Collect ALL processes by name -> [psutil.Process, ...] (no overwrites)."""
    procs = defaultdict(list)
    for p in psutil.process_iter(['name', 'cmdline']):
        try:
            nm = p.info.get('name')
            if nm:
                procs[nm].append(p)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return procs

def find_process(app_config: AppConfig, proc_dict: dict) -> Optional[psutil.Process]:
    procs = proc_dict.get(app_config.process_name, [])
    if not procs:
        return None
    if app_config.script:
        target_script = app_config.script.lower()
        for proc in procs:
            if _match_script(proc, target_script):
                return proc
        return None
    return procs[0] if procs else None

def _log_line(app_name: str, text: str):
    """Append a line to logs/<appname>.log and keep last 50 lines."""
    try:
        _ensure_dirs()
        p = LOG_DIR / f"{app_name}.log"
        line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {text}\n"
        if p.exists():
            # keep last 49, add new one = 50
            lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
            lines = lines[-49:] + [line.rstrip("\n")]
            p.write_text("\n".join(lines) + "\n", encoding="utf-8")
        else:
            p.write_text(line, encoding="utf-8")
    except Exception:
        pass

def start_app(app_config: AppConfig):
    print(f"Starting {app_config.name}...")
    try:
        proc_dict = _aggregate_proc_dict()
        existing = find_process(app_config, proc_dict)
        if existing:
            print(f"{app_config.name} already running (pid={existing.pid}); skip start.")
            return

        if app_config.path and os.path.exists(app_config.path):
            subprocess.Popen([app_config.path])
        elif app_config.command:
            subprocess.Popen(app_config.command, shell=True, cwd=app_config.cwd, creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            messagebox.showerror("Error", f"No valid path or command for {app_config.name}.")
            return
        _log_line(app_config.name, "started")
    except Exception as e:
        messagebox.showerror("Start Error", f"Failed to start {app_config.name}:\n{e}")

def _match_script(proc: psutil.Process, script_name: str) -> bool:
    """
    Precise matching: Returns True ONLY if script_name is found in the 
    command line arguments.
    """
    try:
        # cmdline() returns a list like ['python.exe', 'C:\\Path\\wallch.py', '--interval', '60']
        cmd_args = proc.cmdline()
        
        # We convert everything to lowercase for case-insensitive comparison
        # We verify that 'python' is likely the executable (optional, but safer)
        exe_name = proc.name().lower()
        
        if 'python' not in exe_name:
            # If it's not python, we don't check for script names (unless you wrap other languages)
            return False

        target = script_name.lower()
        
        # Check every argument in the command line
        for arg in cmd_args:
            if target in arg.lower():
                return True
                
        return False
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return False

def _kill_process_tree(proc: psutil.Process):
    """
    Kills the process and all its children (essential for shell=True commands like Rclone).
    """
    try:
        children = proc.children(recursive=True)
        for child in children:
            try: child.terminate()
            except: pass
        
        # Wait briefly for children to die
        _, alive = psutil.wait_procs(children, timeout=0.5)
        for child in alive:
            try: child.kill()
            except: pass
            
        # Now kill the parent
        proc.terminate()
        try:
            proc.wait(timeout=1)
        except psutil.TimeoutExpired:
            proc.kill()
    except psutil.NoSuchProcess:
        pass

def stop_app(app_config: AppConfig, proc: psutil.Process):
    print(f"Stopping {app_config.name}...")
    
    targets = []

    # CASE A: We have a specific process object handled by the UI
    if proc and proc.is_running():
        # Double check: If it's a python script, ensure we didn't get the PIDs mixed up
        if app_config.script:
            if _match_script(proc, app_config.script):
                targets.append(proc)
            else:
                print(f"Warning: Stored PID {proc.pid} no longer matches script {app_config.script}. Scanning system...")
                # PID might have been recycled, fall through to Case B
        else:
            targets.append(proc)

    # CASE B: We don't have a proc handle (or it was invalid), so we scan the system safely
    if not targets:
        for p in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # 1. Check if process name matches (e.g., "pythonw.exe" or "stremio.exe")
                if p.info['name'].lower() != app_config.process_name.lower():
                    continue

                # 2. If it's a script, we MUST match the script name in args
                if app_config.script:
                    if _match_script(p, app_config.script):
                        # SPECIAL SAFETY: Don't let control.py kill itself!
                        if p.pid == os.getpid():
                            continue
                        targets.append(p)
                
                # 3. If it's a regular app (not script), matching process name is enough
                else:
                    targets.append(p)

            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

    # Execute the Kill Order
    if not targets:
        print(f"No running process found for {app_config.name}")
        _log_line(app_config.name, "stop requested but not found")
        return

    for p in targets:
        print(f"Killing PID {p.pid} ({app_config.name})")
        _kill_process_tree(p)
    
    _log_line(app_config.name, "stopped")

# ==============================================================================
# --- 2. STATE (remember last ON/OFF & autostart) + PROFILES
# ==============================================================================
def _load_state() -> dict:
    try:
        if STATE_FILE.exists():
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {"desired": {}, "autostart": False, "last_profile": None}

def _save_state(state: dict):
    try:
        STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")
    except Exception:
        pass

def _load_profiles() -> Dict[str, Dict[str, bool]]:
    """profiles.json structure:
    {
        "Work": {"App A": true, "App B": false},
        "Chill": {"Wallpaper": true}
    }
    """
    try:
        if PROFILES_FILE.exists():
            data = json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                # normalize to bools
                for prof, mapping in data.items():
                    for k, v in list(mapping.items()):
                        mapping[k] = bool(v)
                return data
    except Exception:
        pass
    return {}  # user defines as needed

def _set_startup_enabled(enable: bool):
    if not winreg:
        return
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_SET_VALUE)
        name = "TheControl"
        if enable:
            py = sys.executable.replace("python.exe", "pythonw.exe")
            cmd = f'"{py}" "{Path(__file__).resolve()}"'
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, cmd)
        else:
            try:
                winreg.DeleteValue(key, name)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception:
        pass

def _get_startup_enabled() -> bool:
    if not winreg:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_READ)
        name = "TheControl"
        val, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return bool(val)
    except FileNotFoundError:
        return False
    except Exception:
        return False

def _save_profiles(profiles: dict) -> bool:
    try:
        PROFILES_FILE.write_text(json.dumps(profiles, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False

# ==============================================================================
# --- 3. GUI
# ==============================================================================
def position_dialog(dialog: tk.Toplevel, parent: tk.Tk):
    dialog.update_idletasks()
    w, h = dialog.winfo_width(), dialog.winfo_height()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    sw, sh = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
    x = max(0, min(px + (pw - w) // 2, sw - w))
    y = max(0, min(py + (ph - h) // 2, sh - h))
    dialog.geometry(f"+{x}+{y}")

def set_dark_title_bar(window):
    try:
        window.update()
        hwnd = ctypes.windll.user32.FindWindowW(None, window.title())
        if hwnd:
            value = ctypes.c_int(1)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass

class WallpaperSettingsDialog(tk.Toplevel):
    def __init__(self, parent, app_config, current_settings: dict):
        super().__init__(parent)
        self.title("Wallpaper Settings")
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()
        self.var_folder = tk.StringVar(value=current_settings["folder"])
        self.var_interval = tk.IntVar(value=current_settings["interval"])
        self.var_style = tk.StringVar(value=current_settings["style"])
        self.var_shuffle = tk.BooleanVar(value=current_settings["shuffle"])
        self.var_recursive = tk.BooleanVar(value=current_settings["recursive"])
        self.var_once = tk.BooleanVar(value=current_settings["once"])
        frm = ttk.Frame(self, padding=15)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Folder:", font="-weight bold").grid(row=0, column=0, sticky="w", pady=(0, 6))
        row1 = ttk.Frame(frm)
        row1.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0,10)); row1.columnconfigure(0, weight=1)
        ttk.Entry(row1, textvariable=self.var_folder).grid(row=0, column=0, sticky="ew", padx=(0,8))
        ttk.Button(row1, text="Browse…", command=self.browse_folder, bootstyle=SECONDARY).grid(row=0, column=1)
        ttk.Label(frm, text="Interval (seconds):", font="-weight bold").grid(row=2, column=0, sticky="w")
        ttk.Spinbox(frm, from_=1, to=86400, textvariable=self.var_interval, width=12).grid(row=3, column=0, sticky="w", pady=(0,10))
        ttk.Label(frm, text="Style:", font="-weight bold").grid(row=4, column=0, sticky="w")
        ttk.Combobox(frm, state="readonly", textvariable=self.var_style, values=["fill","fit","stretch","center","tile","span"], width=12).grid(row=5, column=0, sticky="w", pady=(0,10))
        chk_row = ttk.Frame(frm)
        chk_row.grid(row=6, column=0, columnspan=3, sticky="w", pady=(0,10))
        ttk.Checkbutton(chk_row, text="Shuffle", variable=self.var_shuffle, bootstyle="success").pack(side=tk.LEFT, padx=(0,16))
        ttk.Checkbutton(chk_row, text="Recursive", variable=self.var_recursive, bootstyle="info").pack(side=tk.LEFT, padx=(0,16))
        ttk.Checkbutton(chk_row, text="Once", variable=self.var_once, bootstyle="warning").pack(side=tk.LEFT)
        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=3, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btns, text="Save & Apply", command=lambda: self.save(app_config, apply_now=True), bootstyle=SUCCESS).pack(side=tk.RIGHT)
        position_dialog(self, parent)

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.var_folder.get() or str(APP_DIR))
        if folder: self.var_folder.set(folder)

    def save(self, app_config, apply_now=False):
        settings = {"folder": self.var_folder.get().strip(), "interval": self.var_interval.get(), "style": self.var_style.get(), "shuffle": self.var_shuffle.get(), "recursive": self.var_recursive.get(), "once": self.var_once.get()}
        if not settings["folder"] or not os.path.isdir(settings["folder"]):
            messagebox.showerror("Invalid folder", "Please choose a valid folder.", parent=self); return
        save_wallch_settings(settings)
        app_config.command = build_wallch_command(settings)
        if apply_now:
            proc_dict = _aggregate_proc_dict()
            proc = find_process(app_config, proc_dict)
            if proc:
                stop_app(app_config, proc)
                self.after(500, lambda: start_app(app_config))
            else:
                start_app(app_config)
        self.destroy()

class ProfilesManagerDialog(tk.Toplevel):
    """UI to create/rename/delete profiles and set per-app action: Leave / Start / Stop."""
    def __init__(self, parent, profiles: dict, apps: list[str]):
        super().__init__(parent)
        self.title("Manage Profiles")
        self.transient(parent); self.grab_set()
        self.resizable(True, True)

        self.apps = apps[:]                       # list of app names
        self.profiles = {k: dict(v) for k, v in profiles.items()}  # name -> {app: bool}
        self.result = None                         # filled on Save

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)

        # Header
        ttk.Label(outer, text="Profiles", font="-size 14 -weight bold").grid(row=0, column=0, sticky="w")
        ttk.Label(outer, text="Apps in profile (Leave/Start/Stop)", font="-size 14 -weight bold").grid(row=0, column=1, sticky="w")

        # Left: profiles list + buttons
        left = ttk.Frame(outer)
        left.grid(row=1, column=0, sticky="nsw", padx=(0,12))
        self.lb = tk.Listbox(left, height=14)
        self.lb.pack(fill=tk.BOTH, expand=True)
        self.lb.configure(exportselection=False)  # keep selection even when focus moves
        self.current_profile: str | None = None   # sticky selection
        btns = ttk.Frame(left); btns.pack(fill=tk.X, pady=(8,0))
        ttk.Button(btns, text="Add", command=self._add_profile, bootstyle=SUCCESS).pack(side=tk.LEFT)
        ttk.Button(btns, text="Rename", command=self._rename_profile, bootstyle=INFO).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Delete", command=self._delete_profile, bootstyle=DANGER).pack(side=tk.LEFT)

        # Right: scrollable area with per-app comboboxes
        right = ttk.Frame(outer)
        right.grid(row=1, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1); right.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(right, borderwidth=0, highlightthickness=0)
        self.sframe = ttk.Frame(self.canvas)
        self.swin = self.canvas.create_window((0,0), window=self.sframe, anchor="nw")
        sc = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sc.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        sc.grid(row=0, column=1, sticky="ns")

        self.sframe.bind("<Configure>", lambda e: (self.canvas.configure(scrollregion=self.canvas.bbox("all")),
                                                   self.canvas.itemconfig(self.swin, width=self.canvas.winfo_width())))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.swin, width=e.width))

        # Per-app controls: app -> tk.StringVar("Leave"/"Start"/"Stop")
        self.app_modes: dict[str, tk.StringVar] = {}
        for app in self.apps:
            row = ttk.Frame(self.sframe); row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=app).pack(side=tk.LEFT)
            var = tk.StringVar(value="Leave")
            cb = ttk.Combobox(row, state="readonly", width=8, textvariable=var, values=["Leave","Start","Stop"])
            cb.pack(side=tk.RIGHT)
            self.app_modes[app] = var

        # Footer buttons
        footer = ttk.Frame(outer); footer.grid(row=2, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(footer, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(footer, text="Save", command=self._save, bootstyle=SUCCESS).pack(side=tk.RIGHT)

        # Load initial profiles
        for name in sorted(self.profiles.keys()):
            self.lb.insert(tk.END, name)
        self.lb.bind("<<ListboxSelect>>", lambda e: self._load_selected())
        if self.lb.size() > 0:
            self.lb.selection_set(0)
            self._load_selected()

        self.update_idletasks()
        self.minsize(560, 460)
        position_dialog(self, parent)

    # Helpers
    def _load_selected(self):
        sel = self.lb.curselection()
        if not sel:
            self.current_profile = None
            for v in self.app_modes.values(): 
                v.set("Leave")
            return
        name = self.lb.get(sel[0])
        self.current_profile = name
        mapping = self.profiles.get(name, {})
        for app, var in self.app_modes.items():
            if app in mapping:
                var.set("Start" if mapping[app] else "Stop")
            else:
                var.set("Leave")

    def _commit_current_to_profiles(self):
        name = self.current_profile
        if not name:
            return
        # Only store entries that are Start or Stop; omit Leave
        mapping = {}
        for app, var in self.app_modes.items():
            mode = var.get()
            if mode == "Start":
                mapping[app] = True
            elif mode == "Stop":
                mapping[app] = False
        self.profiles[name] = mapping


    def _add_profile(self):
        name = simpledialog.askstring("New Profile", "Profile name:", parent=self)
        if not name: return
        name = name.strip()
        if not name: return
        if name in self.profiles:
            messagebox.showerror("Exists", f"Profile '{name}' already exists.", parent=self); return
        self._commit_current_to_profiles()  # save edits of previous profile
        self.profiles[name] = {}
        self.lb.insert(tk.END, name)
        self.lb.selection_clear(0, tk.END)
        self.lb.selection_set(tk.END)
        self.current_profile = name          # <-- keep sticky
        self._load_selected()


    def _rename_profile(self):
        sel = self.lb.curselection()
        if not sel: return
        old = self.lb.get(sel[0])
        new = simpledialog.askstring("Rename Profile", "New name:", initialvalue=old, parent=self)
        if not new: return
        new = new.strip()
        if not new or new == old: return
        if new in self.profiles:
            messagebox.showerror("Exists", f"Profile '{new}' already exists.", parent=self); return
        self._commit_current_to_profiles()
        self.profiles[new] = self.profiles.pop(old)
        self.lb.delete(sel[0]); self.lb.insert(sel[0], new); self.lb.selection_set(sel[0])
        self.current_profile = new           # <-- keep sticky
        self._load_selected()

    def _delete_profile(self):
        sel = self.lb.curselection()
        if not sel: return
        name = self.lb.get(sel[0])
        if not messagebox.askyesno("Delete", f"Delete profile '{name}'?", parent=self):
            return
        self.profiles.pop(name, None)
        self.lb.delete(sel[0])
        if self.lb.size() == 0:
            self.current_profile = None
            for v in self.app_modes.values(): v.set("Leave")
        else:
            self.lb.selection_set(0)
            self._load_selected()


    def _save(self):
        self._commit_current_to_profiles()   # commit current UI state
        self.result = self.profiles
        self.destroy()


class AddEditAppDialog(tk.Toplevel):
    def __init__(self, parent, app_to_edit: Optional[AppConfig] = None):
        super().__init__(parent)
        self.title("Edit App" if app_to_edit else "Add New App")
        self.transient(parent); self.grab_set()
        self.resizable(False, False)
        self.result: Optional[AppConfig] = None

        self.vars = {
            "name": tk.StringVar(value=app_to_edit.name if app_to_edit else ""),
            "process_name": tk.StringVar(value=app_to_edit.process_name if app_to_edit else ""),
            "path": tk.StringVar(value=app_to_edit.path if app_to_edit else ""),
            "command": tk.StringVar(value=app_to_edit.command if app_to_edit else ""),
            "cwd": tk.StringVar(value=app_to_edit.cwd if app_to_edit else ""),
            "script": tk.StringVar(value=app_to_edit.script if app_to_edit else "")
        }

        frm = ttk.Frame(self, padding=20)
        frm.grid(row=0, column=0, sticky="nsew")

        info_frame = ttk.Labelframe(frm, text="1. Basic Info (Required)", padding=15)
        info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        info_frame.columnconfigure(1, weight=1)
        ttk.Label(info_frame, text="App Name:").grid(row=0, column=0, sticky="w", padx=(0,10), pady=2)
        ttk.Entry(info_frame, textvariable=self.vars["name"]).grid(row=0, column=1, sticky="ew")
        ttk.Label(info_frame, text="Process Name:").grid(row=1, column=0, sticky="w", padx=(0,10), pady=2)
        ttk.Entry(info_frame, textvariable=self.vars["process_name"]).grid(row=1, column=1, sticky="ew")

        launch_frame = ttk.Labelframe(frm, text="2. Launch Method (Choose one)", padding=15)
        launch_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        launch_frame.columnconfigure(0, weight=1)
        ttk.Label(launch_frame, text="Executable Path:").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,5))
        path_entry = ttk.Entry(launch_frame, textvariable=self.vars["path"])
        path_entry.grid(row=1, column=0, sticky="ew", padx=(0,5))
        ttk.Button(launch_frame, text="Browse...", command=self.browse_path).grid(row=1, column=1, sticky="e")
        ttk.Separator(launch_frame, orient="horizontal").grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Label(launch_frame, text="Or Shell Command:").grid(row=3, column=0, columnspan=2, sticky="w", pady=(0,5))
        ttk.Entry(launch_frame, textvariable=self.vars["command"]).grid(row=4, column=0, columnspan=2, sticky="ew")

        adv_frame = ttk.Labelframe(frm, text="3. Advanced (Optional)", padding=15)
        adv_frame.grid(row=2, column=0, sticky="ew")
        adv_frame.columnconfigure(1, weight=1)
        ttk.Label(adv_frame, text="Working Dir (CWD):").grid(row=0, column=0, sticky="w", padx=(0,10), pady=2)
        ttk.Entry(adv_frame, textvariable=self.vars["cwd"]).grid(row=0, column=1, sticky="ew")
        ttk.Label(adv_frame, text="Python Script Name:").grid(row=1, column=0, sticky="w", padx=(0,10), pady=2)
        ttk.Entry(adv_frame, textvariable=self.vars["script"]).grid(row=1, column=1, sticky="ew")

        btn_bar = ttk.Frame(frm, padding=(0, 15, 0, 0))
        btn_bar.grid(row=3, column=0, sticky="e")
        ttk.Button(btn_bar, text="Cancel", command=self.destroy, bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_bar, text="Save", command=self.save, bootstyle=SUCCESS).pack(side=tk.RIGHT)

        position_dialog(self, parent)
        self.bind("<Escape>", lambda e: self.destroy())

    def browse_path(self):
        path = filedialog.askopenfilename(
            title="Select Executable",
            filetypes=[("Executables", "*.exe"), ("All files", "*.*")]
        )
        if path:
            self.vars["path"].set(path)
            # If process name empty, infer from filename
            if not self.vars["process_name"].get().strip():
                self.vars["process_name"].set(os.path.basename(path))

    def save(self):
        name = self.vars["name"].get().strip()
        proc_name = self.vars["process_name"].get().strip()
        path = self.vars["path"].get().strip()
        command = self.vars["command"].get().strip()
        if not name or not proc_name:
            messagebox.showerror("Missing Info", "App Name and Process Name are required.", parent=self); return
        if not path and not command:
            messagebox.showerror("Missing Info", "You must provide an Executable Path or a Shell Command.", parent=self); return
        if path and command:
            messagebox.showwarning("Ambiguous", "Both Path and Command are filled. The Path will be used.", parent=self)
            command = ""

        # ✅ generic, app-agnostic
        script_val = self.vars["script"].get().strip()
        if script_val:
            app_type = "python-script"
        elif path:
            app_type = "executable"
        elif command:
            app_type = "command"
        else:
            app_type = "standard"

        self.result = AppConfig(
            name=name, process_name=proc_name,
            path=path or None, command=command or None,
            cwd=self.vars["cwd"].get().strip() or None,
            script=script_val or None,
            type=app_type
        )
        self.destroy()


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, background=self._get_bg_color())
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _get_bg_color(self):
        try:
            style = ttk.Style()
            return style.lookup('TFrame', 'background')
        except:
            return "#2C2C2C"

    def update_scrollbar(self):
        self.update_idletasks()
        content_height = self.scrollable_frame.winfo_reqheight()
        canvas_height = self.canvas.winfo_height()
        if content_height > canvas_height:
            self.scrollbar.grid(row=0, column=1, sticky='ns')
        else:
            self.scrollbar.grid_forget()

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.update_scrollbar()

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_frame, width=event.width)
        self.update_scrollbar()

    def _on_mousewheel(self, event):
        widget_under_cursor = self.winfo_containing(event.x_root, event.y_root)
        if widget_under_cursor and self.canvas in widget_under_cursor.winfo_parents() or widget_under_cursor == self.canvas:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def destroy(self):
        self.unbind_all("<MouseWheel>")
        super().destroy()

class DraggableAppFrame(ttk.Frame):
    def __init__(self, parent, app_manager, app_config, index, **kwargs):
        super().__init__(parent, **kwargs)
        self.app_manager = app_manager
        self.app_config = app_config
        self.index = index
        self.bind_events()

    def bind_events(self):
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<B1-Motion>", self.on_motion)
        self.bind("<ButtonRelease-1>", self.on_release)
        for widget in self.winfo_children():
            self._bind_recursive(widget)

    def _bind_recursive(self, widget):
        widget.bind("<ButtonPress-1>", self.on_press)
        widget.bind("<B1-Motion>", self.on_motion)
        widget.bind("<ButtonRelease-1>", self.on_release)
        for child in widget.winfo_children():
            self._bind_recursive(child)

    def on_press(self, event):
        self.app_manager.start_drag(self.index, self)

    def on_motion(self, event):
        self.app_manager.handle_drag_motion(event)

    def on_release(self, event):
        self.app_manager.handle_drop()

class AppManager(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("The Control")
        set_dark_title_bar(self)
        _ensure_dirs()

        self.geometry("450x700"); self.resizable(True, True)
        self.apps: List[AppConfig] = load_apps_from_json()

        # Desired ON/OFF map (remembered & for autorestart)
        self.app_state = _load_state()
        self.desired: Dict[str, bool] = dict(self.app_state.get("desired", {}))
        # Profiles
        self.profiles = _load_profiles()
        # Autostart sync
        if _get_startup_enabled() != bool(self.app_state.get("autostart", False)):
            _set_startup_enabled(bool(self.app_state.get("autostart", False)))

        self.ui_elements = {}
        self.tray_icon = None
        self.after_id = None

        self.drag_source_index = None
        self.drag_target_index = None
        self.drag_placeholder = None

        self.main_frame = ttk.Frame(self, padding=20)
        self.main_frame.pack(expand=True, fill=tk.BOTH)

        self.rebuild_ui()
        self.setup_tray_icon()
        self.protocol("WM_DELETE_WINDOW", self.hide_window)

        # Apply remembered state on launch
        self.after(50, self.apply_desired_on_launch)
        self.after(250, self.update_statuses)

        first_run = not STATE_FILE.exists()
        if first_run:
            # show once on first run, after widgets have measured their size
            self.after(120, self.show_window)
        else:
            self.withdraw()
    
    def open_profiles_manager(self):
        app_names = [a.name for a in self.apps]
        dlg = ProfilesManagerDialog(self, self.profiles, app_names)
        self.wait_window(dlg)
        if dlg.result is not None:
            self.profiles = dlg.result
            _save_profiles(self.profiles)
            self.refresh_tray_menu()

    def apply_desired_on_launch(self):
        for app in self.apps:
            want = self.desired.get(app.name, False)
            if want:
                start_app(app)
        self.update_statuses()

    def position_window(self):
        self.update_idletasks()
        width, height = self.winfo_width(), self.winfo_height()
        work_area = wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(work_area), 0)
        x = work_area.right - width - 20
        y = work_area.bottom - height - 30
        self.geometry(f'+{x}+{y}')

    def rebuild_ui(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        self.ui_elements.clear()

        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header_frame, text="Application Status", font="-size 16 -weight bold", bootstyle=PRIMARY).pack(side=tk.LEFT)

        btn_bar = ttk.Frame(header_frame)
        btn_bar.pack(side=tk.RIGHT)
        ttk.Button(btn_bar, text="Manage Profiles", command=self.open_profiles_manager, bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btn_bar, text="Manage Apps", command=self.open_app_manager, bootstyle=INFO).pack(side=tk.RIGHT)

        controls_frame = ScrollableFrame(self.main_frame)
        controls_frame.pack(expand=True, fill=tk.BOTH)
        content_area = controls_frame.scrollable_frame

        self.drag_placeholder = ttk.Frame(content_area, height=6, bootstyle='info')

        for i, app_config in enumerate(self.apps):
            app_frame = DraggableAppFrame(content_area, self, app_config, i, padding=(10, 8))
            app_frame.pack(fill=tk.X)

            info_frame = ttk.Frame(app_frame)
            info_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
            name_lbl = ttk.Label(info_frame, text=app_config.name, font="-size 12")
            name_lbl.pack(side=tk.LEFT, anchor="w")
            stats_lbl = ttk.Label(info_frame, text="", font="-size 9", bootstyle=INFO)
            stats_lbl.pack(side=tk.RIGHT, anchor="e", padx=10)
            switch = ttk.Checkbutton(app_frame, bootstyle="success,round-toggle",
                                     command=lambda cfg=app_config: self.toggle_app(cfg))
            switch.pack(side=tk.RIGHT, anchor="e")

            self.ui_elements[app_config.name] = {'switch': switch, 'stats_lbl': stats_lbl, 'proc': None, 'frame': app_frame}
            app_frame.bind_events()

            if app_config.script == 'wallch.py':
                ttk.Separator(content_area, orient='horizontal').pack(fill=tk.X, pady=(5, 10))
                btn_bar = ttk.Frame(content_area)
                btn_bar.pack(fill=tk.X, pady=(0, 10))
                btn_settings = ttk.Button(btn_bar, text="⚙", command=lambda cfg=app_config: self.open_wallpaper_settings(cfg), bootstyle=SECONDARY)
                btn_settings.pack(side=tk.LEFT, padx=10)
                btn_toggle = ttk.Button(btn_bar, text="▶", command=self.wallch_toggle, bootstyle=INFO)
                btn_toggle.pack(side=tk.LEFT, padx=6)
                btn_next = ttk.Button(btn_bar, text="⏭", command=self.wallch_next, bootstyle=SUCCESS)
                btn_next.pack(side=tk.LEFT, padx=6)
                self.ui_elements[app_config.name].update({'btn_toggle': btn_toggle, 'btn_next': btn_next})
                ttk.Separator(content_area, orient='horizontal').pack(fill=tk.X, pady=5)

        if tk.Tk.state(self) == "normal":
            self.position_window()

    # --- Drag and Drop Handler Methods ---
    def start_drag(self, index, frame_widget):
        self.drag_source_index = index
        self.config(cursor="hand2")

    def handle_drag_motion(self, event):
        if self.drag_source_index is None: return
        target_widget = event.widget.winfo_containing(event.x_root, event.y_root)
        while target_widget and not isinstance(target_widget, DraggableAppFrame):
            target_widget = target_widget.master
        if target_widget and target_widget.index != self.drag_source_index:
            self.drag_target_index = target_widget.index
            self.drag_placeholder.pack_forget()
            self.drag_placeholder.pack(before=target_widget, fill='x', padx=5, pady=2)

    def handle_drop(self):
        if self.drag_source_index is not None and self.drag_target_index is not None and self.drag_source_index != self.drag_target_index:
            item_to_move = self.apps.pop(self.drag_source_index)
            if self.drag_source_index < self.drag_target_index:
                self.drag_target_index -= 1
            self.apps.insert(self.drag_target_index, item_to_move)
            save_apps_to_json(self.apps)
            self.rebuild_ui()
        self.drag_placeholder.pack_forget()
        self.drag_source_index = None
        self.drag_target_index = None
        self.config(cursor="")

    def open_app_manager(self):
        manager_dialog = tk.Toplevel(self)
        manager_dialog.title("Manage Applications")
        manager_dialog.transient(self)
        manager_dialog.grab_set()

        modified = False  # track whether anything changed

        listbox = tk.Listbox(manager_dialog, height=15, width=50, selectmode=tk.SINGLE)
        listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        for app in self.apps:
            listbox.insert(tk.END, app.name)

        btn_frame = ttk.Frame(manager_dialog, padding=10)
        btn_frame.pack(fill=tk.X)

        def add_app():
            nonlocal modified
            dialog = AddEditAppDialog(self)
            self.wait_window(dialog)
            if dialog.result:
                self.apps.append(dialog.result)
                listbox.insert(tk.END, dialog.result.name)
                modified = True

        def edit_app():
            nonlocal modified
            if not listbox.curselection():
                return
            idx = listbox.curselection()[0]
            dialog = AddEditAppDialog(self, app_to_edit=self.apps[idx])
            self.wait_window(dialog)
            if dialog.result:
                self.apps[idx] = dialog.result
                listbox.delete(idx)
                listbox.insert(idx, dialog.result.name)
                modified = True

        def remove_app():
            nonlocal modified
            if not listbox.curselection():
                return
            idx = listbox.curselection()[0]
            if messagebox.askyesno("Confirm", f"Remove '{self.apps[idx].name}'?"):
                del self.apps[idx]
                listbox.delete(idx)
                modified = True

        saved = {'flag': False}  # closure-friendly box

        def do_save_and_close():
            saved['flag'] = True
            manager_dialog.destroy()

        ttk.Button(btn_frame, text="Add", command=add_app, bootstyle=SUCCESS).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Edit", command=edit_app, bootstyle=INFO).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remove", command=remove_app, bootstyle=DANGER).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Save & Close", command=do_save_and_close, bootstyle=PRIMARY).pack(side=tk.RIGHT)

        # If user clicks “X”, it should act like Cancel (no save)
        manager_dialog.protocol("WM_DELETE_WINDOW", manager_dialog.destroy)

        position_dialog(manager_dialog, self)
        self.wait_window(manager_dialog)

        # Only persist+rebuild if saved and something actually changed
        if saved['flag'] and modified:
            save_apps_to_json(self.apps)
            self.rebuild_ui()
            # Force an immediate status refresh to avoid any “all off” frame
            try:
                if self.after_id:
                    self.after_cancel(self.after_id)
            except Exception:
                pass
            self.update_statuses()

    def toggle_app(self, app_config: AppConfig):
        elements = self.ui_elements.get(app_config.name)
        if not elements: return
        btn = elements['switch']
        try: btn.configure(state=tk.DISABLED)
        except Exception: pass

        currently_on = elements['proc'] is not None
        if currently_on:
            stop_app(app_config, elements['proc'])
            self.desired[app_config.name] = False
            _log_line(app_config.name, "toggled OFF")
        else:
            start_app(app_config)
            self.desired[app_config.name] = True
            _log_line(app_config.name, "toggled ON")

        self.app_state["desired"] = self.desired
        _save_state(self.app_state)

        if self.after_id: self.after_cancel(self.after_id)
        self.after(500, self.update_statuses)
        self.after(400, lambda: btn.configure(state=tk.NORMAL))

    # --- PROFILES ---
    def apply_profile(self, profile_name: str | None):
        """Apply a profile non-destructively. Only apps listed in the profile are changed.
        Apps NOT listed are left exactly as they are. If profile_name is None -> clear selection."""
        if not profile_name:
            # Clear selected profile (no changes to running apps), just forget the mark
            self.app_state["last_profile"] = None
            _save_state(self.app_state)
            self.refresh_tray_menu()
            return

        mapping = self.profiles.get(profile_name, {})
        if not mapping:
            return

        # Only touch apps mentioned in mapping; leave others alone.
        for app in self.apps:
            if app.name not in mapping:
                continue
            want = bool(mapping[app.name])
            have = self.ui_elements.get(app.name, {}).get('proc') is not None
            # Update desired ONLY for the apps we touch
            self.desired[app.name] = want
            if want and not have:
                start_app(app); _log_line(app.name, f"profile '{profile_name}': start")
            elif not want and have:
                stop_app(app, self.ui_elements.get(app.name, {}).get('proc')); _log_line(app.name, f"profile '{profile_name}': stop")

        self.app_state["desired"] = self.desired
        self.app_state["last_profile"] = profile_name
        _save_state(self.app_state)
        self.after(300, self.update_statuses)
        self.refresh_tray_menu()


    # --- AUTORESTART + STATS ---
    def update_statuses(self):
        # 1. Only iterate all processes if we absolutely have to find a missing PID
        # We do this lazily inside the loop below only when needed.
        
        running_names_cache = None 

        for app_config in self.apps:
            elements = self.ui_elements.get(app_config.name)
            if not elements: continue
            
            cached_proc = elements.get('proc')
            
            # A. Fast Path: We already have a handle, check if it's alive
            is_alive = False
            if cached_proc:
                try:
                    if cached_proc.is_running():
                        # Double check it hasn't been reused by a different process name
                        # (rare but possible)
                        if cached_proc.name().lower() == app_config.process_name.lower():
                            is_alive = True
                    else:
                        is_alive = False
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    is_alive = False
            
            # B. Slow Path: We don't have a handle, or it died. Scan system.
            if not is_alive:
                # Only fetch the massive system list ONCE per update cycle, and only if needed
                if running_names_cache is None:
                    running_names_cache = _aggregate_proc_dict() # Your existing helper
                
                new_proc = find_process(app_config, running_names_cache)
                if new_proc:
                    elements['proc'] = new_proc
                    is_alive = True
                    
            # Update UI based on is_alive
            if is_alive:
                # Update stats...
                try:
                    p = elements['proc']
                    cpu = p.cpu_percent()
                    mem = p.memory_info().rss / (1024 * 1024)
                    elements['stats_lbl'].config(text=f"CPU: {cpu:.1f}% | Mem: {mem:.1f} MB")
                    elements['switch'].state(['selected'])
                except:
                    elements['proc'] = None # Handle race condition where it dies mid-check
            else:
                elements['switch'].state(['!selected'])
                elements['stats_lbl'].config(text="")
                elements['proc'] = None
                
                # Auto-restart logic here...
                if self.desired.get(app_config.name, False):
                    start_app(app_config)
                    _log_line(app_config.name, "auto-restart (not running)")

        self.update_wallch_ui()
        self.after_id = self.after(2500, self.update_statuses)

    def update_wallch_ui(self):
        wallch_app = next((app for app in self.apps if app.script == 'wallch.py'), None)
        if not (wallch_app and wallch_app.name in self.ui_elements): return
        elements = self.ui_elements[wallch_app.name]
        if 'btn_toggle' not in elements: return
        is_running = elements['proc'] is not None
        state = read_wallch_status() if is_running else "Stopped"
        elements['btn_next'].config(state=tk.NORMAL if is_running else tk.DISABLED)
        elements['btn_toggle'].config(state=tk.NORMAL if is_running else tk.DISABLED)
        if is_running and state == "Playing":
            elements['btn_toggle'].config(text="⏸")
        else:
            elements['btn_toggle'].config(text="▶")
            
    def open_wallpaper_settings(self, app_config):
        current = load_wallch_settings()
        WallpaperSettingsDialog(self, app_config, current)

    def wallch_toggle(self):
        send_wallch_command("toggle"); self.after(250, self.update_wallch_ui)

    def wallch_next(self):
        send_wallch_command("next")

    # --- Tray menu with Profiles + Autostart toggle ---
    def _menu_autostart_checked(self, item):
        return bool(self.app_state.get("autostart", False))

    def _toggle_autostart(self):
        new_val = not bool(self.app_state.get("autostart", False))
        self.app_state["autostart"] = new_val
        _save_state(self.app_state)
        _set_startup_enabled(new_val)
        self.refresh_tray_menu()

    # Add this helper method inside AppManager (just above refresh_tray_menu)
    def _tk_cb(self, fn, *args, **kwargs):
        """Wrap a function so it runs on the Tk main thread when called from pystray."""
        def _cb(icon=None, item=None):
            self.after(0, lambda: fn(*args, **kwargs))
        return _cb

    def refresh_tray_menu(self):
        def clear_profile():
            self.apply_profile(None)

        # Profiles submenu (only apply actions; no “Manage Profiles…” here)
        if self.profiles:
            profile_items = []
            for prof_name in sorted(self.profiles.keys()):
                def make_action(name=prof_name):
                    return self._tk_cb(self.apply_profile, name)
                def is_checked(item, name=prof_name):
                    return self.app_state.get("last_profile") == name
                profile_items.append(pystray.MenuItem(prof_name, make_action(), checked=is_checked))
            profile_items.append(pystray.Menu.SEPARATOR)
            profile_items.append(pystray.MenuItem("Clear profile", self._tk_cb(clear_profile)))
            profiles_menu = pystray.MenuItem("Profiles", pystray.Menu(*profile_items))
        else:
            profiles_menu = pystray.MenuItem("Profiles (none)", lambda icon, item: None, enabled=False)

        # Autostart toggle (thread-safe)
        autostart_item = pystray.MenuItem(
            "Start on Boot",
            self._tk_cb(self._toggle_autostart),
            checked=lambda item: bool(self.app_state.get("autostart", False))
        )

        # Final tray menu (removed “Manage Profiles…” entry to avoid the deadlock)
        menu = (
            pystray.MenuItem('Show', self._tk_cb(self.show_window), default=True),
            pystray.MenuItem('Wallpaper ▶/⏸', self._tk_cb(self.wallch_toggle)),
            pystray.MenuItem('Wallpaper ⏭', self._tk_cb(self.wallch_next)),
            profiles_menu,
            autostart_item,
            pystray.MenuItem('Quit', self._tk_cb(self.quit_window))
        )
        if self.tray_icon:
            self.tray_icon.menu = pystray.Menu(*menu)


    def setup_tray_icon(self):
        try: image = Image.open(APP_DIR / "icon.png")
        except: image = Image.new("RGBA", (64, 64), (30, 30, 30, 255))
        self.tray_icon = pystray.Icon("The Control", image, "The Control", pystray.Menu())  # temp
        self.refresh_tray_menu()
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self):
        # Place before show to avoid flicker at (0, 0)
        self.update_idletasks()
        self.position_window()
        # Fade in quickly to hide any layout jitter
        try:
            self.attributes('-alpha', 0.0)
        except Exception:
            pass
        self.deiconify()
        self.lift()
        self.focus_force()

        def _fade_in():
            try:
                self.attributes('-alpha', 1.0)
            except Exception:
                pass
        self.after(10, _fade_in)

    def hide_window(self):
        self.withdraw()

    def quit_window(self):
        if self.after_id:
            self.after_cancel(self.after_id)
        try:
            CMD_FILE.write_text("quit\n", encoding="utf-8")
        except Exception:
            pass
        time.sleep(0.3)
        self.update_statuses()
        if self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass
        self.destroy()

if __name__ == "__main__":
    try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except: pass

    _kill_other_control_instances()

    app = AppManager()
    app.mainloop()
