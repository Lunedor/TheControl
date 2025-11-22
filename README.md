# 🖥️ The Control

A lightweight, powerful Windows utility to manage background applications, automate their startup behavior, monitor resource usage, and control a wallpaper-changer daemon — all from a clean, modern UI.

**The Control** combines process management, profiles, tray-icon control, autostart, drag-and-drop application ordering, and a fully featured wallpaper engine into a single elegant tool.

---

## ✨ Features

### ✔ Process Manager

* Start/stop any application or script
* Detects running processes reliably
* Automatically restarts apps if they crash
* Shows CPU & Memory usage in real time
* Supports executables, Python scripts, and shell commands

### ✔ Wallpaper Engine (wallch.py)

A fully custom wallpaper changer with:

* Shuffle, interval, recursive folder scanning
* Pause / Resume / Next wallpaper controls
* Windows wallpaper style support (Fill, Fit, Stretch, Tile, Center, Span)
* Fast communication with Control via command file
* No drift thanks to a monotonic timer
* Mutex-based protection to prevent multiple launches
* Status reporting (Playing / Paused / Stopped)

### ✔ Profiles

Create profiles like:

* **Work** → Start some apps, stop others
* **Chill** → Launch only selected apps
* Non-destructive: apps not listed in the profile remain untouched
* Switch profiles instantly from the tray menu

### ✔ Clean Modern UI

* Dark mode via ttkbootstrap
* Drag-and-drop to reorder apps
* Scrollable list for large setups
* Inline wallpaper controls (toggle, next, settings)
* Smooth window placement near the bottom-right corner
* Fade-in animation

### ✔ Tray Integration

* Show / Hide the main window
* Toggle wallpaper play/pause
* Next wallpaper
* Apply profiles
* Enable/disable autostart
* Quit the app entirely

### ✔ Autostart Support

Uses Windows registry to optionally launch **The Control** at login.

### ✔ Logging

Each app gets its own log file with automatic last-50-lines rotation.

---

## 📦 Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/the-control.git
cd the-control
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

Required packages:

* `psutil`
* `ttkbootstrap`
* `Pillow`
* `pystray`

### 3. Run the app

```bash
python control.py
```

> The first launch will show the UI. After that, it runs in the system tray.

---

## 🗂️ Files Overview

| File                     | Description                                                                   |
| ------------------------ | ----------------------------------------------------------------------------- |
| **control.py**           | Main GUI app. Handles process management, profiles, tray, autostart, logging. |
| **wallch.py**            | Wallpaper-changer daemon controlled by Control.                               |
| **apps.json**            | User-defined apps to manage (auto-generated).                                 |
| **wallch_settings.json** | Configuration for the wallpaper engine.                                       |
| **profiles.json**        | Saved user profiles.                                                          |
| **control.state.json**   | Remembers ON/OFF state and autostart settings.                                |
| **logs/**                | Per-application rotating logs.                                                |

---

## 🧩 Supported App Types

When adding a new app:

| Type              | What You Provide                                         |
| ----------------- | -------------------------------------------------------- |
| **Executable**    | Path to `.exe` and process name                          |
| **Command**       | Shell command (e.g., `rclone mount ...`)                 |
| **Python Script** | Script name (e.g., `wallch.py`) + pythonw auto-launching |
| **Custom**        | Combination of the above                                 |

---

## 🎛️ Wallpaper Settings

Accessible through the ⚙ button under the wallpaper engine entry.

Options include:

* Folder to use
* Interval in seconds
* Style
* Shuffle
* Recursive search
* “Once” mode (set one wallpaper and exit)

---

## 🧠 Advanced Features

### Robust Process Matching

Python scripts are matched by exact cmdline arguments to avoid PID confusion.

### Safe Process Termination

Kills entire process trees to ensure shell-launched apps also shut down cleanly.

### Non-blocking Tray Callbacks

All tray actions are marshalled to the Tk main thread.

### Auto Instance Clean-Up

Previous control.py instances are terminated to prevent duplicates.

---

## 🚀 Roadmap (Optional Ideas)

* Export/import app profiles
* Grouping apps (Work tools, Games, Media, Background services)
* Icon extraction from .exe files
* Linux/macOS variants
* Global hotkeys
* Plugin system for power users
* Notification popups for auto-restart events

---

## 📝 License

You can add GPL, MIT, or a custom license depending on how you want the project shared.

---

## 🙌 Credits

Created by **Lunedor** — built with love, Python, and far too many background apps running simultaneously.
