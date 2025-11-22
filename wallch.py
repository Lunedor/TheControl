# wallch.py
import argparse
import ctypes
import os
import random
import time
from pathlib import Path
import atexit

SPI_SETDESKWALLPAPER = 20
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp"}

STATUS_PATH = Path(__file__).resolve().with_name("wallch.status")

def ensure_single_instance(key: str):
    ERROR_ALREADY_EXISTS = 183
    handle = ctypes.windll.kernel32.CreateMutexW(None, False, key)
    # Keep the handle alive for process lifetime; don't close it.
    if ctypes.GetLastError() == ERROR_ALREADY_EXISTS:
        print("[wallch] another instance is already running; exiting.")
        raise SystemExit(0)
    return handle

def write_status(state: str):
    try:
        STATUS_PATH.write_text(state + "\n", encoding="utf-8")
    except Exception:
        pass

def set_wallpaper_style(style: str):
    style = style.lower()
    mapping = {
        "fill":     ("0", "10"),
        "fit":      ("0", "6"),
        "stretch":  ("0", "2"),
        "center":   ("0", "0"),
        "tile":     ("1", "0"),
        "span":     ("0", "22"),  # multi-monitor span
    }
    if style not in mapping:
        raise ValueError(f"Unknown style '{style}'. Use one of: {', '.join(mapping)}")
    tile, wp_style = mapping[style]
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, tile)
        winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, wp_style)

def apply_wallpaper(image_path: Path):
    ok = ctypes.windll.user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER, 0, str(image_path),
        SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
    )
    if not ok:
        raise OSError(f"Failed to set wallpaper: {image_path}")

def gather_images(folder: Path, recursive: bool) -> list[Path]:
    if recursive:
        files = [p for p in folder.rglob("*") if p.suffix.lower() in IMAGE_EXTS]
    else:
        files = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS]
    files.sort()
    return files

def read_command(cmd_file: Path) -> str | None:
    """Read and clear a one-line command file if present."""
    try:
        if cmd_file.exists():
            txt = cmd_file.read_text(encoding="utf-8").strip().lower()
            cmd_file.unlink(missing_ok=True)  # consume once
            if txt:
                return txt
    except Exception:
        # ignore transient file errors
        pass
    return None

def _on_exit():
    try: STATUS_PATH.write_text("Stopped\n", encoding="utf-8")
    except Exception: pass

atexit.register(_on_exit)

def main():
    parser = argparse.ArgumentParser(description="Rotate Windows wallpaper from a folder.")
    parser.add_argument("folder", type=Path, help="Folder containing images")
    parser.add_argument("--interval", type=int, default=600, help="Seconds between changes (default: 600)")
    parser.add_argument("--style", default="fill", choices=["fill","fit","stretch","center","tile","span"],
                        help="Wallpaper display style (default: fill)")
    parser.add_argument("--shuffle", action="store_true", help="Shuffle image order")
    parser.add_argument("--recursive", action="store_true", help="Search folder recursively")
    parser.add_argument("--once", action="store_true", help="Set one image then exit")
    args = parser.parse_args()

    if os.name != "nt":
        raise SystemExit("This script only works on Windows.")

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        raise SystemExit(f"Folder not found: {folder}")

    images = gather_images(folder, args.recursive)

    # Auto-fallback: if no images at top-level and user didn't ask recursive, try recursive before quitting
    if not images and not args.recursive:
        fallback = gather_images(folder, True)
        if fallback:
            print("[wallch] No images at the selected folder level. Found images in subfolders — enabling recursive fallback.")
            images = fallback

    if not images:
        raise SystemExit(f"No images found in {folder} (extensions: {', '.join(sorted(IMAGE_EXTS))})")

    if args.shuffle:
        random.shuffle(images)

    try:
        set_wallpaper_style(args.style)
    except Exception as e:
        raise SystemExit(f"Failed to set style: {e}")

    # command file for control (same folder as this script)
    cmd_file = Path(__file__).resolve().with_name("wallch.cmd")

    paused = False
    index = 0
    next_requested = False  # <-- lives across loop iterations
    _ = ensure_single_instance(f"Global\\wallch::{str(folder).lower()}")
    write_status("Playing" if not args.once else "Playing")

    while True:
        # --- Handle any immediate command at the top ---
        cmd = read_command(cmd_file)
        if cmd:
            if cmd in ("pause", "resume", "toggle", "next", "quit"):
                if cmd == "pause":
                    paused = True
                    write_status("Paused")
                    print("[wallch] paused")
                elif cmd == "resume":
                    paused = False
                    write_status("Playing")
                    print("[wallch] resumed")
                elif cmd == "toggle":
                    paused = not paused
                    write_status("Paused" if paused else "Playing")
                    print(f"[wallch] {'paused' if paused else 'resumed'}")
                elif cmd == "next":
                    index += 1
                    next_requested = True
                elif cmd == "quit":
                    write_status("Stopped")
                    print("[wallch] quitting by command")
                    return
            else:
                print(f"[wallch] unknown command: {cmd}")

        # If paused and Next was requested, apply one image now (stay paused)
        if paused and next_requested:
            img = images[index % len(images)]
            if img.exists():
                try:
                    apply_wallpaper(img)
                    write_status("Paused")  # keep UI correct
                    print(f"[wallch] set (paused-next): {img}")
                except Exception as e:
                    print(f"[wallch] {e}")
            else:
                print(f"[wallch] missing image skipped: {img}")
            next_requested = False  # consumed

        # Normal apply when not paused
        if not paused:
            img = images[index % len(images)]
            if not img.exists():
                print(f"[wallch] missing image skipped: {img}")
                index += 1
            else:
                try:
                    apply_wallpaper(img)
                    write_status("Playing")
                    print(f"[wallch] set: {img}")
                except Exception as e:
                    print(f"[wallch] {e}")

                if args.once:
                    write_status("Stopped")
                    return

                # IMPORTANT: advance after a normal apply
                index += 1

        # --- Monotonic wait (no drift) ---
        remaining = max(1, args.interval)
        next_deadline = time.monotonic() + remaining
        step = 0.25  # polling granularity (you can use 0.5 if you like)

        while True:
            time.sleep(step)

            # consume commands responsively
            cmd = read_command(cmd_file)
            if cmd:
                if cmd == "next":
                    index += 1
                    next_requested = True  # carry to next outer cycle
                    break
                elif cmd == "toggle":
                    paused = not paused
                    write_status("Paused" if paused else "Playing")
                    print(f"[wallch] {'paused' if paused else 'resumed'}")
                elif cmd == "pause":
                    paused = True
                    write_status("Paused")
                    print("[wallch] paused")
                elif cmd == "resume":
                    paused = False
                    write_status("Playing")
                    print("[wallch] resumed")
                elif cmd == "quit":
                    write_status("Stopped")
                    print("[wallch] quitting by command")
                    return

            if paused:
                # hold the deadline while paused (and be a bit gentler on wakeups)
                next_deadline = time.monotonic() + remaining
                time.sleep(0.5)
                continue

            if time.monotonic() >= next_deadline:
                break

if __name__ == "__main__":
    main()
