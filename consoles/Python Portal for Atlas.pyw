"""
Python Portal for Atlas
=======================
A remote support tool that polls GitHub for commands and executes them
on the user's machine.  Stdlib-only so it runs on any Python install.

The portal lives in the folder the user chose when summoning it.
When done, the user can close and delete it.

Commands are pushed to:
  https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/
    manifests/portal-commands.json

Each command has a unique ID.  The portal tracks executed IDs locally
so it never runs the same command twice.

UI: Pulsing portal on the left, chat sidebar on the right that slides
open when Atlas sends a message.  User can reply back to Discord.
Speech can be muted from the chat sidebar.
"""

import ctypes
import json
import os
import platform
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import uuid
import zlib
import struct
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import font as tkfont

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------
PORTAL_VERSION = "2.1.0"
WEBHOOK_URL = (
    "https://discord.com/api/webhooks/1524620703259951104/"
    "fqpIEBXVWsKHy7f1iZ9xoryCpidmjPYIDuITfcwMOjBfMyS2HtJNWpVbfOetapl8vw9O"
)
FIREBASE_URL = "https://mbe-portal-default-rtdb.firebaseio.com"
POLL_INTERVAL = 1.5  # seconds (Firebase — fast for all commands)
CHAT_POLL_INTERVAL = 1  # seconds (Firebase chat — instant feel)
REMINDER_INTERVAL = 25 * 60  # 25 minutes in seconds
CREATE_NO_WINDOW = (
    subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
)

# Session ID — unique per portal instance (for multi-session support)
SESSION_ID = f"portal-{uuid.uuid4().hex[:12]}"

# Colours
_BG = "#0a0a12"
_PANEL = "#12121f"
_PORTAL_IDLE = "#9b59b6"
_PORTAL_ACTIVE = "#00d4ff"
_PORTAL_DONE = "#22c55e"
_PORTAL_ERROR = "#ef4444"
_PORTAL_PAUSED = "#f59e0b"  # amber/yellow
_TEXT = "#ffffff"
_MUTED = "#6b7280"
_CHAT_BG = "#0d0d18"
_CHAT_BUBBLE_ATLAS = "#0a0a12"
_CHAT_BUBBLE_USER = "#2a1a3e"
_CHAT_ENTRY_BG = "#1a0a2e"
_CHAT_BORDER = "#4a2a6e"
_ACCENT = "#9b59b6"

# ------------------------------------------------------------------
# Discord communication — uses session webhook if available, else default
# ------------------------------------------------------------------
_active_webhook_url = WEBHOOK_URL  # overridden by bot-assigned per-session URL

def _set_webhook_url(url):
    global _active_webhook_url
    _active_webhook_url = url or WEBHOOK_URL

def _post_to_discord(content):
    try:
        payload = json.dumps({"content": content[:1900]}).encode("utf-8")
        req = urllib.request.Request(
            _active_webhook_url, data=payload,
            headers={"Content-Type": "application/json",
                     "User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status in (200, 204)
    except Exception:
        return False


def _post_file_to_discord(content, file_path):
    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
        boundary = f"----WebKitFormBoundary{os.urandom(8).hex()}"
        payload_json = json.dumps({"content": content[:1900]})

        # Guess content type from extension
        from mimetypes import guess_type
        content_type = guess_type(file_path)[0] or "application/octet-stream"

        body = b""
        body += f"--{boundary}\r\n".encode()
        body += b'Content-Disposition: form-data; name="payload_json"\r\n'
        body += b"Content-Type: application/json\r\n\r\n"
        body += payload_json.encode() + b"\r\n"

        body += f"--{boundary}\r\n".encode()
        body += (f'Content-Disposition: form-data; name="files[0]"; '
                 f'filename="{Path(file_path).name}"\r\n').encode()
        body += f"Content-Type: {content_type}\r\n\r\n".encode()
        body += file_data + b"\r\n"
        body += f"--{boundary}--\r\n".encode()

        req = urllib.request.Request(
            _active_webhook_url, data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}",
                     "User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status in (200, 204)
    except Exception:
        return _post_to_discord(content + "\n\n[File attachment failed]")


# ------------------------------------------------------------------
# Firebase session management
# ------------------------------------------------------------------
def _firebase_put(path, data):
    """PUT data to a Firebase path."""
    try:
        url = f"{FIREBASE_URL}/{path}.json"
        payload = json.dumps(data).encode("utf-8")
        req = urllib.request.Request(url, data=payload, method="PUT",
                                      headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status in (200, 204)
    except Exception:
        return False


def _firebase_get(path):
    """GET data from a Firebase path. Returns None on error."""
    try:
        url = f"{FIREBASE_URL}/{path}.json"
        req = urllib.request.Request(url,
                                      headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _firebase_delete(path):
    """DELETE a Firebase path."""
    try:
        url = f"{FIREBASE_URL}/{path}.json"
        req = urllib.request.Request(url, method="DELETE")
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status in (200, 204)
    except Exception:
        return False


def _register_session(user, host, folder):
    """Register this portal session in Firebase so the bot can create a channel."""
    data = {
        "user": user,
        "host": host,
        "folder": folder,
        "status": "open",
        "opened_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "version": PORTAL_VERSION,
    }
    return _firebase_put(f"sessions/{SESSION_ID}", data)


def _wait_for_webhook(timeout=30):
    """Poll Firebase for the bot to assign a session-specific webhook URL.
    Falls back to default webhook if bot doesn't respond in time."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        data = _firebase_get(f"sessions/{SESSION_ID}")
        if data and data.get("webhook_url"):
            _set_webhook_url(data["webhook_url"])
            return True
        time.sleep(1)
    return False


def _mark_session_closed():
    """Mark this session as closed so the bot can delete the temp channel."""
    _firebase_put(f"sessions/{SESSION_ID}/status", "closed")


# ------------------------------------------------------------------
# Command execution
# ------------------------------------------------------------------
def _scan_directory(path):
    """Scan the current directory and show files + subfolders."""
    lines = []
    p = Path(path)
    if not p.exists():
        return f"Path does not exist: {path}"

    lines.append(f"Scanned: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append(f"Contents of: {p}")
    if p.parent.exists():
        lines.append(f"Parent: {p.parent}")
    lines.append("")

    # ---- Show subfolders (children) ----
    subdirs = sorted([d for d in p.iterdir() if d.is_dir()])
    if subdirs:
        lines.append("Subfolders:")
        for sub in subdirs[:20]:  # limit to avoid wall of text
            lines.append(f"  {sub.name}/")
        if len(subdirs) > 20:
            lines.append(f"  ... and {len(subdirs) - 20} more")
        lines.append("")

    # ---- Show files in this folder ----
    lines.append("Files in this folder:")
    files = sorted([f for f in p.iterdir() if f.is_file()])
    if not files:
        lines.append("  (no files)")
    for f in files:
        try:
            size = f.stat().st_size
            if size > 1024 * 1024:
                size_str = f"({size / 1024 / 1024:.1f} MB)"
            elif size > 1024:
                size_str = f"({size / 1024:.0f} KB)"
            else:
                size_str = f"({size} B)"
        except Exception:
            size_str = ""
        lines.append(f"  {f.name} {size_str}")

    # ---- Dive one level deeper into subfolders ----
    lines.append("")
    lines.append("One level deeper:")
    for sub in subdirs[:10]:
        subfiles = sorted([f.name for f in sub.iterdir() if f.is_file()][:5])
        if subfiles:
            lines.append(f"  {sub.name}/ → {', '.join(subfiles)}")
        else:
            lines.append(f"  {sub.name}/ → (empty)")
    if len(subdirs) > 10:
        lines.append(f"  ... and {len(subdirs) - 10} more subfolders")

    # ---- Nearby folders (compact) — so the bot can cache them for autocomplete ----
    if p.parent.exists():
        siblings = sorted([d for d in p.parent.iterdir() if d.is_dir()])
        if siblings:
            lines.append("")
            lines.append("Quick options (nearby folders):")
            for sib in siblings[:15]:
                lines.append(f"  {sib.resolve()}/")
            if len(siblings) > 15:
                lines.append(f"  ... and {len(siblings) - 15} more")

    return "\n".join(lines)


def _check_packages():
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "list", "--format=columns"],
            capture_output=True, text=True, timeout=30,
            creationflags=CREATE_NO_WINDOW)
        return proc.stdout or proc.stderr
    except Exception as e:
        return f"Error: {e}"


def _read_file(path, max_bytes=50000):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            data = f.read(max_bytes)
        if len(data) == max_bytes:
            data += "\n... [truncated]"
        return data
    except Exception as e:
        return f"Error reading file: {e}"


# ------------------------------------------------------------------
# Undo / backup helpers
# ------------------------------------------------------------------
def _backup_dir():
    """Return the hidden backup directory inside the portal folder."""
    return Path(__file__).parent / ".portal_backups"


def _backup_log_file():
    return _backup_dir() / "undo_log.json"


def _backup_file(path, operation, extra=None):
    """Backup a file before a destructive operation. Returns backup info or None."""
    p = Path(path)
    if not p.exists() or not p.is_file():
        return None
    try:
        backup_dir = _backup_dir()
        backup_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_name = f"{timestamp}_{p.name}"
        backup_path = backup_dir / backup_name
        shutil.copy2(p, backup_path)
        entry = {
            "original_path": str(p.resolve()),
            "backup_path": str(backup_path),
            "operation": operation,
            "timestamp": datetime.now().isoformat(),
        }
        if extra:
            entry.update(extra)
        logs = []
        log_file = _backup_log_file()
        if log_file.exists():
            try:
                logs = json.load(open(log_file, "r", encoding="utf-8"))
            except Exception:
                pass
        logs.append(entry)
        json.dump(logs, open(log_file, "w", encoding="utf-8"), indent=2)
        return entry
    except Exception:
        return None


def _undo_last():
    """Undo the last destructive file operation. Returns (ok, message)."""
    log_file = _backup_log_file()
    if not log_file.exists():
        return False, "No undo history available."
    try:
        logs = json.load(open(log_file, "r", encoding="utf-8"))
    except Exception:
        return False, "Could not read undo history."
    if not logs:
        return False, "No undo history available."

    entry = logs.pop()
    operation = entry.get("operation", "")
    original_path = entry.get("original_path", "")
    backup_path = entry.get("backup_path", "")

    try:
        if operation in ("delete", "replace", "add"):
            if not backup_path or not Path(backup_path).exists():
                return False, f"Backup missing for {operation} on {original_path}"
            Path(original_path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(backup_path, original_path)
            # Remove the backup entry from log
            json.dump(logs, open(log_file, "w", encoding="utf-8"), indent=2)
            return True, f"Undo {operation}: restored {original_path}"
        elif operation == "rename":
            old_path = entry.get("old_path", "")
            new_path = entry.get("new_path", "")
            if not old_path or not new_path:
                return False, "Rename undo missing path info"
            if not Path(backup_path).exists():
                return False, f"Backup missing for rename: {old_path}"
            # Delete the new file and restore the old one
            if Path(new_path).exists():
                Path(new_path).unlink()
            Path(old_path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(backup_path, old_path)
            json.dump(logs, open(log_file, "w", encoding="utf-8"), indent=2)
            return True, f"Undo rename: restored {old_path}"
        else:
            return False, f"Unknown operation: {operation}"
    except Exception as e:
        return False, f"Undo failed: {e}"


def _download_file(url, dest_path):
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
        Path(dest_path).parent.mkdir(parents=True, exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(data)
        return True, f"Downloaded {len(data)} bytes to {dest_path}"
    except Exception as e:
        return False, f"Download failed: {e}"


def _write_png_rgb(output_path, rgb_bytes, width, height):
    """Write raw RGB24 bytes to a PNG file using only stdlib zlib."""
    raw = bytearray()
    stride = width * 3
    for y in range(height):
        raw.append(0)  # no filter
        raw.extend(rgb_bytes[y * stride:(y + 1) * stride])
    compressed = zlib.compress(bytes(raw))

    def _chunk(tag, data):
        chunk = struct.pack(">I", len(data)) + tag + data
        crc = zlib.crc32(chunk[4:]) & 0xffffffff
        return chunk + struct.pack(">I", crc)

    with open(output_path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)))
        f.write(_chunk(b"IDAT", compressed))
        f.write(_chunk(b"IEND", b""))


def _capture_hbitmap_to_png(hdc_mem, hbitmap, width, height, output_path):
    """Extract a Windows HBITMAP (24-bit BGR top-down) and save as PNG."""
    from ctypes import wintypes
    gdi32 = ctypes.windll.gdi32

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD),
            ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 1)]

    bih = BITMAPINFOHEADER()
    bih.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bih.biWidth = width
    bih.biHeight = -height  # negative = top-down
    bih.biPlanes = 1
    bih.biBitCount = 24
    bih.biCompression = 0
    bih.biSizeImage = 0

    bi = BITMAPINFO()
    bi.bmiHeader = bih

    buffer_size = width * height * 3
    buffer = (ctypes.c_ubyte * buffer_size)()
    gdi32.GetDIBits(hdc_mem, hbitmap, 0, height, buffer, ctypes.byref(bi), 0)

    # Windows gives BGR; PNG wants RGB
    data = bytearray(buffer)
    for i in range(0, len(data), 3):
        data[i], data[i + 2] = data[i + 2], data[i]

    _write_png_rgb(output_path, bytes(data), width, height)


def _take_screenshot(output_path, hwnd=0):
    """Capture the primary screen (hwnd=0) or a specific window and save to output_path (PNG)."""
    system = platform.system()
    if system == "Windows":
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
        try:
            if hwnd:
                # Window-specific capture
                hdc_window = user32.GetDC(hwnd)
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w <= 0 or h <= 0:
                    return False, "Invalid window size"
            else:
                hdc_window = user32.GetDC(0)
                w = user32.GetSystemMetrics(0)
                h = user32.GetSystemMetrics(1)
            hdc_mem = gdi32.CreateCompatibleDC(hdc_window)
            hbitmap = gdi32.CreateCompatibleBitmap(hdc_window, w, h)
            gdi32.SelectObject(hdc_mem, hbitmap)
            SRCCOPY = 0x00CC0020
            if hwnd:
                # Use PrintWindow to capture window content including non-client area
                PW_RENDERFULLCONTENT = 0x00000002
                if not user32.PrintWindow(hwnd, hdc_mem, PW_RENDERFULLCONTENT):
                    user32.PrintWindow(hwnd, hdc_mem, 0)
            else:
                gdi32.BitBlt(hdc_mem, 0, 0, w, h, hdc_window, 0, 0, SRCCOPY)
            _capture_hbitmap_to_png(hdc_mem, hbitmap, w, h, output_path)
            gdi32.DeleteObject(hbitmap)
            gdi32.DeleteDC(hdc_mem)
            user32.ReleaseDC(hwnd, hdc_window)
            return True, output_path
        except Exception as e:
            return False, f"Screenshot error: {e}"
    elif system == "Darwin":
        if hwnd:
            return False, "Window-specific screenshot not supported on Mac"
        try:
            proc = subprocess.run(
                ["/usr/sbin/screencapture", "-x", output_path],
                capture_output=True, text=True, timeout=30)
            if proc.returncode != 0:
                return False, f"Screenshot failed: {proc.stderr.strip()}"
            return True, output_path
        except Exception as e:
            return False, f"Screenshot error: {e}"
    else:
        return False, f"Screenshot unsupported on {system}"


def _list_visible_windows():
    """Return a list of (hwnd, title) for visible top-level windows."""
    if platform.system() != "Windows":
        return []
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    windows = []
    EnumWindowsProc = ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _extra):
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                if rect.right - rect.left > 0 and rect.bottom - rect.top > 0:
                    windows.append((hwnd, title))
        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return windows


def _capture_all_windows(tag):
    """Capture each visible window and post it to Discord. Returns (ok, message)."""
    import tempfile
    if platform.system() != "Windows":
        return False, "Window-specific capture only supported on Windows"

    windows = _list_visible_windows()
    if not windows:
        return False, "No visible windows found"

    posted = 0
    failed = []
    for hwnd, title in windows:
        try:
            fd, path = tempfile.mkstemp(suffix=".png", prefix="win_")
            os.close(fd)
            ok, msg = _take_screenshot(path, hwnd=hwnd)
            if ok and Path(path).exists():
                safe_title = title.replace("`", "'")[:60] or "Untitled"
                _post_file_to_discord(f"{tag} Window: {safe_title}", path)
                posted += 1
                try:
                    Path(path).unlink()
                except Exception:
                    pass
            else:
                failed.append(f"{title[:40]}: {msg}")
                try:
                    Path(path).unlink()
                except Exception:
                    pass
        except Exception as e:
            failed.append(f"{title[:40]}: {e}")
    summary = f"Captured {posted} window(s)"
    if failed:
        summary += f", {len(failed)} failed"
    _post_to_discord(f"{tag} {summary}")
    return True, summary


def _pip_install(packages):
    if isinstance(packages, str):
        packages = [packages]
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--upgrade"] + packages,
            capture_output=True, text=True, timeout=300,
            creationflags=CREATE_NO_WINDOW)
        output = proc.stdout + proc.stderr
        return (proc.returncode == 0, output[-2000:])
    except Exception as e:
        return False, f"pip install error: {e}"


def _speak_text(text):
    try:
        if platform.system() == "Windows":
            escaped = text.replace("'", "''")
            ps_cmd = (
                "Add-Type -AssemblyName System.Speech; "
                "(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{}')"
            ).format(escaped)
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_cmd],
                capture_output=True, timeout=30,
                creationflags=CREATE_NO_WINDOW)
        elif platform.system() == "Darwin":
            subprocess.run(["say", text], capture_output=True, timeout=30)
    except Exception:
        pass


def _write_temp(text):
    import tempfile
    fd, path = tempfile.mkstemp(suffix=".txt", prefix="portal_")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def _write_temp_binary(data, suffix=".bin"):
    import tempfile
    fd, path = tempfile.mkstemp(suffix=suffix, prefix="portal_")
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path


def _zip_directory(path, output_path=None):
    """Zip the contents of a directory, return the zip file path."""
    import tempfile, zipfile
    p = Path(path)
    if not p.exists() or not p.is_dir():
        return None
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix=".zip", prefix="portal_")
        os.close(fd)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in p.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(p))
    return output_path


def _simple_list_files(path):
    """Return a simple flat list of files in the current directory."""
    p = Path(path)
    if not p.exists() or not p.is_dir():
        return f"Path does not exist: {path}"
    lines = [f"Files in: {p}", ""]
    if p.parent.exists():
        lines.append(f"Parent: {p.parent}")
        lines.append("")
    files = sorted([f for f in p.iterdir() if f.is_file()])
    if not files:
        lines.append("  (no files)")
    for f in files:
        size = f.stat().st_size
        size_str = f"{size} bytes" if size < 1024 else f"{size // 1024} KB"
        lines.append(f"  {f.name}  ({size_str})")
    return "\n".join(lines)


# ------------------------------------------------------------------
# Portal UI — pulsing circle + chat sidebar
# ------------------------------------------------------------------
class PortalWindow:
    def __init__(self, root, portal_folder=None, color_override=None):
        self.root = root
        self.portal_folder = portal_folder
        self.portal_dir = Path(__file__).parent
        self.executed_file = self.portal_dir / ".portal_executed.json"
        self.executed_ids = self._load_executed()
        self.pulse_phase = 0.0
        self.idle_color = color_override or _PORTAL_IDLE
        self.portal_color = self.idle_color
        self.paused = False
        self.status_text = "Portal idle — listening for Atlas' commands..."
        self.muted = False
        self.chat_visible = False
        self.user_closed_once = False  # True after first close — second close is permanent

        # Window — wider to accommodate chat sidebar
        root.title(f"Python Portal for Atlas v{PORTAL_VERSION}")
        root.geometry("400x400")
        root.resizable(False, False)
        root.configure(bg=_BG)
        root.protocol("WM_DELETE_WINDOW", self._close)

        # Bring to front on open — stays topmost briefly to survive
        # the "Portal Opened" messagebox from the launcher
        root.attributes("-topmost", True)
        root.after(500, lambda: root.attributes("-topmost", True))
        root.after(2000, lambda: root.attributes("-topmost", False))

        # ---- Main portal area (left) ----
        self.main_frame = tk.Frame(root, bg=_BG, width=400, height=400)
        self.main_frame.pack(side="left", fill="both", expand=True)
        self.main_frame.pack_propagate(False)

        # Canvas for the pulsing portal
        self.canvas = tk.Canvas(self.main_frame, width=400, height=280,
                                 bg=_BG, highlightthickness=0)
        self.canvas.pack()

        # Status text
        self.status_var = tk.StringVar(value=self.status_text)
        tk.Label(self.main_frame, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg=_BG, fg=_TEXT,
                 wraplength=360).pack(pady=(0, 4))

        # Bottom bar: chat bubble icon + footer note
        bottom_bar = tk.Frame(self.main_frame, bg=_BG)
        bottom_bar.pack(side="bottom", fill="x", padx=12, pady=(0, 8))

        # Chat bubble icon (right side of bottom bar — same side chat opens)
        self.chat_btn = tk.Canvas(bottom_bar, width=28, height=28,
                                   bg=_BG, highlightthickness=0, cursor="hand2")
        self.chat_btn.pack(side="right")
        self._draw_chat_bubble(False)
        self.chat_btn.bind("<Button-1>", lambda e: self._toggle_chat())

        # Footer note
        tk.Label(bottom_bar,
                 text="Close to delete this portal\n(you can always open another).",
                 font=("Segoe UI", 7), bg=_BG, fg=_MUTED,
                 justify="left").pack(side="left")

        # ---- Chat sidebar (right, hidden initially) ----
        self.chat_frame = tk.Frame(root, bg=_CHAT_BG, width=340, height=400)
        # Not packed yet — shown when _toggle_chat or message arrives

        # Chat header
        chat_header = tk.Frame(self.chat_frame, bg=_PANEL, height=34)
        chat_header.pack(side="top", fill="x")
        chat_header.pack_propagate(False)

        tk.Label(chat_header, text="Atlas",
                 font=("Segoe UI", 11, "bold"), bg=_PANEL, fg=_TEXT,
                 ).pack(side="left", padx=(10, 0), pady=5)

        # Mute button
        self.mute_btn = tk.Button(chat_header, text="\U0001F50A",
                                   command=self._toggle_mute,
                                   bg=_PANEL, fg=_TEXT, relief="flat",
                                   bd=0, cursor="hand2", font=("Segoe UI", 12))
        self.mute_btn.pack(side="right", padx=(0, 4), pady=3)

        # Close sidebar button
        tk.Button(chat_header, text="\u2715",
                  command=self._hide_chat,
                  bg=_PANEL, fg=_MUTED, relief="flat",
                  bd=0, cursor="hand2", font=("Segoe UI", 10)).pack(
                      side="right", padx=(0, 2), pady=3)

        # Chat messages area — scrollable canvas for bubble rendering
        chat_scroll_frame = tk.Frame(self.chat_frame, bg=_CHAT_BG)
        chat_scroll_frame.pack(side="top", fill="both", expand=True)

        self.chat_canvas = tk.Canvas(chat_scroll_frame, bg=_CHAT_BG,
                                      highlightthickness=0)
        self.chat_scrollbar = tk.Scrollbar(chat_scroll_frame,
                                            command=self.chat_canvas.yview,
                                            bg=_CHAT_BG,
                                            troughcolor=_CHAT_BG,
                                            borderwidth=0,
                                            activebackground=_PANEL)
        self.chat_scrollbar.pack(side="right", fill="y")
        self.chat_canvas.pack(side="left", fill="both", expand=True)
        self.chat_canvas.configure(yscrollcommand=self.chat_scrollbar.set)

        # Inner frame that holds the bubble widgets
        self.chat_inner = tk.Frame(self.chat_canvas, bg=_CHAT_BG)
        self.chat_inner_window = self.chat_canvas.create_window(
            (0, 0), window=self.chat_inner, anchor="n")
        self.chat_inner.bind("<Configure>", self._on_chat_inner_configure)
        self.chat_canvas.bind("<Configure>", self._on_chat_canvas_configure)

        # Mouse wheel / touchpad scrolling
        self.chat_canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.chat_canvas.bind("<Button-4>", self._on_mouse_wheel)
        self.chat_canvas.bind("<Button-5>", self._on_mouse_wheel)
        # Bind to inner frame too so scrolling works over bubbles
        self.chat_inner.bind("<MouseWheel>", self._on_mouse_wheel)
        self.chat_inner.bind("<Button-4>", self._on_mouse_wheel)
        self.chat_inner.bind("<Button-5>", self._on_mouse_wheel)

        # Track bubbles for layout
        self._chat_bubbles = []
        self._bubble_count = 0
        self._welcome_shown = False

        # Welcome screen with stars (drawn on canvas, behind inner frame)
        self._welcome_items = []
        self._draw_welcome()

        # ---- Floating input box (detached from corners) ----
        self.input_container = tk.Frame(self.chat_frame, bg=_CHAT_BG)
        self.input_container.pack(side="bottom", fill="x", padx=12, pady=(4, 10))

        # Input frame with rounded look (using flat bg + padding)
        self.input_wrapper = tk.Frame(self.input_container, bg=_CHAT_ENTRY_BG,
                                       highlightthickness=1,
                                       highlightbackground=_CHAT_BORDER,
                                       highlightcolor=_CHAT_BORDER)
        self.input_wrapper.pack(fill="x", padx=0, pady=0)

        # Auto-growing text widget for input
        self.chat_input = tk.Text(self.input_wrapper, wrap="word",
                                   bg=_CHAT_ENTRY_BG, fg=_TEXT,
                                   insertbackground=_TEXT,
                                   font=("Segoe UI", 10),
                                   relief="flat", bd=0,
                                   padx=10, pady=8,
                                   height=1,
                                   highlightthickness=0)
        self.chat_input.pack(side="left", fill="x", expand=True)
        self.chat_input.bind("<KeyRelease>", self._on_input_change)
        self.chat_input.bind("<Return>", lambda e: self._send_user_message())
        self.chat_input.bind("<Shift-Return>", lambda e: None)

        # Input placeholder
        self._input_placeholder = "Type here to message Atlas..."
        self._input_has_placeholder = True
        self.chat_input.insert("1.0", self._input_placeholder)
        self.chat_input.configure(fg=_MUTED)
        self.chat_input.bind("<FocusIn>", self._on_input_focus_in)
        self.chat_input.bind("<FocusOut>", self._on_input_focus_out)

        # Send button
        send_btn = tk.Button(self.input_wrapper, text="\u279C",
                             command=self._send_user_message,
                             bg=_CHAT_ENTRY_BG, fg=_ACCENT, relief="flat",
                             bd=0, cursor="hand2",
                             font=("Segoe UI", 14))
        send_btn.pack(side="right", padx=(4, 8), pady=6)

        # Hint text below input
        tk.Label(self.input_container,
                 text="Enter to send \u00b7 Shift+Enter for new line",
                 font=("Segoe UI", 7), bg=_CHAT_BG, fg=_MUTED,
                 ).pack(anchor="w", padx=4, pady=(2, 0))

        # Adjust window width when chat is shown
        self._chat_open_width = 740
        self._chat_closed_width = 400

        # Track when portal was opened
        self._opened_at = time.time()
        self._first_poll = True

        # Clear old Firebase chat + commands on startup so we don't replay
        # messages/commands from a previous session
        try:
            _firebase_clear = json.dumps({}).encode()
            for path in ("/chat.json", "/commands.json"):
                req = urllib.request.Request(
                    f"{FIREBASE_URL}{path}",
                    data=_firebase_clear,
                    method="PUT",
                    headers={"Content-Type": "application/json"})
                urllib.request.urlopen(req, timeout=5)
        except Exception:
            pass

        # Start animation + polling + reminder
        self._animate()
        threading.Thread(target=self._poll_loop, daemon=True).start()
        threading.Thread(target=self._reminder_loop, daemon=True).start()
        threading.Thread(target=self._chat_poll_loop, daemon=True).start()

        # Ensure input doesn't grab focus on startup — keep placeholder visible
        self.root.focus_set()
        self.chat_input.tk_focusNext()

    # ---- Chat bubble icon ----
    def _draw_chat_bubble(self, has_unread):
        """Draw a chat bubble icon on the canvas button."""
        self.chat_btn.delete("all")
        color = _ACCENT if has_unread else _MUTED
        # Bubble outline
        self.chat_btn.create_oval(4, 4, 24, 24, outline=color, width=2)
        # Small speech dot
        self.chat_btn.create_oval(10, 10, 18, 18, fill=color, outline="")

    def _toggle_chat(self):
        if self.chat_visible:
            self._hide_chat()
        else:
            self._show_chat()

    def _show_chat(self):
        if not self.chat_visible:
            self.chat_frame.pack(side="right", fill="both")
            self.root.geometry(f"{self._chat_open_width}x400")
            self.chat_visible = True
            self._draw_chat_bubble(False)
            # Keep focus on root so input placeholder stays visible
            self.root.focus_set()

    def _hide_chat(self):
        if self.chat_visible:
            self.chat_frame.pack_forget()
            self.root.geometry(f"{self._chat_closed_width}x400")
            self.chat_visible = False
            self._draw_chat_bubble(False)

    def _toggle_mute(self):
        self.muted = not self.muted
        self.mute_btn.configure(
            text="\U0001F507" if self.muted else "\U0001F50A")

    # ---- Chat canvas scrolling ----
    def _on_chat_inner_configure(self, event):
        """Update scroll region when inner frame changes size."""
        self.chat_canvas.configure(scrollregion=self.chat_canvas.bbox("all"))

    def _on_chat_canvas_configure(self, event):
        """Adjust inner frame width to match canvas width + redraw welcome."""
        self.chat_canvas.itemconfig(self.chat_inner_window, width=event.width)
        # Redraw welcome screen once canvas has real dimensions
        if self._welcome_shown and self._bubble_count == 0 and event.width > 1:
            for item in self._welcome_items:
                try:
                    self.chat_canvas.delete(item)
                except Exception:
                    pass
            self._welcome_items = []
            self._draw_welcome(event.width, event.height)

    def _on_mouse_wheel(self, event):
        """Handle mouse wheel / touchpad scrolling over chat area."""
        if event.num == 5 or event.delta < 0:
            self.chat_canvas.yview_scroll(3, "units")
        elif event.num == 4 or event.delta > 0:
            self.chat_canvas.yview_scroll(-3, "units")

    # ---- Welcome screen (stars) ----
    def _draw_welcome(self, canvas_w=None, canvas_h=None):
        """Draw a 'Welcome to Space' placeholder with stars on the chat canvas."""
        if self._bubble_count > 0:
            return
        self._welcome_shown = True
        # Use provided dimensions (from configure event) or fall back to winfo
        w = canvas_w or self.chat_canvas.winfo_width() or 300
        h = canvas_h or self.chat_canvas.winfo_height() or 250

        # Stars at fixed positions
        import random
        random.seed(42)  # consistent star positions
        stars = []
        for _ in range(15):
            sx = random.randint(20, max(40, w - 20))
            sy = random.randint(20, max(40, h - 60))
            size = random.choice([1, 1, 1, 2, 2, 3])  # mostly small
            brightness = random.choice(["#444466", "#555577", "#666688", "#8888aa"])
            stars.append((sx, sy, size, brightness))

        for sx, sy, size, brightness in stars:
            item = self.chat_canvas.create_oval(
                sx - size, sy - size, sx + size, sy + size,
                fill=brightness, outline="", tags="welcome")
            self._welcome_items.append(item)

        # "Welcome to Space" text — manually shifted left for visual centering
        cx = (w // 2) - 30
        cy = h // 2 - 20
        item = self.chat_canvas.create_text(
            cx, cy, text="Welcome to Space",
            fill=_ACCENT, font=("Segoe UI", 14, "bold"),
            anchor="center", tags="welcome")
        self._welcome_items.append(item)

        item = self.chat_canvas.create_text(
            cx, cy + 24, text="Atlas will appear here when you're needed.",
            fill=_MUTED, font=("Segoe UI", 8),
            anchor="center", tags="welcome")
        self._welcome_items.append(item)

    def _remove_welcome(self):
        """Remove the welcome screen when first bubble arrives."""
        if not self._welcome_shown:
            return
        for item in self._welcome_items:
            try:
                self.chat_canvas.delete(item)
            except Exception:
                pass
        self.chat_canvas.delete("welcome")
        self._welcome_items = []
        self._welcome_shown = False

    # ---- Input placeholder ----
    def _on_input_focus_in(self, event):
        if self._input_has_placeholder:
            self.chat_input.delete("1.0", "end")
            self.chat_input.configure(fg=_TEXT)
            self._input_has_placeholder = False

    def _on_input_focus_out(self, event):
        content = self.chat_input.get("1.0", "end-1c").strip()
        if not content:
            self.chat_input.insert("1.0", self._input_placeholder)
            self.chat_input.configure(fg=_MUTED)
            self._input_has_placeholder = True

    # ---- Chat bubbles ----
    def _add_chat_message(self, sender, text, is_atlas=True):
        """Add a chat bubble to the message area."""
        # Remove welcome screen on first message
        if self._bubble_count == 0:
            self._remove_welcome()
        bubble_bg = _CHAT_BUBBLE_ATLAS if is_atlas else _CHAT_BUBBLE_USER
        name_color = _ACCENT if is_atlas else "#c8a8e8"
        name_label = "Atlas" if is_atlas else "You"

        # Container for this bubble (full width row)
        row = tk.Frame(self.chat_inner, bg=_CHAT_BG)
        row.pack(fill="x", padx=8, pady=(6, 2))

        if is_atlas:
            # Atlas messages: left-aligned, black with purple border
            bubble = tk.Frame(row, bg=bubble_bg, padx=12, pady=8,
                              highlightthickness=1,
                              highlightbackground=_CHAT_BORDER,
                              highlightcolor=_CHAT_BORDER)
            bubble.pack(side="left", anchor="w")
        else:
            # User messages: right-aligned, dark purple
            bubble = tk.Frame(row, bg=bubble_bg, padx=12, pady=8)
            bubble.pack(side="right", anchor="e")

        # Name label
        tk.Label(bubble, text=name_label, font=("Segoe UI", 8, "bold"),
                 bg=bubble_bg, fg=name_color).pack(anchor="w")

        # Message text
        msg_label = tk.Label(bubble, text=text, font=("Segoe UI", 10),
                             bg=bubble_bg, fg=_TEXT, wraplength=240,
                             justify="left")
        msg_label.pack(anchor="w")

        # Timestamp
        ts = datetime.now().strftime("%H:%M")
        tk.Label(bubble, text=ts, font=("Segoe UI", 7),
                 bg=bubble_bg, fg=_MUTED).pack(anchor="w", pady=(2, 0))

        self._bubble_count += 1
        self._chat_bubbles.append(bubble)

        # Auto-scroll to bottom
        self.chat_canvas.update_idletasks()
        self.chat_canvas.yview_moveto(1.0)

    def _add_system_message(self, text):
        """Add a centered system message (no bubble)."""
        row = tk.Frame(self.chat_inner, bg=_CHAT_BG)
        row.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(row, text=text, font=("Segoe UI", 8, "italic"),
                 bg=_CHAT_BG, fg=_MUTED, wraplength=240,
                 justify="center").pack()
        self.chat_canvas.update_idletasks()
        self.chat_canvas.yview_moveto(1.0)

    # ---- Input auto-grow ----
    def _on_input_change(self, event=None):
        """Auto-grow the input text widget as the user types.
        Counts both explicit newlines AND visual wrap lines."""
        if self._input_has_placeholder:
            return  # Don't resize for placeholder
        content = self.chat_input.get("1.0", "end-1c")
        if not content:
            self.chat_input.configure(height=1)
            return
        # Count display lines: each explicit newline + wrapped lines
        total_lines = 0
        for line in content.split("\n"):
            if len(line) == 0:
                total_lines += 1
            else:
                # Approximate wrap: ~28 chars per line at width 200px, font size 10
                wrapped = max(1, -(-len(line) // 28))  # ceiling division
                total_lines += wrapped
        # Cap at 5 lines visible
        new_height = min(max(1, total_lines), 5)
        if self.chat_input.cget("height") != new_height:
            self.chat_input.configure(height=new_height)

    def _send_user_message(self):
        """Send user's chat message to Discord."""
        if self._input_has_placeholder:
            return  # Don't send placeholder text
        msg = self.chat_input.get("1.0", "end-1c").strip()
        if not msg:
            return
        self.chat_input.delete("1.0", "end")
        self.chat_input.configure(height=1)
        self._add_chat_message("You", msg, is_atlas=False)

        # Send to Discord
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        threading.Thread(
            target=lambda: _post_to_discord(
                f"[Portal Chat @ {user}@{host} | {SESSION_ID}] {msg}"),
            daemon=True).start()

    def _receive_atlas_message(self, text, speak=False):
        """Show a message from Atlas in the chat + optionally speak it."""
        # Don't override the paused color
        if not self.paused:
            self._set_color(_PORTAL_ACTIVE, "Message from Atlas...")
        self._add_chat_message("Atlas", text, is_atlas=True)
        # Auto-open chat if closed
        if not self.chat_visible:
            self._show_chat()
        else:
            self._draw_chat_bubble(False)
        # Speak unless muted
        if speak and not self.muted:
            threading.Thread(target=lambda: _speak_text(text), daemon=True).start()
        # Return to idle after 2 seconds (but stay paused-amber if paused)
        if not self.paused:
            self.root.after(2000, lambda: self._set_color(
                self.idle_color, "Portal idle — listening for Atlas' commands..."))

    # ---- Executed IDs persistence ----
    def _load_executed(self):
        try:
            if self.executed_file.exists():
                with open(self.executed_file, "r") as f:
                    return set(json.load(f))
        except Exception:
            pass
        return set()

    def _save_executed(self):
        try:
            with open(self.executed_file, "w") as f:
                json.dump(list(self.executed_ids), f)
        except Exception:
            pass

    def _resolve_path(self, raw_path):
        """Resolve a path relative to the portal folder."""
        if os.path.isabs(raw_path):
            return raw_path
        base = Path(self.portal_folder) if self.portal_folder else Path.cwd()
        return str((base / raw_path).resolve())

    # ---- Animation ----
    def _hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def _lerp_color(self, c1, c2, t):
        r1, g1, b1 = self._hex_to_rgb(c1)
        r2, g2, b2 = self._hex_to_rgb(c2)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _animate(self):
        self.pulse_phase += 0.05
        if self.pulse_phase > 6.283:
            self.pulse_phase = 0.0

        self.canvas.delete("portal")
        cx, cy = 200, 160

        # Force amber color + special styling while paused
        import math
        if self.paused:
            base_r = 55
            max_glow_r = base_r + 85
            glow_layers = 38
            core_layers = 8
            portal_color = _PORTAL_PAUSED
            pulse_t = 0.5 + 0.5 * math.sin(self.pulse_phase)
            r = int(base_r * (1 + 0.08 * pulse_t))
        else:
            base_r = 80
            max_glow_r = base_r + 50
            glow_layers = 30
            core_layers = 12
            portal_color = self.portal_color
            pulse_t = 0.5 + 0.5 * math.sin(self.pulse_phase)
            r = int(base_r * (1 + 0.12 * pulse_t))

        # ---- Outer glow halo (many concentric circles fading outward) ----
        for i in range(glow_layers, 0, -1):
            layer_r = r + int((max_glow_r - r) * (i / glow_layers))
            alpha = (1 - (i / glow_layers)) ** 2 * 0.4
            color = self._lerp_color(_BG, portal_color, alpha)
            self.canvas.create_oval(
                cx - layer_r, cy - layer_r, cx + layer_r, cy + layer_r,
                fill=color, outline="", tags="portal")

        # ---- Expanding pulse rings ----
        for i in range(3):
            wave = (self.pulse_phase + i * 2.0) % 6.283
            progress = wave / 6.283
            ring_r = r + int(40 * progress)
            ring_alpha = (1 - progress) * 0.5
            ring_color = self._lerp_color(_BG, portal_color, ring_alpha)
            self.canvas.create_oval(
                cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r,
                outline=ring_color, width=1, tags="portal")

        # ---- Core gradient (bright center fading to color) ----
        for i in range(core_layers, 0, -1):
            core_r = int(r * (i / core_layers))
            blend = (i / core_layers) ** 1.5
            color = self._lerp_color(portal_color, "#ffffff", 1 - blend)
            self.canvas.create_oval(
                cx - core_r, cy - core_r, cx + core_r, cy + core_r,
                fill=color, outline="", tags="portal")

        # ---- Bright white center ----
        center_r = int(r * 0.15 * (0.8 + 0.2 * pulse_t))
        self.canvas.create_oval(
            cx - center_r, cy - center_r, cx + center_r, cy + center_r,
            fill="#ffffff", outline="", tags="portal")

        self.root.after(40, self._animate)

    def _set_color(self, color, status):
        self.portal_color = color
        self.status_var.set(status)

    # ---- Polling ----
    def _poll_loop(self):
        while True:
            try:
                self._poll_once()
            except Exception as e:
                if not self.paused:
                    self.root.after(0, lambda: self._set_color(
                        _PORTAL_ERROR, f"Poll error: {e}"))
            # When paused, poll every 2 minutes (saves Firebase reads)
            # When active, poll every 1.5 seconds
            interval = 120 if self.paused else POLL_INTERVAL
            time.sleep(interval)

    # ---- Reminder loop — nag Atlas every 25 min if portal still open ----
    def _reminder_loop(self):
        while True:
            time.sleep(REMINDER_INTERVAL)
            try:
                user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
                host = platform.node() or "unknown"
                elapsed = int((time.time() - self._opened_at) / 60)
                _post_to_discord(
                    f"**Portal Still Active**\n"
                    f"User: {user}@{host}\n"
                    f"Open for: {elapsed} minutes\n"
                    f"Use /close to shut it down remotely if done.")
            except Exception:
                pass

    # ---- Firebase chat polling (instant message delivery) ----
    def _chat_poll_loop(self):
        """Poll Firebase every 1 second for chat messages from Atlas."""
        seen_ids = set()
        while True:
            try:
                req = urllib.request.Request(
                    f"{FIREBASE_URL}/sessions/{SESSION_ID}/chat.json",
                    headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
                with urllib.request.urlopen(req, timeout=5) as resp:
                    data = json.loads(resp.read().decode("utf-8"))

                if data:
                    for msg_id, msg_data in data.items():
                        if msg_id in seen_ids:
                            continue
                        seen_ids.add(msg_id)
                        text = msg_data.get("text", "")
                        msg_type = msg_data.get("type", "msg")
                        speak = (msg_type == "speak")
                        if text:
                            self.root.after(0, lambda t=text, s=speak:
                                            self._receive_atlas_message(t, speak=s))
            except Exception:
                pass
            time.sleep(CHAT_POLL_INTERVAL)

    def _poll_once(self):
        """Poll Firebase for commands. Session paths are unique per portal, so
        commands execute immediately without a warm-up skip."""
        try:
            req = urllib.request.Request(
                f"{FIREBASE_URL}/sessions/{SESSION_ID}/commands.json",
                headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception:
            return

        if not data:
            self.root.after(0, lambda: self._set_color(
                self.idle_color, "Portal idle — listening for Atlas' commands..."))
            return

        commands = list(data.values())

        # Only execute commands we haven't seen before
        new_commands = [
            c for c in commands
            if c.get("id") and c.get("id") not in self.executed_ids
        ]

        if not new_commands:
            # While paused, keep the amber paused color + status text
            if not self.paused:
                self.root.after(0, lambda: self._set_color(
                    self.idle_color, "Portal idle — listening for Atlas' commands..."))
            else:
                self.root.after(0, lambda: self._set_color(
                    _PORTAL_PAUSED, "Portal paused — waiting for Atlas to resume..."))
            return

        # If paused, only look for resume command — skip everything else
        if self.paused:
            for cmd in new_commands:
                if cmd.get("type") == "resume_portal":
                    self.paused = False
                    self.executed_ids.add(cmd.get("id"))
                    self._save_executed()
                    self.root.after(0, lambda: self._set_color(
                        self.idle_color, "Portal resumed — listening for Atlas' commands..."))
                    _post_to_discord(f"[Portal] Resumed by Atlas")
                    return
                # Mark other commands as executed so they don't pile up
                self.executed_ids.add(cmd.get("id"))
            self._save_executed()
            return

        self.root.after(0, lambda: self._set_color(
            _PORTAL_ACTIVE, f"Processing {len(new_commands)} command(s)..."))

        for cmd in new_commands:
            cmd_id = cmd.get("id")
            try:
                ok, result = self._execute_command(cmd)
                # Don't override paused/resumed state with "done" status
                if self.paused and cmd.get("type") == "pause_portal":
                    pass  # paused color/status already set by _execute_command
                else:
                    color = _PORTAL_DONE if ok else _PORTAL_ERROR
                    status = f"Command '{cmd.get('type')}' {'done' if ok else 'failed'}"
                    self.root.after(0, lambda c=color, s=status: self._set_color(c, s))
            except Exception as e:
                color = _PORTAL_ERROR
                status = f"Command error: {e}"
                ok = False
                self.root.after(0, lambda c=color, s=status: self._set_color(c, s))

            time.sleep(1)

            self.executed_ids.add(cmd_id)
            self._save_executed()

        if not self.paused:
            self.root.after(0, lambda: self._set_color(
                self.idle_color, "Portal idle — listening for Atlas' commands..."))

    # ---- Command execution (routes msg/speak to chat) ----
    def _execute_command(self, cmd):
        cmd_type = cmd.get("type", "")
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        tag = f"[Portal @ {user}@{host}]"

        if cmd_type == "scan":
            raw_path = cmd.get("path", ".")
            path = self._resolve_path(raw_path)
            result = _scan_directory(path)
            _post_file_to_discord(f"{tag} Scan result for: {path}", _write_temp(result))
            return True, result

        elif cmd_type == "view":
            raw_path = cmd.get("path", ".")
            path = self._resolve_path(raw_path)
            result = _simple_list_files(path)
            _post_to_discord(f"{tag} {result}")
            return True, result

        elif cmd_type == "fetch":
            path = cmd.get("path", "")
            if not path or not Path(path).exists():
                return False, f"File not found: {path}"
            _post_file_to_discord(f"{tag} Fetched: {path}", path)
            return True, f"Fetched: {path}"

        elif cmd_type == "fetchall":
            raw_path = cmd.get("path", ".")
            if raw_path == "." and self.portal_folder:
                path = self.portal_folder
            else:
                path = raw_path
            zip_path = _zip_directory(path)
            if not zip_path:
                return False, f"Could not zip: {path}"
            _post_file_to_discord(f"{tag} All files from: {path}", zip_path)
            return True, f"Sent zip of {path}"

        elif cmd_type == "check_packages":
            result = _check_packages()
            _post_file_to_discord(f"{tag} Installed packages", _write_temp(result))
            return True, result

        elif cmd_type == "download":
            url = cmd.get("url", "")
            dest = cmd.get("dest", "")
            if not url or not dest:
                return False, "Missing url or dest"
            ok, msg = _download_file(url, dest)
            _post_to_discord(f"{tag} {msg}")
            return ok, msg

        elif cmd_type == "pip_install":
            packages = cmd.get("packages", [])
            ok, msg = _pip_install(packages)
            _post_to_discord(f"{tag} pip install {' '.join(packages)}: {'OK' if ok else 'FAILED'}\n{msg[:500]}")
            return ok, msg

        elif cmd_type == "read_file":
            path = cmd.get("path", "")
            if not path:
                return False, "Missing path"
            result = _read_file(path)
            _post_file_to_discord(f"{tag} Contents of {path}", _write_temp(result))
            return True, result

        elif cmd_type == "read_file_inline":
            path = cmd.get("path", "")
            if not path:
                return False, "Missing path"
            result = _read_file(path, max_bytes=1800)
            _post_to_discord(f"{tag} `{path}`:\n```\n{result[:1800]}\n```")
            return True, result

        elif cmd_type == "delete_file":
            path = cmd.get("path", "")
            if not path:
                return False, "Missing path"
            try:
                _backup_file(path, "delete")
                Path(path).unlink()
                _post_to_discord(f"{tag} Deleted: {path}")
                return True, f"Deleted: {path}"
            except Exception as e:
                return False, f"Delete failed: {e}"

        elif cmd_type == "delete_all":
            path = cmd.get("path", "")
            if not path:
                return False, "Missing path"
            try:
                p = Path(path)
                if not p.exists() or not p.is_dir():
                    return False, f"Not a directory: {path}"
                deleted = []
                failed = []
                for f in sorted(p.iterdir()):
                    if f.is_file():
                        try:
                            _backup_file(f, "delete")
                            f.unlink()
                            deleted.append(f.name)
                        except Exception as e:
                            failed.append(f"{f.name} ({e})")
                msg_lines = [f"{tag} Deleted {len(deleted)} file(s) from: {p}"]
                if deleted:
                    msg_lines.append("Deleted:")
                    msg_lines.extend(f"  {name}" for name in deleted)
                if failed:
                    msg_lines.append("Failed:")
                    msg_lines.extend(f"  {name}" for name in failed)
                result = "\n".join(msg_lines)
                _post_to_discord(result[:1900])
                return True, result
            except Exception as e:
                return False, f"Delete all failed: {e}"

        elif cmd_type == "rename_file":
            old = cmd.get("old_path", "")
            new = cmd.get("new_path", "")
            if not old or not new:
                return False, "Missing old_path or new_path"
            try:
                _backup_file(old, "rename", {"old_path": old, "new_path": new})
                Path(old).rename(new)
                _post_to_discord(f"{tag} Renamed: {old} -> {new}")
                return True, f"Renamed: {old} -> {new}"
            except Exception as e:
                return False, f"Rename failed: {e}"

        elif cmd_type == "add":
            path = cmd.get("path", "")
            url = cmd.get("url", "")
            if not path or not url:
                return False, "Missing path or url"
            if Path(path).exists():
                return False, f"File already exists: {path}. Use /replace to overwrite."
            try:
                ok, msg = _download_file(url, path)
                _post_to_discord(f"{tag} {msg}")
                return ok, msg
            except Exception as e:
                return False, f"Add failed: {e}"

        elif cmd_type == "replace":
            path = cmd.get("path", "")
            url = cmd.get("url", "")
            if not path or not url:
                return False, "Missing path or url"
            try:
                _backup_file(path, "replace")
                Path(path).unlink()
            except Exception:
                pass
            ok, msg = _download_file(url, path)
            _post_to_discord(f"{tag} {msg}")
            return ok, msg

        elif cmd_type == "message":
            # Chat is now handled by Firebase (instant).
            # GitHub command just confirms delivery to Discord.
            msg = cmd.get("text", "")
            if msg:
                _post_to_discord(f"{tag} Message sent via Firebase: {msg[:200]}")
            return True, msg

        elif cmd_type == "speak":
            # Chat + speech is now handled by Firebase (instant).
            # GitHub command just confirms delivery to Discord.
            msg = cmd.get("text", "")
            if msg:
                _post_to_discord(f"{tag} Speak sent via Firebase: {msg[:200]}")
            return True, msg

        elif cmd_type == "run_script":
            path = cmd.get("path", "")
            if not path or not Path(path).exists():
                return False, f"Script not found: {path}"
            try:
                # Use unbuffered output + force stderr to merge with stdout
                proc = subprocess.run(
                    [sys.executable, "-u", path],
                    capture_output=True, text=True, timeout=120,
                    creationflags=CREATE_NO_WINDOW,
                    env={**os.environ, "PYTHONUNBUFFERED": "1"})
                output = proc.stdout + proc.stderr
                if not output.strip():
                    output = "(Script produced no output — it may be a GUI app that ran successfully)"
                _post_file_to_discord(f"{tag} Output of {path}", _write_temp(output))
                return proc.returncode == 0, output
            except subprocess.TimeoutExpired:
                return False, f"Script timed out after 120 seconds"
            except Exception as e:
                return False, f"Script error: {e}"

        elif cmd_type == "terminal":
            command = cmd.get("command", "")
            if not command:
                return False, "Missing command"
            try:
                is_mac = platform.system() == "Darwin"
                if is_mac:
                    proc = subprocess.run(
                        command, shell=True, capture_output=True,
                        text=True, timeout=60)
                else:
                    proc = subprocess.run(
                        command, shell=True, capture_output=True,
                        text=True, timeout=60,
                        creationflags=CREATE_NO_WINDOW)
                output = f"$ {command}\n\n--- stdout ---\n{proc.stdout}\n--- stderr ---\n{proc.stderr}\n--- exit code: {proc.returncode} ---"
                _post_file_to_discord(f"{tag} Terminal: {command}", _write_temp(output))
                return proc.returncode == 0, output
            except subprocess.TimeoutExpired:
                return False, f"Command timed out after 60 seconds"
            except Exception as e:
                return False, f"Terminal error: {e}"

        elif cmd_type == "close_portal":
            # Atlas remotely closes the portal — no confirmation needed
            _post_to_discord(f"{tag} Portal closed by Atlas.")
            self.root.after(0, lambda: self._force_close())
            return True, "Portal closing..."

        elif cmd_type == "pause_portal":
            # Atlas pauses the portal — stops executing commands
            self.paused = True
            self.root.after(0, lambda: self._set_color(
                _PORTAL_PAUSED, "Portal paused — waiting for Atlas to resume..."))
            _post_to_discord(f"{tag} Portal paused by Atlas. Use /resume to continue.")
            return True, "Portal paused"

        elif cmd_type == "resume_portal":
            # Handled in _poll_once (checked before execution when paused)
            return True, "Portal resumed"

        elif cmd_type == "resurrect_portal":
            # User closed the window but the portal is still running in background
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(1000, lambda: self.root.attributes("-topmost", False))
            self.paused = False
            _firebase_put(f"sessions/{SESSION_ID}/status", "open")
            _post_to_discord(f"{tag} Portal resurrected by Atlas.")
            return True, "Portal resurrected"

        elif cmd_type == "undo_last":
            ok, msg = _undo_last()
            _post_to_discord(f"{tag} {msg}")
            return ok, msg

        elif cmd_type == "snap_screen":
            # Screenshot requires GUI interaction; schedule on main thread.
            self.root.after(0, self._ask_screenshot)
            return True, "Screenshot permission requested"

        elif cmd_type == "snap_windows":
            self.root.after(0, self._ask_snap_windows)
            return True, "Window screenshot permission requested"

        else:
            return False, f"Unknown command type: {cmd_type}"

    def _ask_screenshot(self):
        """Show a local allow/deny dialog and capture the screen if allowed.
        The portal window is hidden before capture so it doesn't appear in the shot."""
        import tempfile
        from tkinter import Toplevel, Label, Button, Frame

        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        tag = f"[Portal @ {user}@{host}]"

        dialog = Toplevel(self.root)
        dialog.title("Atlas: Screenshot Request")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#2b2b2b")
        Label(
            dialog, text="Atlas is requesting a screenshot of your screen.",
            bg="#2b2b2b", fg="white", font=("Segoe UI", 11),
            wraplength=320, justify="center"
        ).pack(padx=20, pady=(20, 10))
        Label(
            dialog, text="The portal window will be hidden before capturing.",
            bg="#2b2b2b", fg="#aaaaaa", font=("Segoe UI", 9),
            wraplength=320, justify="center"
        ).pack(padx=20, pady=(0, 20))

        def _on_deny():
            dialog.destroy()
            _post_to_discord(f"{tag} Screenshot denied by user.")

        def _on_allow():
            dialog.destroy()
            # Hide the portal entirely before capturing so it doesn't show in the shot
            was_withdrawn = not self.root.winfo_viewable()
            self.root.withdraw()
            # Small delay so the window fully disappears
            self.root.after(500, lambda: _do_capture(was_withdrawn))

        def _do_capture(was_withdrawn):
            fd, path = tempfile.mkstemp(suffix=".png", prefix="screenshot_")
            os.close(fd)
            ok, msg = _take_screenshot(path)
            if ok and Path(path).exists():
                _post_file_to_discord(f"{tag} Screenshot", path)
                try:
                    Path(path).unlink()
                except Exception:
                    pass
            else:
                _post_to_discord(f"{tag} Screenshot failed: {msg}")
            if not was_withdrawn:
                self.root.deiconify()
                self.root.lift()

        button_row = Frame(dialog, bg="#2b2b2b")
        button_row.pack(padx=20, pady=(0, 20))
        Button(
            button_row, text="Deny", command=_on_deny,
            bg="#555555", fg="white", activebackground="#666666",
            font=("Segoe UI", 10, "bold"), width=10
        ).pack(side="left", padx=10)
        Button(
            button_row, text="Allow", command=_on_allow,
            bg="#4a9eff", fg="white", activebackground="#6ab2ff",
            font=("Segoe UI", 10, "bold"), width=10
        ).pack(side="left", padx=10)

        # Center dialog on screen
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        sw = dialog.winfo_screenwidth()
        sh = dialog.winfo_screenheight()
        dialog.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _ask_snap_windows(self):
        """Show a local allow/deny dialog and capture each visible window if allowed."""
        import tempfile
        from tkinter import Toplevel, Label, Button, Frame

        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        tag = f"[Portal @ {user}@{host}]"

        dialog = Toplevel(self.root)
        dialog.title("Atlas: Window Screenshots")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#2b2b2b")
        Label(
            dialog, text="Atlas wants a screenshot of each open window.",
            bg="#2b2b2b", fg="white", font=("Segoe UI", 11),
            wraplength=320, justify="center"
        ).pack(padx=20, pady=(20, 10))
        Label(
            dialog, text="The portal window will be hidden before capturing.",
            bg="#2b2b2b", fg="#aaaaaa", font=("Segoe UI", 9),
            wraplength=320, justify="center"
        ).pack(padx=20, pady=(0, 20))

        def _on_deny():
            dialog.destroy()
            _post_to_discord(f"{tag} Window screenshots denied by user.")

        def _on_allow():
            dialog.destroy()
            was_withdrawn = not self.root.winfo_viewable()
            self.root.withdraw()
            self.root.after(500, lambda: _do_capture(was_withdrawn))

        def _do_capture(was_withdrawn):
            try:
                ok, msg = _capture_all_windows(tag)
                if not ok:
                    _post_to_discord(f"{tag} Window screenshots failed: {msg}")
            except Exception as e:
                _post_to_discord(f"{tag} Window screenshots error: {e}")
            finally:
                if not was_withdrawn:
                    self.root.deiconify()
                    self.root.lift()

        button_row = Frame(dialog, bg="#2b2b2b")
        button_row.pack(padx=20, pady=(0, 20))
        Button(
            button_row, text="Deny", command=_on_deny,
            bg="#555555", fg="white", activebackground="#666666",
            font=("Segoe UI", 10, "bold"), width=10
        ).pack(side="left", padx=10)
        Button(
            button_row, text="Allow", command=_on_allow,
            bg="#4a9eff", fg="white", activebackground="#6ab2ff",
            font=("Segoe UI", 10, "bold"), width=10
        ).pack(side="left", padx=10)

        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        sw = dialog.winfo_screenwidth()
        sh = dialog.winfo_screenheight()
        dialog.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _close(self):
        """User clicked the X button — show a custom confirmation dialog.
        First close offers hide, fully close, or cancel.
        Second close is permanent."""
        if self.user_closed_once:
            result = self._ask_yes_no(
                "Close Portal?",
                "Are you sure you want to close the portal?\n\n"
                "This will permanently close the portal and delete the file.")
            if result:
                self._final_user_close()
        else:
            choice = self._ask_three_way(
                "Close Portal?",
                "What would you like to do?",
                "Hide", "Fully Close", "Cancel")
            if choice == "hide":
                self._user_close()
            elif choice == "fully_close":
                self._final_user_close()

    def _ask_yes_no(self, title, message):
        """Custom modal yes/no dialog centered over the portal window."""
        from tkinter import Toplevel, Label, Button, Frame
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=_BG)

        # Size
        dialog.update_idletasks()
        width, height = 360, 160

        # Position centered over the portal window
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w - width) // 2
        y = root_y + (root_h - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        # Content
        Label(dialog, text=title, font=("Segoe UI", 11, "bold"),
              bg=_BG, fg=_TEXT).pack(pady=(16, 8), padx=16)
        Label(dialog, text=message, font=("Segoe UI", 9),
              bg=_BG, fg=_TEXT, justify="center").pack(padx=16)

        result = {"value": False}

        def yes():
            result["value"] = True
            dialog.destroy()

        def no():
            result["value"] = False
            dialog.destroy()

        btn_frame = Frame(dialog, bg=_BG)
        btn_frame.pack(pady=(16, 12))
        Button(btn_frame, text="Yes", command=yes, width=10,
               bg=_PORTAL_IDLE, fg=_TEXT, activebackground="#b07bc8",
               relief="flat", cursor="hand2").pack(side="left", padx=6)
        Button(btn_frame, text="No", command=no, width=10,
               bg=_PANEL, fg=_TEXT, activebackground="#2a2a45",
               relief="flat", cursor="hand2").pack(side="left", padx=6)

        # Center dialog over root and wait
        self.root.wait_window(dialog)
        return result["value"]

    def _ask_three_way(self, title, message, btn1, btn2, btn3):
        """Custom modal three-button dialog centered over the portal window.
        Returns one of: 'btn1', 'btn2', 'btn3' (lowercased, spaces to underscores)."""
        from tkinter import Toplevel, Label, Button, Frame
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=_BG)

        width, height = 400, 180
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w - width) // 2
        y = root_y + (root_h - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        Label(dialog, text=title, font=("Segoe UI", 11, "bold"),
              bg=_BG, fg=_TEXT).pack(pady=(16, 8), padx=16)
        Label(dialog, text=message, font=("Segoe UI", 9),
              bg=_BG, fg=_TEXT, justify="center").pack(padx=16)

        result = {"value": None}

        def make_cb(val):
            def cb():
                result["value"] = val
                dialog.destroy()
            return cb

        btn_frame = Frame(dialog, bg=_BG)
        btn_frame.pack(pady=(16, 12))
        for text, val in [(btn1, btn1.lower().replace(" ", "_")),
                          (btn2, btn2.lower().replace(" ", "_")),
                          (btn3, btn3.lower().replace(" ", "_"))]:
            Button(btn_frame, text=text, command=make_cb(val), width=10,
                   bg=_PANEL, fg=_TEXT, activebackground="#2a2a45",
                   relief="flat", cursor="hand2").pack(side="left", padx=6)

        self.root.wait_window(dialog)
        return result["value"]

    def _user_close(self):
        """First user close: hide window and keep polling in background."""
        self.user_closed_once = True
        try:
            _firebase_put(f"sessions/{SESSION_ID}/status", "user-closed")
        except Exception:
            pass
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        _post_to_discord(
            f"**User closed the portal (background mode)**\n"
            f"[Portal @ {user}@{host}]\n"
            f"Session: `{SESSION_ID}`\n"
            f"Atlas can use `/resurrect` to reopen it or `/close` to end the session.\n"
            f"Closing again will permanently delete the portal.")
        self.root.withdraw()  # hide window, keep process alive

    def _final_user_close(self):
        """Second user close: permanently close and delete the portal file."""
        try:
            _mark_session_closed()
        except Exception:
            pass
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        _post_to_discord(
            f"**User permanently closed the portal**\n"
            f"[Portal @ {user}@{host}]\n"
            f"Session: `{SESSION_ID}`")
        self.root.destroy()
        try:
            portal_path = Path(__file__).resolve()
            if self.executed_file.exists():
                self.executed_file.unlink()
            if portal_path.exists():
                portal_path.unlink()
            # Clean up backup folder on final close
            bdir = _backup_dir()
            if bdir.exists():
                shutil.rmtree(bdir)
        except Exception:
            pass

    def _force_close(self):
        """Official close (remote /close): delete the portal file."""
        # Mark session as closed so the bot can delete the temp channel
        try:
            _mark_session_closed()
        except Exception:
            pass
        self.root.destroy()
        try:
            portal_path = Path(__file__).resolve()
            if self.executed_file.exists():
                self.executed_file.unlink()
            if portal_path.exists():
                portal_path.unlink()
            # Clean up backup folder on final close
            bdir = _backup_dir()
            if bdir.exists():
                shutil.rmtree(bdir)
        except Exception:
            pass


if __name__ == "__main__":
    PORTAL_FOLDER = str(Path(__file__).parent)
    # Allow color override via command-line: --color=#a0c4ff
    portal_color_override = None
    for arg in sys.argv[1:]:
        if arg.startswith("--color="):
            portal_color_override = arg.split("=", 1)[1]
    try:
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        # Register session in Firebase so the bot can create a temp channel
        _register_session(user, host, PORTAL_FOLDER)
        # Wait for bot to assign a session-specific webhook URL
        _wait_for_webhook(timeout=30)
        # Post opening message (goes to session channel if webhook was assigned)
        _post_to_discord(
            f"**Portal Opened**\n"
            f"Session: `{SESSION_ID}`\n"
            f"User: {user}@{host}\n"
            f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"Python: {platform.python_version()}\n"
            f"Folder: {PORTAL_FOLDER}")
    except Exception:
        pass

    root = tk.Tk()
    PortalWindow(root, portal_folder=PORTAL_FOLDER,
                 color_override=portal_color_override)
    root.mainloop()
