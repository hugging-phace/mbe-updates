import getpass
import io
import json
import os
import platform
import queue
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import traceback
import urllib.request
import uuid
from datetime import datetime
from tkinter import messagebox, ttk


APP_NAME = "MBE Setup Console"
APP_VERSION = "1.2.0"
DEVELOPER_NAME = "Atlas Ramoon"

# Bug reports POST to a Discord webhook (goes to developer's phone).
# Uses only the standard library so a brand-new Python install can send it.
BUG_REPORT_WEBHOOK_URL = (
    "https://discord.com/api/webhooks/1524620703259951104/"
    "fqpIEBXVWsKHy7f1iZ9xoryCpidmjPYIDuITfcwMOjBfMyS2HtJNWpVbfOetapl8vw9O"
)

# Remote package list — fetched from GitHub at runtime and merged with the
# baked-in lists below.  If the fetch fails (no internet, GitHub down), the
# baked-in lists are used as a fallback so the setup console always works.
REMOTE_PACKAGES_URL = (
    "https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/"
    "manifests/packages.json"
)

# Baked-in fallback (used if remote fetch fails)
PINNED_PACKAGES = [
    "customtkinter==6.0.0",
    "tkhtmlview==0.3.2",
    "selenium==4.45.0",
    "Pillow==12.3.0",
    "numpy==2.4.6",
    "pandas==3.0.3",
    "openpyxl==3.1.5",
    "PyMuPDF==1.28.0",
    "pypdf==6.14.2",
    "O365==2.1.9",
    "google-genai==2.10.0",
    "lxml==6.1.1",
    "fpdf2==2.8.7",
    "beautifulsoup4==4.15.0",
    "textblob==0.20.0",
]

LATEST_PACKAGES = [
    "attrs", "certifi", "cffi", "et_xmlfile", "exceptiongroup", "h11",
    "idna", "natsort", "outcome", "pycparser", "PySocks",
    "python-dateutil", "pytz", "six", "sniffio", "sortedcontainers",
    "trio", "trio-websocket", "typing_extensions", "tzdata", "urllib3",
    "websocket-client", "wsproto", "xlrd", "oauthlib",
    "language-tool-python", "keyring",
    "requests",       # HTTP client -- pairs with beautifulsoup4 / selenium
    "xlsxwriter",     # Excel writing engine -- pairs with pandas / openpyxl
]

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0

# In-memory log buffer -- no file is written to the user's machine.
# The log is only materialised as a temp file when sending a Discord
# error report, and that temp file is deleted immediately after.
LOG_BUFFER = io.StringIO()

# Strip ANSI escape sequences (colour codes, cursor moves, etc.) so pip
# output renders cleanly inside a tk.Text widget.
_ANSI_RE = re.compile(r"\x1b\[[0-9;]*[a-zA-Z]")


def _strip_ansi(text):
    return _ANSI_RE.sub("", text)


# ------------------------------------------------------------------
# Remote package list — fetch from GitHub and merge with baked-in.
# Standard-library only (urllib.request + json).
# ------------------------------------------------------------------
def _fetch_remote_packages():
    """Fetch packages.json from GitHub. Returns a dict with keys
    'pinned_packages', 'latest_packages', 'windows_only' or None on failure."""
    try:
        req = urllib.request.Request(
            REMOTE_PACKAGES_URL,
            headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            # Validate structure
            if "pinned_packages" in data and "latest_packages" in data:
                return data
            return None
    except Exception:
        return None


def _get_merged_packages():
    """Return (pinned, latest, windows_only) lists.
    Merges remote packages.json with baked-in fallback.
    Remote entries are appended (deduped by normalized name)."""
    pinned = list(PINNED_PACKAGES)
    latest = list(LATEST_PACKAGES)
    win_only = ["pywin32"]

    remote = _fetch_remote_packages()
    if remote:
        # Build sets of existing package names (normalized)
        existing_pinned = {_normalize_pkg_name(p) for p in pinned}
        existing_latest = {_normalize_pkg_name(p) for p in latest}
        # Merge remote pinned
        for pkg in remote.get("pinned_packages", []):
            if _normalize_pkg_name(pkg) not in existing_pinned:
                pinned.append(pkg)
                existing_pinned.add(_normalize_pkg_name(pkg))
        # Merge remote latest
        for pkg in remote.get("latest_packages", []):
            if _normalize_pkg_name(pkg) not in existing_latest:
                latest.append(pkg)
                existing_latest.add(_normalize_pkg_name(pkg))
        # Merge remote windows_only
        for pkg in remote.get("windows_only", []):
            if pkg not in win_only:
                win_only.append(pkg)

    return pinned, latest, win_only


# ------------------------------------------------------------------
# Remote support: error reporting to the developer's Discord.
# Standard-library only (urllib.request) so a fresh Python install
# that has never run pip can still send a report.
# ------------------------------------------------------------------
def _post_to_discord(content):
    """POST a text message to the Discord webhook. Returns (ok, error_msg)."""
    if not BUG_REPORT_WEBHOOK_URL:
        return False, "No bug-report channel is configured."
    payload = json.dumps({"content": content[:1900]}).encode("utf-8")
    try:
        req = urllib.request.Request(
            BUG_REPORT_WEBHOOK_URL,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": f"{APP_NAME}/{APP_VERSION}",
            },
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status in (200, 204):
                return True, None
            return False, f"Server returned status {resp.status}"
    except Exception as e:
        return False, str(e)


def _post_bug_report_with_log(description, log_text):
    """POST a bug report to Discord with the install log attached as a file.

    Writes log_text to a temp file, attaches it, then deletes the temp file.
    Falls back to a text-only message (with the log inlined) if the
    multipart upload fails.
    Returns (ok, error_msg).
    """
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown"
    host = platform.node() or "unknown"

    content = (
        f"**Bug Report - {APP_NAME}**\n"
        f"**Version:** {APP_VERSION}\n"
        f"**From:** {user}@{host}\n"
        f"**When:** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"**Python:** {platform.python_version()} ({platform.system()} {platform.machine()})\n"
        f"**Details:**\n{description}"
    )

    file_data = log_text.encode("utf-8") if log_text else None
    file_name = "install requirements.log"

    # If there's no log to attach, send text only.
    if not file_data:
        return _post_to_discord(content)

    # Multipart: text + log file in a single Discord message.
    try:
        boundary = f"----WebKitFormBoundary{uuid.uuid4().hex[:16]}"
        payload_json = json.dumps({"content": content[:1900]})

        body = b""
        body += f"--{boundary}\r\n".encode()
        body += b'Content-Disposition: form-data; name="payload_json"\r\n'
        body += b"Content-Type: application/json\r\n\r\n"
        body += payload_json.encode() + b"\r\n"

        body += f"--{boundary}\r\n".encode()
        body += (
            f'Content-Disposition: form-data; name="files[0]"; '
            f'filename="{file_name}"\r\n'
        ).encode()
        body += b"Content-Type: text/plain\r\n\r\n"
        body += file_data + b"\r\n"

        body += f"--{boundary}--\r\n".encode()

        req = urllib.request.Request(
            BUG_REPORT_WEBHOOK_URL,
            data=body,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "User-Agent": f"{APP_NAME}/{APP_VERSION}",
            },
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            if resp.status in (200, 204):
                return True, None
            return False, f"Server returned status {resp.status}"
    except Exception as e:
        # Multipart failed -- fall back to text-only with inlined log tail.
        tail = log_text[-1500:] if log_text else ""
        content += f"\n\n--- log (tail) ---\n{tail}"
        ok, err = _post_to_discord(content)
        if not ok:
            return False, f"{e}; fallback also failed: {err}"
        return True, None


def _send_install_error_report(title, detail):
    """Send an auto error report (with log) to Discord.

    Pulls the in-memory log buffer.  Returns (ok, error_msg).
    """
    full = f"{title}\n\n{detail}"
    log_text = LOG_BUFFER.getvalue()
    return _post_bug_report_with_log(full, log_text)


# ------------------------------------------------------------------
# Pip execution -- streams stdout in real time so the UI can show
# live progress bars and status lines instead of a blank spinner.
# ------------------------------------------------------------------
def _pip_env():
    """Build a subprocess environment that keeps pip's progress bars
    simple (no fancy Unicode) so they render cleanly in a tk.Text widget."""
    env = dict(os.environ)
    env["PIP_PROGRESS_BAR"] = "raw"
    return env


def _stream_output(proc, log_file, on_output):
    """Read proc.stdout character-by-character, parse \\r and \\n, and
    call on_output(text, is_progress) for each segment.

    is_progress=True  -> \\r was seen: the text overwrites the current line.
    is_progress=False -> \\n was seen: the text is a completed new line.

    Everything is also written to log_file.  Returns True if the process
    exited with code 0.
    """
    buffer = ""
    while True:
        byte = proc.stdout.read(1)
        if not byte:
            break
        char = byte.decode("utf-8", errors="replace")
        if char == "\r":
            if buffer:
                log_file.write(buffer)
                log_file.flush()
                on_output(buffer, True)
            buffer = ""
        elif char == "\n":
            log_file.write(buffer + "\n")
            log_file.flush()
            on_output(buffer, False)
            buffer = ""
        else:
            buffer += char

    if buffer:
        log_file.write(buffer)
        log_file.flush()
        on_output(buffer, False)

    proc.wait()
    return proc.returncode == 0


def _pip_call_streaming(args, log_file, on_output):
    """Run pip with the given args, streaming output in real time.

    on_output(text, is_progress) is called for each line / progress update.
    Returns True on success (exit code 0).
    """
    commands = [
        ([sys.executable, "-m", "pip"] + args, False),
        (subprocess.list2cmdline([sys.executable, "-m", "pip"] + args), True),
    ]

    for command, use_shell in commands:
        try:
            proc = subprocess.Popen(
                command,
                shell=use_shell,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=CREATE_NO_WINDOW,
                env=_pip_env(),
            )
        except OSError as error:
            log_file.write(f"\nCould not run pip: {error}\n")
            log_file.flush()
            on_output(f"Could not run pip: {error}", False)
            continue

        if _stream_output(proc, log_file, on_output):
            return True
    return False


def _install_package(package, log_file, on_output):
    if _pip_call_streaming(["install", "--upgrade", package], log_file, on_output):
        return True
    log_file.write(f"\nRetrying {package} with --user.\n")
    log_file.flush()
    on_output(f"Retrying {package} with --user...", False)
    return _pip_call_streaming(
        ["install", "--upgrade", "--user", package], log_file, on_output
    )


def run_installation(report):
    def on_output(text, is_progress):
        report("output", (text, is_progress))

    pinned, latest, win_only = _get_merged_packages()
    packages = pinned + latest
    if platform.system() == "Windows":
        packages += win_only

    log_file = LOG_BUFFER
    log_file.write(
        "PYTHON ENVIRONMENT\n"
        f"Version: {platform.python_version()}\n"
        f"Path: {sys.executable}\n"
        f"System: {platform.system()} {platform.machine()}\n\n"
    )

    report("status", "Updating the package installer...")
    on_output("Updating pip...", False)
    if not _pip_call_streaming(["install", "--upgrade", "pip"], log_file, on_output):
        on_output("Retrying pip with --user...", False)
        if not _pip_call_streaming(
            ["install", "--upgrade", "--user", "pip"], log_file, on_output
        ):
            raise RuntimeError(
                "Could not update pip. Check Python and your internet connection."
            )

    report("status", "Installing required packages...")
    on_output("Installing all packages...", False)
    if _pip_call_streaming(["install", "--upgrade"] + packages, log_file, on_output):
        failed = []
    else:
        failed = []
        on_output("Batch install failed -- installing packages individually...", False)
        report("progress_mode", len(packages))
        for index, package in enumerate(packages, 1):
            report("status", f"Installing {package} ({index} of {len(packages)})...")
            report("progress", index - 1)
            if not _install_package(package, log_file, on_output):
                failed.append(package)
            report("progress", index)

    report("complete", failed)


# ------------------------------------------------------------------
# Installer window -- dark, terminal-style live console.
# Standard-library tkinter/ttk only; no customtkinter needed.
# ------------------------------------------------------------------

# Colour palette (Devin-inspired: black bg, dark green console, white text).
_BG       = "#000000"
_PANEL    = "#0d1f17"
_TEXT     = "#ffffff"
_MUTED    = "#6b7d72"
_ACCENT   = "#1a8f4e"
_GREEN    = "#22c55e"
_RED      = "#ef4444"
_YELLOW   = "#eab308"
_TEAL     = "#14b8a6"
_BTN_BG   = "#0a2e1a"
_BTN_HOV  = "#1a8f4e"


def _get_installed_packages():
    """Return a set of installed package names (lowercase, no version)."""
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "list", "--format=freeze"],
            capture_output=True, text=True, timeout=30,
            creationflags=CREATE_NO_WINDOW,
        )
        installed = set()
        for line in proc.stdout.splitlines():
            line = line.strip()
            if "==" in line:
                name = line.split("==")[0].lower().replace("_", "-")
                installed.add(name)
            elif line:
                installed.add(line.lower().replace("_", "-"))
        return installed
    except Exception:
        return set()


def _normalize_pkg_name(pkg_spec):
    """Extract the package name from a spec like 'customtkinter==6.0.0'."""
    for sep in ["==", ">=", "<=", ">", "<", "~="]:
        if sep in pkg_spec:
            return pkg_spec.split(sep)[0].lower().replace("_", "-")
    return pkg_spec.lower().replace("_", "-")


def run_check_only(report):
    """Check which required packages are missing without installing."""
    pinned, latest, win_only = _get_merged_packages()
    all_packages = pinned + latest
    if platform.system() == "Windows":
        all_packages += win_only

    report("status", "Checking installed packages...")
    report("output", ("Scanning your Python environment...", False))

    installed = _get_installed_packages()
    missing = []
    present = []

    for pkg in all_packages:
        name = _normalize_pkg_name(pkg)
        if name in installed:
            present.append(pkg)
        else:
            missing.append(pkg)

    report("output", ("", False))
    if not missing:
        report("output", (f"All {len(all_packages)} packages are installed -- no missing packages.", False))
        report("output", ("Everything is up to date!", False))
        report("check_result", {"missing": [], "present": len(present), "total": len(all_packages)})
    else:
        report("output", (f"{len(missing)} missing package(s) found out of {len(all_packages)} required:", False))
        report("output", ("", False))
        for pkg in missing:
            report("output", (f"  MISSING: {pkg}", False))
        report("output", ("", False))
        report("output", (f"{len(present)} packages are already installed.", False))
        report("check_result", {"missing": missing, "present": len(present), "total": len(all_packages)})


class InstallerWindow:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.running = True
        self._line_has_content = False
        self._line_count = 0
        self._mode = None  # "install" or "check"

        root.title(f"{APP_NAME} v{APP_VERSION}")
        root.geometry("640x440")
        root.resizable(False, False)
        root.configure(bg=_BG)
        root.protocol("WM_DELETE_WINDOW", self._close)

        # ---- Choice screen first ----
        self._build_choice_screen()

        self._failure_title = ""
        self._failure_detail = ""
        self._sending = False

        self._center()
        root.after(50, self._process_events)

    # ---- Choice screen -----------------------------------------------
    def _build_choice_screen(self):
        """Show two buttons: Install All or Check For Missing."""
        for widget in self.root.winfo_children():
            widget.destroy()

        self.running = False  # not running yet

        # Header
        header = tk.Frame(self.root, bg=_BG)
        header.pack(fill="x", padx=22, pady=(40, 0))
        tk.Label(
            header,
            text="Setting up MBE tools",
            font=("Segoe UI", 14, "bold"),
            bg=_BG, fg=_TEXT, anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text="Choose an option to get started.",
            font=("Segoe UI", 10),
            bg=_BG, fg=_MUTED, anchor="w",
        ).pack(fill="x", pady=(4, 0))

        # Buttons
        btn_frame = tk.Frame(self.root, bg=_BG)
        btn_frame.pack(fill="x", padx=22, pady=(40, 8))

        tk.Button(
            btn_frame,
            text="Install All Packages",
            command=self._start_install,
            bg=_BTN_BG, fg="#ffffff",
            activebackground=_BTN_HOV, activeforeground="#ffffff",
            relief="flat", font=("Segoe UI", 12, "bold"),
            padx=30, pady=12, bd=0, cursor="hand2",
            width=30,
        ).pack(fill="x", pady=(0, 12))

        tk.Button(
            btn_frame,
            text="Check For Missing Packages",
            command=self._start_check,
            bg=_BTN_BG, fg="#ffffff",
            activebackground=_BTN_HOV, activeforeground="#ffffff",
            relief="flat", font=("Segoe UI", 12, "bold"),
            padx=30, pady=12, bd=0, cursor="hand2",
            width=30,
        ).pack(fill="x")

        # Description
        tk.Label(
            self.root,
            text=(
                "Install All Packages — installs or upgrades everything (shows full pip output).\n"
                "This can take up to 5 minutes.\n\n"
                "Check For Missing Packages — scans your environment and only reports\n"
                "what's missing, without installing anything."
            ),
            font=("Segoe UI", 9),
            bg=_BG, fg=_MUTED, anchor="w", justify="left",
        ).pack(fill="x", padx=22, pady=(16, 0))

    def _start_install(self):
        self._mode = "install"
        self._build_installer_screen()
        self.running = True
        threading.Thread(target=self._worker, daemon=True).start()
        self.root.after(50, self._process_events)

    def _start_check(self):
        self._mode = "check"
        self._build_installer_screen()
        self.running = True
        threading.Thread(target=self._check_worker, daemon=True).start()
        self.root.after(50, self._process_events)

    # ---- Installer screen --------------------------------------------
    def _build_installer_screen(self):
        """Build the console + progress + footer UI."""
        for widget in self.root.winfo_children():
            widget.destroy()

        # ---- Header ---------------------------------------------------
        header = tk.Frame(self.root, bg=_BG)
        header.pack(fill="x", padx=22, pady=(18, 0))

        title = "Installing packages..." if self._mode == "install" else "Checking packages..."
        tk.Label(
            header,
            text="Setting up MBE tools",
            font=("Segoe UI", 14, "bold"),
            bg=_BG, fg=_TEXT, anchor="w",
        ).pack(fill="x")

        self.status = tk.StringVar(value=title)
        tk.Label(
            header,
            textvariable=self.status,
            font=("Segoe UI", 10),
            bg=_BG, fg=_MUTED, anchor="w", wraplength=596,
        ).pack(fill="x", pady=(4, 0))

        # ---- Console --------------------------------------------------
        console_frame = tk.Frame(self.root, bg=_BG)
        console_frame.pack(fill="both", expand=True, padx=22, pady=12)

        self.console = tk.Text(
            console_frame,
            wrap="word",
            bg=_PANEL, fg=_TEXT,
            font=("Consolas", 9),
            relief="solid", bd=1,
            padx=10, pady=8,
            height=18,
            insertontime=0,
            highlightthickness=1,
            highlightbackground="#1a3a28",
            highlightcolor="#1a3a28",
            cursor="arrow",
        )
        self.console.pack(fill="both", expand=True)
        self.console.configure(state="disabled")

        self.console.tag_config("default",   foreground=_TEXT)
        self.console.tag_config("download",  foreground=_ACCENT)
        self.console.tag_config("install",   foreground=_YELLOW)
        self.console.tag_config("success",   foreground=_GREEN)
        self.console.tag_config("satisfied", foreground=_TEAL)
        self.console.tag_config("error",     foreground=_RED)
        self.console.tag_config("warning",   foreground=_YELLOW)
        self.console.tag_config("info",      foreground=_MUTED)

        self.console.mark_set("line_start", "1.0")
        self.console.mark_gravity("line_start", "left")

        # ---- Progress bar ---------------------------------------------
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            "Dark.Horizontal.TProgressbar",
            background=_ACCENT,
            troughcolor=_PANEL,
            borderwidth=0,
            thickness=6,
        )

        self.progress = ttk.Progressbar(
            self.root,
            style="Dark.Horizontal.TProgressbar",
            mode="indeterminate",
            length=596,
        )
        self.progress.pack(fill="x", padx=22, pady=(0, 10))
        self.progress.start(12)

        # ---- Footer ---------------------------------------------------
        self._footer = tk.Frame(self.root, bg=_BG)
        self._footer.pack(fill="x", padx=22, pady=(0, 16))

        self.report_status = tk.StringVar(value="")
        tk.Label(
            self._footer,
            textvariable=self.report_status,
            bg=_BG, fg=_MUTED,
            font=("Segoe UI", 9), anchor="w",
        ).pack(side="left")

        self.report_button = tk.Button(
            self._footer,
            text="Report Issue",
            command=self._send_report,
            bg=_BTN_BG, fg="#ffffff",
            activebackground=_BTN_HOV, activeforeground="#ffffff",
            relief="flat", font=("Segoe UI", 10),
            padx=14, pady=5, state="disabled",
            bd=0, cursor="hand2",
        )

        self.close_button = tk.Button(
            self._footer,
            text="Close",
            command=self._close,
            bg=_BTN_BG, fg="#ffffff",
            activebackground=_BTN_HOV, activeforeground="#ffffff",
            relief="flat", font=("Segoe UI", 10),
            padx=18, pady=5, state="disabled",
            bd=0, cursor="hand2",
        )
        self.close_button.pack(side="right")

        self._center()

    # ---- Layout helper ------------------------------------------------
    def _center(self):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_width()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_height()) // 2
        self.root.geometry(f"+{x}+{y}")

    # ---- Console output helpers --------------------------------------
    @staticmethod
    def _classify(text):
        """Return the colour-tag name for a line of pip output."""
        lower = text.lower().strip()
        if lower.startswith("successfully installed"):
            return "success"
        if "already satisfied" in lower:
            return "satisfied"
        if lower.startswith("error") or "no matching distribution" in lower:
            return "error"
        if "could not find a version" in lower or "fail" in lower:
            return "error"
        if lower.startswith("warning"):
            return "warning"
        if "downloading" in lower:
            return "download"
        if lower.startswith("installing collected") or lower.startswith("installing"):
            return "install"
        if "|" in text and "%" in text:
            return "download"
        if lower.startswith("using cached"):
            return "satisfied"
        if lower.startswith("collecting"):
            return "info"
        return "default"

    def _append_output(self, text, is_progress):
        """Append a line or progress-bar update to the console widget."""
        text = _strip_ansi(text)
        tag = self._classify(text) if text else "default"

        self.console.configure(state="normal")

        if is_progress:
            # \r: overwrite the current line in place.
            if not text:
                self.console.configure(state="disabled")
                return
            line_start = self.console.index("line_start")
            self.console.delete(line_start, f"{line_start} lineend")
            self.console.insert(line_start, text, tag)
            self._line_has_content = True
        else:
            # \n: advance to a new line, then insert the text.
            if self._line_has_content:
                self.console.insert("end", "\n")
            self.console.insert("end", text + "\n", tag)
            self.console.mark_set("line_start", "end")
            self._line_has_content = False
            self._line_count += 1

        # Trim very old lines to keep the widget responsive.
        if self._line_count > 2000:
            self.console.delete("1.0", "1000.0")
            self._line_count -= 1000

        self.console.see("end")
        self.console.configure(state="disabled")

    # ---- Worker thread -----------------------------------------------
    def _worker(self):
        try:
            run_installation(self._report)
        except Exception as error:
            tb = "".join(
                traceback.format_exception(type(error), error, error.__traceback__)
            )
            self._report("error", (str(error), tb))

    def _check_worker(self):
        try:
            run_check_only(self._report)
        except Exception as error:
            tb = "".join(
                traceback.format_exception(type(error), error, error.__traceback__)
            )
            self._report("error", (str(error), tb))

    def _report(self, event, value):
        self.events.put((event, value))

    # ---- Event loop (UI thread) --------------------------------------
    def _process_events(self):
        try:
            while True:
                event, value = self.events.get_nowait()
                if event == "status":
                    self.status.set(value)
                elif event == "output":
                    self._append_output(value[0], value[1])
                elif event == "progress_mode":
                    self.progress.stop()
                    self.progress.configure(
                        mode="determinate", maximum=value, value=0
                    )
                elif event == "progress":
                    self.progress.configure(value=value)
                elif event == "complete":
                    self._finish(value)
                elif event == "check_result":
                    self._finish_check(value)
                elif event == "error":
                    message, tb = value
                    self._fail(message, tb)
        except queue.Empty:
            pass

        if self.running:
            self.root.after(50, self._process_events)

    # ---- Failure / report helpers ------------------------------------
    def _show_report_button(self, title, detail):
        self._failure_title = title
        self._failure_detail = detail
        self.report_button.configure(state="normal")
        self.report_button.pack(side="right", padx=(0, 8))

    def _send_report(self):
        if self._sending:
            return
        self._sending = True
        self.report_button.configure(state="disabled", text="Sending...")
        title = self._failure_title or "Installation failed"
        detail = self._failure_detail

        def worker():
            ok, err = _send_install_error_report(title, detail)

            def done():
                self._sending = False
                if ok:
                    self.report_status.set("Report sent -- thank you!")
                    self.report_button.configure(text="Sent \u2713")
                else:
                    self.report_status.set(f"Failed to send: {err}")
                    self.report_button.configure(state="normal", text="Report Issue")

            try:
                self.root.after(0, done)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self, failed):
        self.running = False
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=1, value=1)
        if failed:
            self.status.set("Installation finished with errors.")
            self._append_output("", False)
            self._append_output(
                f"Failed: {', '.join(failed)}", False
            )
            self._append_output(
                "Click 'Report Issue' to send the error details to the developer.", False
            )
            self._show_report_button(
                "Installation finished with errors",
                f"Failed packages: {', '.join(failed)}",
            )
        else:
            self.status.set("Installation complete.")
            self._append_output("", False)
            self._append_output("All packages installed successfully.", False)
            self._append_output("The MBE tools are ready to use.", False)
        self.close_button.configure(state="normal")

    def _finish_check(self, result):
        self.running = False
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=1, value=1)
        missing = result.get("missing", [])
        present = result.get("present", 0)
        total = result.get("total", 0)
        if not missing:
            self.status.set("No Missing Packages - All Up to Date")
            self._append_output("", False)
            self._append_output(
                f"All {total} required packages are installed.", False
            )
        else:
            self.status.set(f"{len(missing)} Missing Package(s) Found")
            self._append_output("", False)
            self._append_output(
                f"Found {len(missing)} missing out of {total} required.", False
            )
            self._append_output(
                f"{present} packages are already installed.", False
            )
            # Show "Install Missing" button in the footer
            self._missing_packages = missing
            self.install_missing_button = tk.Button(
                self._footer,
                text=f"Install {len(missing)} Missing Package(s)",
                command=self._install_missing,
                bg=_BTN_BG, fg="#ffffff",
                activebackground=_BTN_HOV, activeforeground="#ffffff",
                relief="flat", font=("Segoe UI", 10, "bold"),
                padx=14, pady=5, bd=0, cursor="hand2",
            )
            self.install_missing_button.pack(side="right", padx=(0, 8))
        self.close_button.configure(state="normal")

    def _install_missing(self):
        """Install the packages that were found missing during check."""
        if not self._missing_packages:
            return
        self.install_missing_button.pack_forget()
        self.running = True
        self._mode = "install"
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self.status.set("Installing missing packages...")
        self._append_output("", False)
        self._append_output("Installing missing packages...", False)

        missing = self._missing_packages

        def _worker():
            try:
                log_file = LOG_BUFFER
                on_output = lambda text, is_progress: self._report("output", (text, is_progress))

                report = self._report
                report("progress_mode", len(missing))
                failed = []
                for index, package in enumerate(missing, 1):
                    report("status", f"Installing {package} ({index} of {len(missing)})...")
                    report("progress", index - 1)
                    if not _install_package(package, log_file, on_output):
                        failed.append(package)
                    report("progress", index)

                report("complete", failed)
            except Exception as error:
                tb = "".join(
                    traceback.format_exception(type(error), error, error.__traceback__)
                )
                self._report("error", (str(error), tb))

        self.root.after(50, self._process_events)
        threading.Thread(target=_worker, daemon=True).start()

    def _fail(self, error, tb_text=""):
        self.running = False
        self.progress.stop()
        self.status.set("Installation could not be completed.")
        self._append_output("", False)
        self._append_output(f"ERROR: {error}", False)
        if tb_text:
            self._append_output(tb_text, False)
        full_detail = error
        if tb_text:
            full_detail += f"\n\n--- Traceback ---\n{tb_text}"
        self._show_report_button("Installation could not be completed", full_detail)
        self.close_button.configure(state="normal")

    def _close(self):
        if self.running and self._mode == "install":
            messagebox.showinfo("Installation in progress", "Please wait for setup to finish.")
            return
        self.root.destroy()


if __name__ == "__main__":
    app_root = tk.Tk()
    InstallerWindow(app_root)
    app_root.mainloop()
