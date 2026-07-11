"""
Summon Portal for Atlas
=======================
Standalone remote-support launcher. Keep this file on your computer and
run it whenever you need Atlas to connect. It always downloads the latest
Python Portal for Atlas from GitHub before opening it.

Peace of mind:
  - Atlas cannot see your screen or control your mouse.
  - He can only read files you choose to share, run scripts you send,
    and send short text messages back.
"""

import os
import platform
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from tkinter import filedialog, messagebox, Tk

PORTAL_URL = (
    "https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/"
    "consoles/Python%20Portal%20for%20Atlas.pyw"
)


def main():
    root = Tk()
    root.withdraw()

    confirm = messagebox.askyesno(
        "Open a Portal for Atlas?",
        "This will open a remote IT support portal that lets Atlas\n"
        "diagnose and fix issues on your machine from afar.\n\n"
        "Peace of mind:\n"
        "Atlas cannot see your screen or control your mouse.\n"
        "He can only read files you choose to share, run scripts\n"
        "you send, and send short text messages back.\n\n"
        "You'll choose where the problem is, then a small portal\n"
        "file will be saved there for you to open.\n\n"
        "Continue?")
    if not confirm:
        root.destroy()
        return

    folder = filedialog.askdirectory(
        title="Where is the problem located? Choose a folder:")
    if not folder:
        root.destroy()
        return

    is_mac = platform.system() == "Darwin"
    ext = ".py" if is_mac else ".pyw"
    dest = os.path.join(folder, f"Python Portal for Atlas{ext}")

    try:
        busted_url = f"{PORTAL_URL}?t={int(time.time())}"
        req = urllib.request.Request(
            busted_url,
            headers={"User-Agent": "SummonPortal/1.0",
                     "Cache-Control": "no-cache"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        with open(dest, "wb") as f:
            f.write(data)
    except Exception as e:
        messagebox.showerror(
            "Download Failed",
            f"Could not download the portal:\n\n{e}\n\n"
            "Please check your internet connection and try again.")
        root.destroy()
        return

    try:
        if is_mac:
            subprocess.Popen(
                ["python3", dest],
                start_new_session=True,
            )
        else:
            subprocess.Popen(
                [sys.executable, dest],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
    except Exception as e:
        messagebox.showerror(
            "Could Not Launch",
            f"The portal was saved to:\n\n{dest}\n\n"
            f"But it could not be launched automatically:\n{e}\n\n"
            f"Please open it manually.")
        root.destroy()
        return

    messagebox.showinfo(
        "Portal Opened",
        "The portal is now opening.\n\n"
        "Leave it running and let Atlas know it's open.")
    root.destroy()


if __name__ == "__main__":
    main()
