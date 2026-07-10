# -*- coding: utf-8 -*-
"""
Unified XML Declaration Console
Launches a small picker window -> opens Air or Ocean console.
Both modes share the same UI framework; colours, header fields,
XML structure and Excel parsing differ by mode.
"""
import os
import sys
import re
import platform
import warnings
import json
import getpass
import urllib.request
import uuid
import subprocess
import traceback
from datetime import datetime
from collections import Counter
import io
import base64
import xml.etree.ElementTree as ET
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from difflib import get_close_matches
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import tkinter.ttk as ttk
import time

from PIL import Image
import customtkinter as ctk

# ==============================================================================
# RIGHT-CLICK CONTEXT MENU for CTkEntry fields (Cut / Copy / Paste / Select All)
# Monkey-patched globally so every entry in the app gets it automatically.
# ==============================================================================
def _ctk_entry_context_menu(event):
    """Show a Cut/Copy/Paste/Select All menu on right-click."""
    widget = event.widget
    # Walk up to find the underlying tkinter Entry inside the CTkEntry frame
    entry_widget = widget
    if not isinstance(entry_widget, tk.Entry):
        # CTkEntry wraps a tk.Entry — search children
        for child in widget.winfo_children():
            if isinstance(child, tk.Entry):
                entry_widget = child
                break
    if not isinstance(entry_widget, tk.Entry):
        return
    # Detect the background color from the widget hierarchy so the menu
    # blends with whatever console/window the entry lives in.
    bg_color = "#1a1a2e"  # sensible default
    w = entry_widget
    for _ in range(10):
        try:
            c = w.cget("bg")
            if c and str(c) not in ("", "SystemButtonFace"):
                bg_color = str(c)
                break
        except Exception:
            pass
        try:
            fg = w.cget("fg_color")
            if fg and str(fg) not in ("", "SystemButtonFace"):
                bg_color = str(fg)
                break
        except Exception:
            pass
        w = w.master
        if w is None:
            break
    _menu_font = ("Segoe UI", 11) if platform.system() == "Windows" else ("Helvetica", 13)
    try:
        menu = tk.Menu(entry_widget, tearoff=0, bg=bg_color, fg="#e8e8e8",
                       activebackground="#2e8b57", activeforeground="#ffffff",
                       borderwidth=1, relief="solid", activeborderwidth=0,
                       font=_menu_font)
    except Exception:
        menu = tk.Menu(entry_widget, tearoff=0, bg=bg_color, fg="#e8e8e8",
                       activebackground="#2e8b57", activeforeground="#ffffff",
                       font=_menu_font)
    menu.add_command(label="Cut",
                     command=lambda: entry_widget.event_generate("<<Cut>>"))
    menu.add_command(label="Copy",
                     command=lambda: entry_widget.event_generate("<<Copy>>"))
    menu.add_command(label="Paste",
                     command=lambda: entry_widget.event_generate("<<Paste>>"))
    menu.add_separator()
    menu.add_command(label="Select All",
                     command=lambda: entry_widget.select_range(0, "end"))
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()

# Bind right-click on the CTkEntry class itself
_orig_ctk_entry_init = ctk.CTkEntry.__init__
def _patched_ctk_entry_init(self, *args, **kwargs):
    _orig_ctk_entry_init(self, *args, **kwargs)
    self.bind("<Button-3>", _ctk_entry_context_menu, add="+")
ctk.CTkEntry.__init__ = _patched_ctk_entry_init

# Selenium for COLS upload (imported lazily so app doesn't crash if not installed)
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from natsort import natsorted
    _SELENIUM_AVAILABLE = True
except ImportError:
    _SELENIUM_AVAILABLE = False

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ==============================================================================
# EMBEDDED MBE LOGO (base64 PNG – no external file needed)
# ==============================================================================
MBE_LOGO_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAABLgAAALaCAYAAAAySeo9AAAAAXNSR0IArs4c6QAAAARnQU1BAACx"
    "jwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAFOQSURBVHhe7d0L0CVVYSfwM+/3A7emEt3ZBYpF"
    "sTABMcTJIhlZUSDBAAsuJCADArsYLdlCYyglOxg1SgSzlo9VQMH3A1gkgI+IATQmUCzMsobFlaBO"
    "YGVZNg6Iw4wyw+yc5tyx+eZ+39xH9719+v5+VV1fn7739u139/1/p08HAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIxK/0FoOF27JR6ASB7s3ZKvQAwNCcVgEwIuABokzoD"
    "runOmUI1gPZygAfIhIALgDapOmzq9zwp7AJoFwd1gEwIuABok6oCpmHPj4IugHZwMAfIhIALgDap"
    "Ilia7tw43bj7fT8A+XAgB8jEdBflAJCjYUOlqefFfsbX7Zw67PQAMF6z018AAIAsDBNuRfH9USoW"
    "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"
    "D+263ampLr/88vDv//2/T6VmEXDt7q/+6q/CK1/5ylSanoBrd9MFXDHcPfXUU3cFUDFkimHT1ACq"
    "EzzF2xLj33JY1bkVsXPrYflWw7YTcAEAwASKAVeT/fznP9+x88fdrltkmtDleoti7L70pS+lqWyu"
    "Qw89tOu0N6Fzi+LuXbxVcfv27WmJTM8tirt3092iyHB2LlsAqIynKAJQiR/84Adh27ZtqcSwer2d"
    "bFy++93vFk/qIx/33HNP+PznP59KAADtIuACoBL3339/6qMKN998c9i4cWMqNU/TAzi6+9CHPpT6"
    "AADaxX3vAJlo+u0csVH0t7zlLanUDHtqg+uLX/xiuO+++1Kpu7jYr7766j2+rw4XXnhheOc735lK"
    "zbF169bwvOc9r2iXaNTOPffc8Cu/8iupNL0zzjgj7LPPPqkUwsUXXxy2bNmSSt3F+YrvG9ZBBx1U"
    "tO/Wq7iNxfbBfvaznxV/Y1gc27Krq0bk//gf/6NohHw6cTm8973vTaXpVbW8cjBdG1wMRxtcAFTJ"
    "SQUgE00PuN7znveEa665pmg0uazTuHJHp9Hljk6jzGVTf0hOHUenMeaOTmPNZbEcx/vSl740DRnc"
    "uBqjX716dRHQlee9CT772c+G0047LZVGa8OGDeHggw9OpWpV1bj6unXrwlVXXZVKg3nyySeLWnyx"
    "If8bb7wxDa3Gf/pP/ym84x3vSKXBVbW81q5d+6wHAlTp1ltvDbfddlsqDU7AVQ8BFwAATKAYcDEe"
    "42yM/sYbb0xT0Rxr167tOq2j6DZs2JCmonpVNa6+bt26NMZqfOtb39qx//77d/2uQbrDDjssjXk4"
    "VS2v9evXpzFWL46723f228V5pXo7ly0AVEYbXADQYB//+MdTXzM88MADldSIoXeHH3540aB/rOlU"
    "hTvuuKO4vRAAoE0EXADQYDfccEN4+OGHU2n84i1zjN6KFSuKWxUPPPDANGRwsW0vD4UAANpGwAUA"
    "DRbDiE996lOpNF5PPfVU+MQnPpFKjFpsUy62fxbbrRvW97///dQHANAOAi4AaLh4m2ITmqu56aab"
    "wqOPPppKjEN8QmMVDfw/+OCDqQ8AoB0EXADQcPF2sia0e9W09sAm1bnnnpv6Brd58+bUBwDQDgIu"
    "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"
    "Sv2LIdkDDzxQNKQ9iNi4/EknnZRKg/nBD35QTEc3MTyr4umKGzZsCAcffHAqVSuun7322iuV8rV2"
    "7dpw8cUXh5e+9KVpSD1yXF7XXXddsT9X5YknnnjWk1DjbaHlds+WL1++K/SdO3duWLp0adEfxfd1"
    "biON74nv7YgPBpg/f34qhbBy5crUF8KCBQuKhzl0xGNLZ7/v9v1xPOXvbZKd0+23CAAATJoYcDXR"
    "unXrYvDWyG7vvfdOUzmc4447ruv4++l2/ghNY+tuy5YtO1atWtX1s712N998cxpb/4466qiu4+y1"
    "O+KII9KYurvyyiu7fq7fbsOGDWmM1du0aVPX78y1O/fcc3c88cQTae6ql+Pyuu6669LUVyOHZbB2"
    "7do0tc2zc/oAoDLuIwCABli4cGE49dRTU2kwV1xxRerrz8aNG8PXv/71VBpMFbWzqNZHP/rRcMgh"
    "h4S77rorDWESxdpgADAJVAsGyERT/9v9mte8JlxzzTWp1Cw53aIY3XvvveFFL3pRKvUv3gL1yCOP"
    "hOc85zlpSG8uuuii8I53vCOV+hfn7X//7//9rFujpnKL4vjE29muvPLKcPLJJ6ch1XCL4jNtn8Xj"
    "TPmWws7thLHcCZeWLVtWPDG1fCtiXC/xdsN4l17ch8rDos5nyrcddm5rjMPj61H5M92+u8ncoggA"
    "ABOouJ+jgaq4fa+uLqdbFDvWrFnT9fO9dn/xF3+RxtSb7du371i9enXXcfXanXXWWWls03OL4vi7"
    "j3zkI2lOq+EWRYa1c50AQGXcoggADXL22WenvsH0e5tivDXxoYceSqXBuD0xD3/4h3/Y2AdCAAAM"
    "S8AFwFA6TwijGvE2splu9duTeJvjHXfckUp7dvnll6e+wey///7hsMMOSyWaLgaot9xySyoBALSH"
    "XyUADKX8aHuGFx/nf8opp6TSYHoNrWJ7XTfccEMqDeass85KfXSsW7euCJH21P3VX/1V0SZUp4vt"
    "lF1yySXhvPPOC0ceeeSudpmqtG3btiJEffjhh9MQAIB20LAjQCaa2l7JG9/4xvDhD384lcYr1nyK"
    "Da137LfffpU8QW5Ujcx3xBpYa9asSaX+xeUQA4xOI9TTufjii8MFF1yQSv2Ly/qHP/xhWL16dRoy"
    "vUlqZH79+vVFw/3D2r59e7j99tvDpz71qfDpT386bNmyJb0yvBNOOCH81//6X1NpMFUtr6OOOioc"
    "ffTRqVSvV7/61cVxgWbQyDwAVXJSAchEUwOuWBPlnnvuSaXiB8tuNU86TwPriE/8ik/+6ig/Eaxj"
    "T58pP42sbqMOuKL4NMV4u+GgLrvssnDOOeek0u7i5vSCF7wg3H///WlI/4455pjwla98JZVmJuAa"
    "Tgws3/KWt4TPfe5zacjwbrzxxvC7v/u7qdS/Ji8v8iDgAgCACRQDLsZjlE9R7IhPQ+w2nl67Qw89"
    "NI2pu1tvvbXr5/rpvvSlL6Wx7dkkPUVx/fr1aYzVi8tx7ty5Xb+33+7AAw8snqI5qByWF822c/0D"
    "QGW0wQUADXT66ac/63bLft15553hu9/9birt7uMf/3jqG0ysufN7v/d7qcSonHHGGeHaa68datvo"
    "iDUEY9tfAABtIOACgAZ6znOeE0466aRUGky8TbGbeGvZNddck0qDOfXUU4vbRhm9GCy+973vTaXh"
    "fPCDH0x9AAB5E3ABQEOdffbZqW8wn/3sZ8PWrVtT6ZeqaLA81iRifM4///xw2GGHpdLgbrvttrBx"
    "48ZUAgDIl4ALABrqiCOOCPvuu28q9W/Tpk3F7WxTDXt74kEHHRRe8pKXpBLjENvmfs973pNKwxn2"
    "aYoAAE0g4AKAhopPihy2ptTll1+e+p5x1113Peupl4NQe6sZDj/88HDggQem0uC+8Y1vpD4AgHwJ"
    "uACgwV73utcN1aB4vAXtgQceSKXp2+XqVZyW2P4WzXDCCSekvsF961vfik9pTSUAgDwJuACgwVav"
    "Xh2OPvroVBpMpxbXz372s/DFL36x6B/Uq1/96rBq1apUYtxe9rKXpb7Bbd68OfzgBz9IJQCAPAm4"
    "AKDhzjrrrNQ3mE984hNh+/btRbj1+OOPp6GDOfPMM1MfTXDAAQekvuF8//vfT30AAHkScAFAwx17"
    "7LFD1Zp69NFHww033BCuuOKKNGQwcRqOOeaYVKIJVq5cmfqG8+CDD6Y+AIA8CbgAoOFiu1fD1py6"
    "8MILw+23355Kg4mNyw/THhjVq2p9xBC0jS666KLiiZPDdo899lgaIwDQVAIuAMjA2WefnfoGc++9"
    "96a+wa1bty710RQ/+clPUt9wnnrqqdQHAJAnARcAZGD//fcPa9euTaXRW7NmTTjwwANTiab4n//z"
    "f6Y+AIDJJuACgEwM29j8ME4//fTUR5N861vfSn3DWbRoUeoDAMiTgAsAMvGa17wmrFixIpVGJ4Yf"
    "f/AHf5BKNMWOHTvC1VdfnUrDWb58eeoDAMiTgAsAMrFw4cJw2mmnpdLoHHfccWMJ1pjZV77ylXD/"
    "/fen0nD+5b/8l6kPACBPAi4AyMg4blN83etel/poim3btoU//uM/TqXhHXDAAakPACBPAi4AyMiL"
    "X/zicMghh6RS/VavXh1e8YpXpBJN8ba3va2SJ2NGS5YsCfvuu28qAQDkScAFQGW+/OUvh6uuuqro"
    "rr322qIcu1tuuSXceuutRbdhw4bw3//7fy+6H/3oR0W3cePG8NhjjxXd448/nsbGdM4+++zUV78z"
    "zjgjzJ7tcqFJ/st/+S/hfe97XyoN77d/+7etYwAAAEZjRwaOPPLIHXFSq+xWrFhRdHvttdeOvffe"
    "u+j23XffHQcddFDRHXLIITvWrl1bdPH7jzvuuKI76aSTdpx//vlpyoYTx9dt2vrp4jxU5bHHHtux"
    "aNGirt9TdfcP//AP6VuHc+WVV3Ydf7/dhg0b0hirt2nTpq7f2W+3fv36NMZqPfXUUzv++I//uOt3"
    "DtO9//3vT9/Qn6YvryiOu9t39tvFeaV6O5ctAFTGv+sAqEwd7fjEGl2x2/kDs6jpFbsf/vCH4Z57"
    "7im6u+++O9x2221Fd/PNN4frr7++6K655pqiFlkbxQbf4xMV67Z27dqw3377pRLj8vTTT4ebbrop"
    "HHzwweHiiy9OQ6vze7/3e6kPACBfAi4AKvP85z8/9VG3UdymeOaZZ6Y+RikGWg888EC47rrrwn/8"
    "j/8x7L333uHYY4+trM2tMiEmANAWs9JfABouh9s5vvGNb4RXvepVqTR+MRiIbXxN59FHHw2bN29O"
    "pemdc845Re2wYcRaV7H9sVmz9nzqXb58eXjOc56TSt3FzeEFL3hBuP/++9OQasWGxx955JHi70y2"
    "bt0a/s//+T+pNL1Ym+4tb3lLKg3uK1/5SnjhC1+YStP71V/91bBw4cJUCkXNvz3tQk888UT49V//"
    "9VQaXNzu9tlnn1TqTWx/7sknnwz/7//9v2I64lMSR+Ezn/lMOPXUU1PpGTFg+8d//MdUmt44l1ev"
    "Om38DSvWIF25cmUqUZWdx0O/RQAAYNLEgKvp/umf/mlXmzVN6GJ7XTOpom2tOrrzzjsvTeHM/vzP"
    "/7zr56vozjrrrPQtM7vlllu6fn7cXZyustgGWrf3TXK3//77F+16TVVV21pt6rTBVY+dyxYAKuMW"
    "RQAqE2sd7fzRnErU7fTTTw9z585NpWqtW7cu9dFWb3vb22rbfgAARk3ABUClDj300NRH3X7lV34l"
    "vPrVr06l6sSQ8mUve1kq0UZr1qwpAlIAgLZw3ztAJnK5neODH/xgeNOb3pRK47WnNriOP/744omL"
    "TXPeeeeF//yf/3MqzeyrX/1q+J3f+Z1Uqsa73vWu8Pa3vz2VZnbrrbeGI444IpWa45Zbbgkvf/nL"
    "UykU7SfFp3ESilpb3/3ud6d96mlsD2yvvfZKJaLp2uCK21Q8zkSzZ88u2s+b2j9nzpywbNmyon/e"
    "vHm72rUr9y9YsCAsWrSo6I9tx3Xaj4vD4mvR4sWLw/z584v++Ln4+an98Xv23XffPbbh1xTa4AIA"
    "gAlUNFiSgb//+79/Vts14+za3gZXtG3bth2rV6/uOp5Burlz5+548MEH09j3TBtc+XWXXHJJWird"
    "aYNr9266NriauKz+5m/+Jk1d8+2cXgCojFsUAajUgQceGJ773OemEnWLtUPOOOOMVBreK17xirB6"
    "9epUom1OOeWUcP7556cSOYo18OJTWWO3atWqogZZ7OLTReOtp7EDgEmkWjBAJnL6b/frX//68NGP"
    "fjSVxmcSblGM4jzG25Kq8IUvfCGcfPLJqbRnblHMR1xPX/va13bd5jYdtyjubrpbFLdu3RouuOCC"
    "oj8u13gbYVS+5bDXWwtjWB11vifevRdDrKh8m2ObuEURAAAmUHE/RyZuvvnmZ90yM65uEm5R7Djy"
    "yCO7jqufLt7Gt/MHexpjb9yimEd31FFH7di8eXNaGjNzi+Lu3XS3KDKcncsWACrjFkUAKhdrzrhN"
    "cbTOPvvs1De40047bVeD1rTHOeecE/7yL/9yV+0iAIA2EnABULl4O81rX/vaVGIU4u2WsT2eYZx5"
    "5pmpjzaIt8BdccUV4bLLLtvjbYkAALkTcAFQiypqFNG7WPNqmFDxoIMOCi95yUtSidwde+yx4e//"
    "/u/DWWedlYYAALSbgAuAWuy///7Fj+xRi7VWOk8Y69YodJu97nWvS339U+OuHY466qiigf0bbrgh"
    "7LPPPmkoAED7eXIJQCZybJD33nvvDbfffvuup4NF5aeFRf2Wly9fHmbP/uX/Z6aW+9GWpyiW/dZv"
    "/VaxzPsxd+7c8OMf/3igWxw9RXH8Ypgct+XY1lbsH5anKO5uuqcoMhxPUQQAgAlUPHKKSrXpKYod"
    "V1xxRddxztSdcMIJ6dP98xTF0XarVq3asWbNmh3nnnvujssuu2zH97///TSH1fEUxd07T1Gsx85l"
    "CwCV8V8TgEz4MVC9b37zm+HBBx9MpeZ44QtfGF760pemUn82b94crr766lTqzZo1a8IBBxyQSv15"
    "+OGHw9e//vVUao54q175SZ6f+9znwi9+8YtUaq5YoSXeXtsR+2Mtxdj9i3/xL8LChQvTK/WJyyku"
    "L37pD/7gDzTUXwM1uACokpMKQCYEXAC0iYALgCppZB4AAACArAm4AAAAAMiagAsAAACArAm4AAAA"
    "AMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4"
    "AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMjarPQXgIbbsVPqzd7TTz8d/vEf/zHM"
    "mjUrrFixIg0NRX8cBkD77TzeO+ADUBknFYBMtCngip7//OeH+++/P5V2t2TJkjB37tyif/HixWH+"
    "/Pm79ce/sRzNmzev+Ew0e/bssHz58qI/Kgdny5YtC3PmzCn64/vj56Jexhs/Fz/fsXLlytQXiu+L"
    "3xstXbp017QD0J2AC4AqOakAZKJtAdfFF18cLrjgglRqr0WLFnUNzsr9gwRn5XAufseCBQuK/vg3"
    "lqPyeAcJ/RYuXFh0AHUQcAFQJScVgEy0LeB65JFHwurVq8O2bdvSEJosBmudmmzlcG5qcFYO56YL"
    "zqYL5HqtLVceb7m/l/ECzSHgAqBKTioAmWhbwBWdfvrp4dOf/nQqQf3KQV1UDs7K/eXacuXbV8sh"
    "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"
    "KYkrV64shse/8+bNK17rPCUxPlUxvj/efth5rfOUxP322y8cfvjhxWehQ8AFQJWcVAAyIeCim1tv"
    "vTW8+c1vDnfffXcaQpvFsCgGTTFMiiHS8uXLi3IMoDohVBw2Z86cImjqvNYJoWLoFD/XLaDqvBY/"
    "F4Oq+B1QJwEXAFVyUgHIhICL6cRN4/Of/3y48MILww9/+MM0lFGIYdDs2bOfFTSVazR1QqgYGC1d"
    "ujQsXLiw6OJrMYSKQVP8G8ud2k6d18pBU+dz0CYCLgCq5KQCkAkBF3vyi1/8InzsYx8L73znO8Oj"
    "jz6ahk6Ozu10U2s0xRCpE0JNrbUUw6Opt9V1q+00NYTqvAYMTsAFQJWcVAAyIeCiVz/72c/CpZde"
    "WjTovXnz5jR09DqBUzloirrdHtctaJpao6kcNHVqNHWCpvg5v5UhLwIuAKrkpAKQCQEX/Yq1uN71"
    "rneFj3zkI8UTF2NYFIOmGBDFUEj7TcA4CbgAqJKTCkAmBFwMavv27UVoBdAkAi4AquSkApAJARcA"
    "bSLgAqBKs9NfAAAAAMiSgAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAA"
    "AMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4"
    "AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACA"
    "rAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsA"
    "AACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMia"
    "gAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAA"
    "AMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAACArAm4AAAAAMiagAsAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "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"
    "HdBsTT6+OvYzagKulnOxk6eq1ltVJxXbEcCeDXKsbNLFf13H+ibNI8Co1XVsjXI4vjoHMEoCLmpR"
    "54GcPXMigTzZdydTXeu9n3Ox8zajErf3YaVRwUTLaV+w3zIqAi5oGD8ygLo4vtRv0GVc18X/uNe5"
    "HzXAJKvrGJzjsdX5gFEQcE2AUV/cjvr7mFlVJxPrFepXxX7mApKp9rRdVbHddWNbHC3LGyaDfR2m"
    "Z+eoSVUXi/EAVsW4RnkgrGreqzTK+R/WsMuv27xWtU4maTmW1Tnfk7huxmGStofppi+XZTBOTVhG"
    "VU5D2XTTM+rvG6Wq5q0J89KrKtdnTvM9DpO4fZUNM//9znNVyzoa5XfH76py2qN+p7+JqlwmbVge"
    "VEsNrglR9cG1bpN6sGr6esptOwLITV3nv27H77qO6ZN6Dgcoq+sYmzvnCOpk46pJVQe0eACoclyp"
    "tzZNnO9oFPNehSrmebp5rWp5TtKy7KhznidtvYzLJG0P001fLstgnJq0jKqclrLOdNU1/mjYea9K"
    "VfPYlPnpRdXrNad5H7VJ3L7GpcrtepTLu8rpjmwrsGdqcGXAwWwyVHESHMW2UvXJGnhGnccA++3o"
    "DbvM6zqe170t1DXdjEfd2wsAVEnANUFyuUhxcVwPyxWay49IRqmu7c15pp0cn6B/Ve83jq/QGwFX"
    "JnI4qLkAGi8nPsiTYyfTyem47hzUDHWtB8cpAHIg4KJRJvUCObcLRxe6UI0q96Xpjp/21/GpYtnn"
    "cF7MYRoZXtyeo1QERsQxFnon4JowdV2YuODJgxMkNEM8ZkapCDNy7KZJ0uGrkAYBQCMIuDLiAred"
    "qrhAHMe24cKWSRW3/WGlUVXG+aH9mrqObXvNM8p1kg5phTQIJp79AcZHwDWBmnrQdZEM0D/HTsbF"
    "tkdZvL4sS4OBITjOQn/sMDWp6sQ+9aBW13iH0fR5jaqc3yrVtez2ZFzfO0pVzWNU53xOwrpogiq3"
    "hybpZb3XMe9t3d5yWFZ1TOMgctgGmrKs+lHlcm3S/OewvfRr0revUapyWY9iGeQ2vdAmanBlxkGO"
    "qca5TVR5Agf643wwmZqw3m17eWjSeorXCx1pEABUTsA1oaq6wKhqPC6WR8vyhrz1ug9XdYxmcHWs"
    "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"
    "AABJRU5ErkJggg=="
)
# ==============================================================================
# FONT
# ==============================================================================
def get_platform_font():
    system = platform.system()
    if system == "Windows":
        return "Segoe UI"
    elif system == "Darwin":
        return "SF Pro Display"
    return "Arial"

MODERN_FONT = get_platform_font()


# ==============================================================================
# SHARED CONSTANTS
# ==============================================================================
SCRIPT_DIR = Path(__file__).parent
SCRIPT_PATH = Path(__file__).resolve()
PARENT_DIR = SCRIPT_DIR.parent

# ------------------------------------------------------------------
# Remote support: bug reporting + self-update
#   Bug reports POST to a Discord webhook (goes to developer's phone).
#   Updates are pulled from a JSON manifest hosted on GitHub.
#   Updates PRESERVE embedded data blocks (BUILTIN_TIN_NUMBERS and
#   BUILTIN_CODES) so user edits are never lost when code is replaced.
# ------------------------------------------------------------------
APP_NAME = "XML Declaration Console"
APP_VERSION = "1.1.7"
DEVELOPER_NAME = "Atlas Ramoon"
DEVELOPER_EMAIL = "atlasramoon@gmail.com"

BUG_REPORT_WEBHOOK_URL = "https://discord.com/api/webhooks/1524620703259951104/fqpIEBXVWsKHy7f1iZ9xoryCpidmjPYIDuITfcwMOjBfMyS2HtJNWpVbfOetapl8vw9O"

UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/hugging-phace/mbe-updates/main/"
    "manifests/xml-declaration-console.json"
)

# Patterns for locating embedded data blocks (used to find the start)
_TIN_DATA_PATTERN = r'BUILTIN_TIN_NUMBERS\s*=\s*\{'
_CODES_DATA_PATTERN = r'BUILTIN_CODES\s*=\s*\['


def _extract_braced_block(text, start_pattern, open_ch, close_ch):
    """Find the variable assignment matching *start_pattern* and return the
    full block from the variable name through the matching closing bracket,
    using depth counting so brackets inside string literals don't confuse it.
    Returns (block_text, start_index, end_index) or (None, -1, -1)."""
    m = re.search(start_pattern, text)
    if not m:
        return None, -1, -1
    # Position of the opening bracket
    bracket_pos = m.end() - 1
    depth = 0
    i = bracket_pos
    in_str = False
    str_ch = ""
    while i < len(text):
        ch = text[i]
        if in_str:
            if ch == "\\":
                i += 2  # skip escaped char
                continue
            if ch == str_ch:
                in_str = False
        else:
            if ch in ('"', "'"):
                in_str = True
                str_ch = ch
            elif ch == open_ch:
                depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    return text[m.start():i+1], m.start(), i + 1
        i += 1
    return None, -1, -1


def _splice_block(new_text, start_pattern, open_ch, close_ch, local_block):
    """Replace the block in *new_text* (found by *start_pattern*) with
    *local_block*.  Returns the modified text, or *new_text* unchanged if
    the block can't be found in either side."""
    _, ns, ne = _extract_braced_block(new_text, start_pattern, open_ch, close_ch)
    if ns == -1:
        return new_text
    return new_text[:ns] + local_block + new_text[ne:]


def _version_tuple(v):
    out = []
    for p in str(v).split("."):
        try:
            out.append(int(p))
        except ValueError:
            out.append(0)
    return tuple(out)


def _http_get(url, timeout=10):
    req = urllib.request.Request(
        url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8")


def _post_to_discord(content):
    """POST a message to the Discord webhook. Returns (ok, error_msg)."""
    if not BUG_REPORT_WEBHOOK_URL:
        return False, "No bug-report channel is configured."
    payload = json.dumps({"content": content[:1900]}).encode("utf-8")
    try:
        req = urllib.request.Request(
            BUG_REPORT_WEBHOOK_URL, data=payload,
            headers={"Content-Type": "application/json",
                     "User-Agent": f"{APP_NAME}/{APP_VERSION}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status in (200, 204):
                return True, None
            return False, f"Server returned status {resp.status}"
    except Exception as e:
        return False, str(e)


# ==============================================================================
# AUTOMATIC ERROR REPORTING
#   Every messagebox.showerror(...) call in this file — and any otherwise
#   unhandled exception raised inside a Tkinter callback anywhere in the
#   app — is routed through this enhanced dialog. It adds a "Report Issue"
#   button that sends the exact error text (plus a traceback, for crashes)
#   straight to the developer's Discord, so remote colleagues don't have
#   to describe the problem or screen-share.
# ==============================================================================
_WINDOW_NAME_REGISTRY = {}  # str(toplevel widget) -> (name, colors_dict)


def _register_window_name(win, name, colors=None):
    """Remember which human-readable name (and optional colour scheme) belongs
    to a given Toplevel, so error reports can say exactly which window the
    error came from and match its visual style."""
    try:
        _WINDOW_NAME_REGISTRY[str(win)] = (name, colors or {})
    except Exception:
        pass


def _get_active_window_info():
    """Best-effort guess at which window is currently active/focused.
    Returns (window_name, colors_dict)."""
    try:
        root = tk._default_root
        if not root:
            return "unknown", {}
        focused = root.focus_get()
        if focused is None:
            return "unknown", {}
        top = focused.winfo_toplevel()
        entry = _WINDOW_NAME_REGISTRY.get(str(top))
        if entry:
            return entry[0], entry[1]
        return "unknown", {}
    except Exception:
        return "unknown", {}


def _send_error_report(title, detail, window_name=""):
    """Send a diagnostic error report straight to the developer's Discord.
    If the message is too long for a single Discord message (2000 char
    limit), it is split across multiple messages so the full traceback
    always arrives."""
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown"
    host = platform.node() or "unknown"
    header = (
        f"**Auto Error Report - {APP_NAME}**\n"
        f"**Window:** {window_name or 'unknown'}\n"
        f"**Title:** {title}\n"
        f"**From:** {user}@{host}\n"
        f"**Version:** {APP_VERSION}\n"
        f"**When:** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    )
    full = header + f"**Details:**\n{detail}"
    # Discord's message limit is 2000 chars.  Leave room for the
    # "(continued)" marker and split on a newline boundary if possible.
    LIMIT = 1900
    if len(full) <= LIMIT:
        return _post_to_discord(full)
    # Multi-part send
    parts = []
    while full:
        if len(full) <= LIMIT:
            parts.append(full)
            break
        cut = full.rfind("\n", 0, LIMIT)
        if cut < LIMIT // 2:
            cut = LIMIT
        parts.append(full[:cut])
        full = full[cut:].lstrip("\n")
    ok_all = True
    err_all = None
    for i, part in enumerate(parts):
        prefix = f"[{i+1}/{len(parts)}] " if len(parts) > 1 else ""
        ok, err = _post_to_discord(prefix + part)
        if not ok:
            ok_all = False
            err_all = err
    return ok_all, err_all


_tk_messagebox_showerror = messagebox.showerror  # keep the real one as a fallback


def _show_error_with_report(title="Error", message="", parent=None,
                            traceback_text=None, window_name="", **kwargs):
    """Drop-in replacement for messagebox.showerror.

    For REAL errors (called inside an active except block, or with an
    explicit traceback_text), shows a 'Report Issue' button that sends
    the full diagnostic to Discord.

    For VALIDATION messages (called outside any except block — e.g.
    'Please enter the AWB first', 'No line items to generate'), shows
    a plain OK-only dialog.  These aren't bugs worth reporting.
    """
    # Figure out which window this error came from and its colour scheme.
    if not window_name:
        window_name, win_colors = _get_active_window_info()
    else:
        win_colors = {}

    # Detect whether we're inside an active except block.  If yes, this is
    # a real error → auto-attach the traceback and show the Report button.
    # If no, it's a validation message → plain dialog, no Report button.
    is_real_error = bool(traceback_text)
    if not traceback_text:
        try:
            exc_type, exc_val, exc_tb = sys.exc_info()
            if exc_type is not None:
                traceback_text = "".join(
                    traceback.format_exception(exc_type, exc_val, exc_tb))
                is_real_error = True
        except Exception:
            pass

    # Resolve colours to match the Report-a-Bug dialog exactly:
    #   dialog background = panel, text area = input, borders = border,
    #   text = text, muted = light, buttons = accent / accent_hover.
    # Fall back to a neutral dark palette if the window didn't register any.
    dlg_bg     = win_colors.get("panel", "#1a1a2e")
    input_bg   = win_colors.get("input", "#0f0f1a")
    border_col = win_colors.get("border", "#333333")
    accent     = win_colors.get("accent", "#b8860b")
    accent_h   = win_colors.get("accent_hover", "#daa520")
    text_col   = win_colors.get("text", "#e8e8e8")
    muted_col  = win_colors.get("light", "#aaaaaa")

    try:
        win = tk.Toplevel(parent) if parent is not None else tk.Toplevel()
    except Exception:
        return _tk_messagebox_showerror(title, message, **kwargs)

    win.title(title or "Error")
    win.configure(bg=dlg_bg)
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass
    try:
        win.grab_set()
    except Exception:
        pass

    msg_text = str(message) if message is not None else ""
    lines = msg_text.count("\n") + 1
    text_height = max(3, min(10, lines))
    w = 460

    # Helper: colour-aware right-click menu for read-only Text widgets
    # (Copy + Select All only — no Cut or Paste since the text is disabled).
    _menu_font = ("Segoe UI", 11) if platform.system() == "Windows" else ("Helvetica", 13)

    def _text_context_menu(event, text_widget, menu_bg):
        try:
            menu = tk.Menu(text_widget, tearoff=0, bg=menu_bg, fg=text_col,
                           activebackground=accent, activeforeground="#ffffff",
                           borderwidth=1, relief="solid", activeborderwidth=0,
                           font=_menu_font)
        except Exception:
            menu = tk.Menu(text_widget, tearoff=0, bg=menu_bg, fg=text_col,
                           activebackground=accent, activeforeground="#ffffff",
                           font=_menu_font)
        menu.add_command(label="Copy",
                         command=lambda: text_widget.event_generate("<<Copy>>"))
        menu.add_separator()
        menu.add_command(label="Select All",
                         command=lambda: text_widget.tag_add("sel", "1.0", "end"))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tk.Label(win, text="\u26a0  " + (title or "Error"), bg=dlg_bg,
             fg="#ff6b6b", font=(MODERN_FONT, 13, "bold"),
             anchor="w", justify="left").pack(fill="x", padx=16, pady=(14, 6))

    text_frame = tk.Frame(win, bg=dlg_bg)
    text_frame.pack(fill="both", expand=True, padx=16)
    txt = tk.Text(text_frame, wrap="word", bg=input_bg, fg=text_col,
                  relief="solid", borderwidth=1, highlightbackground=border_col,
                  highlightcolor=border_col, highlightthickness=1,
                  font=(MODERN_FONT, 10), height=text_height,
                  padx=8, pady=8)
    txt.insert("1.0", msg_text)
    txt.configure(state="disabled")
    txt.pack(fill="both", expand=True)
    txt.bind("<Button-3>", lambda e: _text_context_menu(e, txt, input_bg))

    # Collapsible traceback — only for real errors.
    tb_frame = None
    tb_text_widget = None
    if traceback_text:
        tb_toggle = tk.Label(win, text="\u25b6 Show details", bg=dlg_bg,
                             fg=muted_col, font=(MODERN_FONT, 9, "underline"),
                             cursor="hand2", anchor="w")
        tb_toggle.pack(fill="x", padx=16, pady=(4, 0))

        tb_frame = tk.Frame(win, bg=input_bg)
        tb_text_widget = tk.Text(tb_frame, wrap="word", bg=input_bg,
                                 fg=muted_col, relief="flat",
                                 font=(MODERN_FONT, 8), height=10,
                                 highlightthickness=0, padx=8, pady=6)
        tb_text_widget.insert("1.0", traceback_text)
        tb_text_widget.configure(state="disabled")
        tb_text_widget.bind("<Button-3>", lambda e: _text_context_menu(e, tb_text_widget, input_bg))

        def _toggle_tb(event=None):
            if tb_frame.winfo_ismapped():
                tb_frame.pack_forget()
                tb_toggle.configure(text="\u25b6 Show details")
            else:
                tb_frame.pack(fill="both", expand=False, padx=16, pady=(2, 4))
                tb_text_widget.pack(fill="both", expand=True)
                tb_toggle.configure(text="\u25bc Hide details")
            win.update_idletasks()
            try:
                req_h = win.winfo_reqheight()
                h = max(200, min(600, req_h))
                sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
                win.geometry(f"{w}x{h}+{int((sw-w)/2)}+{int((sh-h)/2)}")
            except Exception:
                pass
        tb_toggle.bind("<Button-1>", _toggle_tb)

    btn_frame = tk.Frame(win, bg=dlg_bg)
    btn_frame.pack(fill="x", padx=16, pady=(10, 14))

    status_lbl = tk.Label(btn_frame, text="", bg=dlg_bg,
                          fg=muted_col, font=(MODERN_FONT, 9))
    status_lbl.pack(side="left")

    full_report = msg_text if not traceback_text else (
        f"{msg_text}\n\n--- Traceback ---\n{traceback_text}")

    # Only show the Report Issue button for real errors (inside except).
    # Validation messages get a plain OK-only dialog.
    if is_real_error:
        def _report():
            report_btn.configure(state="disabled", text="Sending...")
            def worker():
                ok, err = _send_error_report(title or "Error", full_report, window_name)
                def done():
                    if ok:
                        status_lbl.configure(text="Report sent \u2014 thank you!", fg="#4caf50")
                        report_btn.configure(text="Sent \u2713")
                    else:
                        status_lbl.configure(text=f"Failed to send: {err}", fg="#ff6b6b")
                        report_btn.configure(state="normal", text="Report Issue")
                try:
                    win.after(0, done)
                except Exception:
                    pass
            threading.Thread(target=worker, daemon=True).start()

        report_btn = tk.Button(btn_frame, text="Report Issue", command=_report,
                               bg=accent, fg="white", activebackground=accent_h,
                               activeforeground="white", relief="flat",
                               font=(MODERN_FONT, 10, "bold"), padx=12, pady=5,
                               cursor="hand2", bd=0)
        report_btn.pack(side="right", padx=(6, 0))

    ok_btn = tk.Button(btn_frame, text="OK", command=win.destroy,
                       bg="#333344", fg="white", activebackground="#44445a",
                       activeforeground="white", relief="flat",
                       font=(MODERN_FONT, 10, "bold"), padx=18, pady=5,
                       cursor="hand2", bd=0)
    ok_btn.pack(side="right")

    # Size the window to fit everything we just packed (label + text +
    # button row) instead of guessing pixels from line count — guarantees
    # the Report Issue / OK buttons are never clipped off-screen.
    try:
        win.update_idletasks()
        req_h = win.winfo_reqheight()
        h = max(200, min(500, req_h))
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{int((sw-w)/2)}+{int((sh-h)/2)}")
    except Exception:
        pass

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    try:
        win.wait_window(win)
    except Exception:
        pass


messagebox.showerror = _show_error_with_report


def _tk_global_exception_handler(exc_type, exc_value, exc_tb):
    """Installed as root.report_callback_exception — catches ANY otherwise
    unhandled exception raised inside a Tkinter callback anywhere in the
    app (button clicks, bindings, etc.) and shows the same enhanced error
    dialog with a full traceback attached to the report."""
    try:
        tb_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    except Exception:
        tb_text = f"{exc_type}: {exc_value}"
    print(tb_text)  # still visible in the console window for local debugging
    try:
        _show_error_with_report(
            "Unexpected Error",
            f"Something went wrong:\n\n{exc_value}\n\n"
            f"This wasn't a message we expected, so please use 'Report Issue'\n"
            f"below to send the details straight to the developer.",
            traceback_text=tb_text)
    except Exception:
        pass


def _generate_case_number():
    """Generate a short case number like CASE-2025-A4F2."""
    year = datetime.now().year
    tag = uuid.uuid4().hex[:4].upper()
    return f"CASE-{year}-{tag}"


def _post_bug_report_with_files(description, case_number, file_paths,
                                reporter_email="", category="Bug Fix",
                                window_name=""):
    """POST a bug report to Discord with optional file attachments in ONE message.
    file_paths: list of file paths to attach (can be empty).
    reporter_email: optional email for follow-up (included in message text).
    category: Bug Fix, Feature Request, Environmental Change, or Other.
    window_name: which window the report originated from.
    Returns (ok, error_msg)."""
    if not BUG_REPORT_WEBHOOK_URL:
        return False, "No bug-report channel is configured."
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown"
    host = platform.node() or "unknown"

    content = (
        f"**Bug Report - {APP_NAME}**\n"
        f"**Case:** {case_number}\n"
        f"**Category:** {category}\n"
        + (f"**Window:** {window_name}\n" if window_name else "")
        + f"**Version:** {APP_VERSION}\n"
        f"**From:** {user}@{host}\n"
        + (f"**Email:** {reporter_email}\n" if reporter_email else "")
        + f"**When:** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"**Contact:** {DEVELOPER_EMAIL}\n"
        f"**Details:**\n{description}"
    )

    if not file_paths:
        return _post_to_discord(content)

    # Multipart: text + files in a single Discord message
    try:
        boundary = f"----WebKitFormBoundary{uuid.uuid4().hex[:16]}"
        payload_json = json.dumps({"content": content[:1900]})

        body = b""
        body += f"--{boundary}\r\n".encode()
        body += b'Content-Disposition: form-data; name="payload_json"\r\n'
        body += b"Content-Type: application/json\r\n\r\n"
        body += payload_json.encode() + b"\r\n"

        for i, fp in enumerate(file_paths):
            p = Path(fp)
            file_data = p.read_bytes()
            body += f"--{boundary}\r\n".encode()
            body += f'Content-Disposition: form-data; name="files[{i}]"; filename="{p.name}"\r\n'.encode()
            body += b"Content-Type: application/octet-stream\r\n\r\n"
            body += file_data + b"\r\n"

        body += f"--{boundary}--\r\n".encode()

        req = urllib.request.Request(
            BUG_REPORT_WEBHOOK_URL, data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}",
                     "User-Agent": f"{APP_NAME}/{APP_VERSION}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            if resp.status in (200, 204):
                return True, None
            return False, f"Server returned status {resp.status}"
    except Exception as e:
        return False, str(e)


def _post_update_applied(old_ver, new_ver):
    """Notify the developer that a user applied an update (for invoicing)."""
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown"
    host = platform.node() or "unknown"
    content = (
        f"**Update Applied - {APP_NAME}**\n"
        f"**Updated:** v{old_ver} -> v{new_ver}\n"
        f"**By:** {user}@{host}\n"
        f"**When:** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    )
    return _post_to_discord(content)


def _check_for_update():
    try:
        data = json.loads(_http_get(UPDATE_MANIFEST_URL))
        remote = data.get("version", "")
        if remote and _version_tuple(remote) > _version_tuple(APP_VERSION):
            return data
    except Exception:
        pass
    return None


def _download_and_apply_update(new_url):
    """Download the new script and apply it, preserving embedded data blocks."""
    try:
        new_text = _http_get(new_url, timeout=30)
        local_text = SCRIPT_PATH.read_text(encoding="utf-8")

        # Preserve BUILTIN_TIN_NUMBERS block
        local_tin, _, _ = _extract_braced_block(
            local_text, _TIN_DATA_PATTERN, "{", "}")
        if local_tin:
            new_text = _splice_block(
                new_text, _TIN_DATA_PATTERN, "{", "}", local_tin)

        # Preserve BUILTIN_CODES block
        local_codes, _, _ = _extract_braced_block(
            local_text, _CODES_DATA_PATTERN, "[", "]")
        if local_codes:
            new_text = _splice_block(
                new_text, _CODES_DATA_PATTERN, "[", "]", local_codes)

        tmp = SCRIPT_PATH.with_name(SCRIPT_PATH.name + ".new")
        tmp.write_text(new_text, encoding="utf-8")
        os.replace(str(tmp), str(SCRIPT_PATH))
        return True, None
    except Exception as e:
        return False, str(e)


FALLBACK_IMPORTER_NUMBER = "21640166"

BUILTIN_TIN_NUMBERS = {}

MASTER_IMPORTER_DEFAULT = "20000561"
DEFAULT_COMMODITY_CODE = "98010029"

# Common customs procedure codes for the dropdown
PROCEDURE_OPTIONS = ["HOME", "BLD MAT", "SCHOOL", "RETAILER", "SPCL ECO ZONE"]
UNIT_OPTIONS = ["NO", "KG", "G", "L", "LB", "M2", "M3", "PCS", "PR", "DOZ"]

BUILTIN_CODES = []


def truncate_desc(text, limit=50):
    """Truncate a description to `limit` chars at the nearest word boundary.
    COLS rejects overly long descriptions; the commodity code already
    carries the full legal text on the officer's side."""
    text = text.strip()
    if len(text) <= limit:
        return text
    trunc = text[:limit]
    last_space = trunc.rfind(" ")
    if last_space > 20:
        trunc = trunc[:last_space]
    return trunc


class _TreeValue:
    """Wrapper so _generate_xml can call row['cby'].get() on tree values."""
    __slots__ = ("_v",)
    def __init__(self, value=""):
        self._v = str(value) if value is not None else ""
    def get(self):
        return self._v
    def set(self, value):
        self._v = str(value) if value is not None else ""

# ==============================================================================
# SHARED HELPERS
# ==============================================================================
def money(value):
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0.00")

def money_str(value):
    return format(money(value), ".2f")

def clean_text(text):
    text = str(text).lower().strip()
    text = re.sub(r"[^a-z0-9 ]", "", text)
    return " ".join(text.split())

def clean_dock_receipt(val):
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        return val_str[:-2]
    return val_str

# Build lookup dicts from BUILTIN_CODES (description -> (code, unit, proc))
_BUILTIN_DESC_MAP = {}   # cleaned desc -> (code, unit, proc)
_BUILTIN_CODE_MAP = {}   # code -> (desc, unit, proc)
for _desc, _code, _unit, _proc in BUILTIN_CODES:
    _clean = clean_text(_desc)
    _BUILTIN_DESC_MAP[_clean] = (_code, _unit, _proc)
    _BUILTIN_CODE_MAP[_code] = (_desc, _unit, _proc)

# ==============================================================================
# MODE CONFIGURATIONS
# ==============================================================================
OCEAN_CONFIG = {
    "mode": "ocean",
    "title": "Ocean XML Declaration Console",
    "subtitle": "Build Customs SAD XML files from manifest data",
    "excel_file": PARENT_DIR / "Create Declaration Files for Ocean" / "Input Ocean Manifest Data.xlsx",
    "output_folder": PARENT_DIR / "Create Declaration Files for Ocean" / "Output XML Files",
    "transport_mode": "SEA",
    "discharge_port": "KYGEC",
    "exporter_address": "8001 NW 79th Ave",
    "vessel_fixed": None,
    "vessels": [
        ("Caribe Navigator", "SBCN"),
        ("Caribe Voyager", "CBVG"),
        ("Caribe Legend", "CL"),
        ("Tropic Mariner", "CV"),
        ("Seaboard Sun", "SS"),
        ("Bf Fortaleza", "FORTA"),
    ],
    "header_fields": [
        ("awb", "AWB / Master BL #"),
        ("vessel", "Vessel Code"),
        ("weight_lb", "Total Weight (LB)"),
        ("departure", "Departure Date (YYYY-MM-DD)"),
        ("arrival", "Arrival Date (YYYY-MM-DD)"),
        ("weight_cf", "Total Weight (CF)"),
        ("voyage", "Voyage #"),
        ("pkg_count", "Master Pkg Count"),
        ("master_importer", "Master Importer #"),
    ],
    "col_defs": [
        ("cby", "CBY", 70),
        ("dock", "Dock Receipt#", 110),
        ("desc", "Description", 210),
        ("code", "Comm. Code", 90),
        ("unit", "Unit", 55),
        ("proc", "Procedure", 85),
        ("qty", "Qty", 45),
        ("value", "Value USD", 85),
        ("freight", "Freight USD", 90),
        ("insurance", "Insurance USD", 100),
        ("importer", "Importer #", 95),
        ("origin", "Origin", 55),
    ],
    "colors": {
        "bg":         "#0a2417",
        "panel":      "#0f3320",
        "input":      "#10381f",
        "border":     "#3cb371",
        "text":       "#e8f5e9",
        "light":      "#90ee90",
        "accent":     "#2e8b57",
        "accent_hover": "#3cb371",
        "danger":     "#7a1f1f",
        "danger_hover": "#a52a2a",
        "btn_text":   "#ffffff",
    },
    "house_weight_zero": False,
    "excel_sheet": 0,
    "excel_header_labels": {
        "manifest_date": "Manifest Date",
        "arrival_date":  "Date of Arrival",
        "master_bl":     "BL#",
        "pkg_count":     "No. Pcs. BG",
        "weight_cf":     "Total Weight CF",
        "weight_lb":     "Total Weight LB",
        "vessel_code":   "Vessel Code",
        "voyage_no":     "Voyage #",
    },
    "excel_dock_col": "Dock Receipt#",
    "cols_sheet": "COLS",
    "cols_header_row": 9,
    "cols_data_start_row": 10,
    "cols_header_label_col": 2,
    "cols_header_value_col": 3,
    "cols_column_map": {
        "cby": 1,
        "dock": 0,
        "desc": 4,
        "value": 7,
        "freight": 8,
        "insurance": 9,
        "qty": 5,
        "importer": None,
    },
    "cols_header_labels": {
        "manifest_date": "Manifest Date",
        "arrival_date":  "Date of Arrival",
        "master_bl":     "BL#",
        "pkg_count":     "No. Pcs.",
        "weight_cf":     "Total Weight CF",
        "weight_lb":     None,
        "vessel_code":   None,
        "voyage_no":     None,
    },
    "commodity_sheet": "Commodity Codes",
    "commodity_desc_col": 1,
    "commodity_code_col": 2,
    "commodity_unit_col": None,
    "commodity_proc_col": None,
}

AIR_CONFIG = {
    "mode": "air",
    "title": "Air XML Declaration Console",
    "subtitle": "Build Customs SAD XML files from air manifest data",
    "excel_file": PARENT_DIR / "Create Declaration Files for Air" / "Input Air Manifest Data.xlsx",
    "output_folder": PARENT_DIR / "Create Declaration Files for Air" / "Output XML Files",
    "transport_mode": "AIR",
    "discharge_port": "KYGCM",
    "exporter_address": "2250 NW 114th Ave",
    "vessel_fixed": "KX",
    "vessels": [],
    "header_fields": [
        ("awb", "AWB #"),
        ("voyage", "KX Flight Number"),
        ("weight_lb", "Gross Weight (LB)"),
        ("departure", "Departure Date (YYYY-MM-DD)"),
        ("arrival", "Arrival Date (YYYY-MM-DD)"),
        ("weight_cf", "Gross Volume (CF)"),
        ("pkg_count", "# Bags/Pcs"),
        ("master_importer", "Master Importer #"),
    ],
    "col_defs": [
        ("cby", "CBY", 70),
        ("dock", "Package #", 110),
        ("desc", "Description", 210),
        ("code", "Comm. Code", 90),
        ("unit", "Unit", 55),
        ("proc", "Procedure", 85),
        ("qty", "Qty", 45),
        ("value", "Value USD", 85),
        ("freight", "Freight USD", 90),
        ("insurance", "Insurance USD", 100),
        ("importer", "Importer #", 95),
        ("origin", "Origin", 55),
    ],
    "colors": {
        "bg":         "#1a0a1e",
        "panel":      "#2a1030",
        "input":      "#3a1840",
        "border":     "#d946b5",
        "text":       "#fce4ec",
        "light":      "#f48fb1",
        "accent":     "#c2185b",
        "accent_hover": "#e91e63",
        "danger":     "#7a1f1f",
        "danger_hover": "#a52a2a",
        "btn_text":   "#ffffff",
    },
    "house_weight_zero": True,
    "excel_sheet": "Create",
    "excel_header_labels": {
        "manifest_date": "Manifest Date",
        "arrival_date":  "Date of Arrival",
        "master_bl":     "AWB#",
        "pkg_count":     "# Bags/Pcs AWB",
        "weight_cf":     "GrossVol CF",
        "weight_lb":     "GrossWt LB",
        "vessel_code":   None,
        "voyage_no":     "KX Flight Number",
    },
    "excel_dock_col": "Package#",
    "cols_sheet": " COLS",
    "cols_header_row": 8,
    "cols_data_start_row": 9,
    "cols_header_label_col": 3,
    "cols_header_value_col": 4,
    "cols_column_map": {
        "cby": 2,
        "dock": 16,
        "desc": 6,
        "value": 8,
        "freight": 9,
        "insurance": 10,
        "qty": None,
        "importer": 4,
    },
    "cols_header_labels": {
        "manifest_date": "Manifest Date",
        "arrival_date":  "Date of Arrival",
        "master_bl":     "AWB#",
        "pkg_count":     "# Bags/Pcs AWB",
        "weight_cf":     None,
        "weight_lb":     None,
        "vessel_code":   None,
        "voyage_no":     None,
    },
    "commodity_sheet": "Commodity Codes",
    "commodity_desc_col": 1,
    "commodity_code_col": 2,
    "commodity_unit_col": 3,
    "commodity_proc_col": None,
}

CONFIGS = {"ocean": OCEAN_CONFIG, "air": AIR_CONFIG}


# ==============================================================================
# XML STRUCTURE BUILDER  (parameterised by mode config)
# ==============================================================================
def build_sad_structure(cfg, hdr, bill_type, bill_no, content_text,
                        pkg_count, pkg_type="BG",
                        gross_wt=None, gross_vol=None):
    root = ET.Element("SADEntry")
    ET.SubElement(root, "Date").text = hdr["arrival"]
    ET.SubElement(root, "Regime").text = "IM1"

    exporter = ET.SubElement(root, "Exporter")
    ET.SubElement(exporter, "Name").text = "E-box Logistics"
    ET.SubElement(exporter, "Address").text = cfg["exporter_address"]
    ET.SubElement(exporter, "City").text = "Miami"
    ET.SubElement(exporter, "State").text = "FL"
    ET.SubElement(exporter, "Country").text = "USA"
    ET.SubElement(exporter, "Phone").text = "3457451400"

    consignment = ET.SubElement(root, "Consignment")
    ET.SubElement(consignment, "DepartureDate").text = hdr["departure"]
    ET.SubElement(consignment, "ArrivalDate").text = hdr["arrival"]
    ET.SubElement(consignment, "ExportCountry").text = "USA"
    ET.SubElement(consignment, "ImportCountry").text = "CYM"
    ET.SubElement(consignment, "ShippingPort").text = "USMIA"
    ET.SubElement(consignment, "DischargePort").text = cfg["discharge_port"]
    ET.SubElement(consignment, "TransportMode").text = cfg["transport_mode"]

    shipment = ET.SubElement(root, "Shipment")
    vessel_code = cfg["vessel_fixed"] or hdr["vessel"]
    ET.SubElement(shipment, "VesselCode").text = vessel_code
    ET.SubElement(shipment, "VoyageNo").text = str(hdr["voyage"])
    ET.SubElement(shipment, "ShippingAgent").text = "KX"
    ET.SubElement(shipment, "ShippingAgentName").text = "KX"
    ET.SubElement(shipment, "BillNumber").text = bill_no
    ET.SubElement(shipment, "BillType").text = bill_type
    if bill_type == "HOUSE":
        ET.SubElement(shipment, "MasterBillNumber").text = hdr["awb"]

    packages = ET.SubElement(root, "Packages")
    ET.SubElement(packages, "PkgCount").text = str(pkg_count)
    ET.SubElement(packages, "PkgType").text = pkg_type
    wt = gross_wt if gross_wt is not None else hdr["weight_lb"]
    vol = gross_vol if gross_vol is not None else hdr["weight_cf"]
    ET.SubElement(packages, "GrossWt").text = money_str(wt)
    ET.SubElement(packages, "GrossWtUnit").text = "LB"
    ET.SubElement(packages, "GrossVol").text = money_str(vol)
    ET.SubElement(packages, "GrossVolUnit").text = "CF"
    ET.SubElement(packages, "Contents").text = content_text

    return root


# ==============================================================================
# LAUNCHER WINDOW  (small dark picker)
# ==============================================================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

LAUNCHER_BG = "#1a1a2e"
LAUNCHER_PANEL = "#16213e"
LAUNCHER_ACCENT = "#0f3460"
LAUNCHER_TEXT = "#e0e0e0"


# ==============================================================================
# SUPPORT MIXIN — bug report + self-update functionality shared by all windows
# ==============================================================================
class SupportMixin:
    """Provides bug-reporting and self-update features to any window class
    that has a ``self.win`` (CTkToplevel), ``self._support_bg`` (bg color
    string), and ``self._window_name`` (identifying label for bug reports).

    Call ``_set_support_palette(colors)`` in ``__init__`` to theme the
    support dialogs to match the window's own colour scheme."""

    # ---- Palette setup --------------------------------------------------
    def _set_support_palette(self, colors):
        """Populate instance-level dialog colours from a *colors* dict
        (keys: bg, input, border, text, light, accent, accent_hover).
        Falls back to sensible defaults for missing keys."""
        bg       = colors.get("bg", "#1a1a2e")
        panel    = colors.get("panel", bg)
        inp      = colors.get("input", "#0f0f1a")
        border   = colors.get("border", "#333333")
        text     = colors.get("text", "#e8e8e8")
        light    = colors.get("light", "#aaaaaa")
        accent   = colors.get("accent", "#2e8b57")
        accent_h = colors.get("accent_hover", accent)

        self._SUP_DLG_BG       = panel
        self._SUP_DLG_LIGHT    = text
        self._SUP_DLG_MUTED    = light
        self._SUP_DLG_INPUT    = inp
        self._SUP_DLG_BORDER   = border
        self._SUP_DLG_ACCENT   = accent
        self._SUP_DLG_ACCENT_H = accent_h
        self._SUP_DLG_GREEN    = accent
        self._SUP_DLG_GREEN_H  = accent_h
        # Info-box: dark background (input colour) so light text is readable,
        # with the accent colour used only for the heading
        self._SUP_DLG_INFO_BG  = inp
        self._SUP_DLG_INFO_HDR = accent
        self._SUP_DLG_INFO_TXT = text

    # ---- Tooltip helpers ------------------------------------------------
    def _show_tooltip(self, widget, text):
        self._hide_tooltip()
        x = widget.winfo_rootx() + 20
        y = widget.winfo_rooty() + 18
        line_count = text.count("\n") + 1
        est_height = line_count * 16 + 16
        screen_h = widget.winfo_screenheight()
        if y + est_height > screen_h - 20:
            y = widget.winfo_rooty() - est_height - 8
        tw = tk.Toplevel(self.win)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=text, justify="left",
                         bg=self._SUP_DLG_BG, fg=self._SUP_DLG_LIGHT,
                         relief="solid", bd=1, padx=10, pady=8,
                         font=(MODERN_FONT, 10), wraplength=320)
        label.pack()
        self._tooltip_win = tw

    def _hide_tooltip(self):
        if self._tooltip_win is not None:
            try:
                self._tooltip_win.destroy()
            except Exception:
                pass
            self._tooltip_win = None

    # ---- Support icon ---------------------------------------------------
    def _attach_support_icon(self, parent_frame):
        """Create the bug / apply-fixes icon button in *parent_frame*."""
        self._support_btn = ctk.CTkButton(
            parent_frame, text="\U0001f41e", width=34, height=28,
            fg_color=self._support_bg, hover_color="#24507a",
            corner_radius=6,
            font=("Segoe UI Emoji", 15),
            command=lambda: self._on_support_click(self._window_name))
        self._support_btn.pack(side="right")
        self._support_btn.bind("<Enter>",
            lambda e: self._show_tooltip(self._support_btn, self._support_tooltip))
        self._support_btn.bind("<Leave>",
            lambda e: self._hide_tooltip())

    # ---- Background update check ---------------------------------------
    def _check_update_bg(self):
        result = _check_for_update()
        if result:
            self._pending_update = result
            try:
                self.win.after(0, self._refresh_support_icon)
            except Exception:
                pass

    def _refresh_support_icon(self):
        try:
            if self._pending_update:
                ver = self._pending_update.get("version", "?")
                self._support_btn.configure(
                    text="Apply Fixes", fg_color="#b8860b",
                    hover_color="#daa520",
                    font=(MODERN_FONT, 11, "bold"), width=90)
                self._support_tooltip = f"Apply Fixes \u2014 v{ver} available"
            else:
                self._support_btn.configure(
                    text="\U0001f41e", fg_color=self._support_bg,
                    hover_color="#24507a",
                    font=("Segoe UI Emoji", 15), width=34)
                self._support_tooltip = "Report a Bug"
        except Exception:
            pass

    # ---- Click handler --------------------------------------------------
    def _on_support_click(self, window_name):
        self._hide_tooltip()
        if self._pending_update:
            self._apply_fixes_dialog()
        else:
            self._report_bug_dialog(window_name)

    # ---- Bug report dialog ---------------------------------------------
    def _report_bug_dialog(self, window_name):
        dlg = ctk.CTkToplevel(self.win)
        dlg.title(f"Report a Bug \u2014 {window_name}")
        dlg.configure(fg_color=self._SUP_DLG_BG)
        dlg.transient(self.win)
        dlg.grab_set()
        dlg.resizable(False, False)
        w, h = 460, 640
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(dlg, text=f"Report a Bug \u2014 {window_name}",
                     font=(MODERN_FONT, 15, "bold"),
                     text_color=self._SUP_DLG_LIGHT).pack(
            anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(dlg,
            text=f"Describe the problem in as much detail as you can \u2014 "
                 f"what you\ndid, what happened, and what you expected. This "
                 f"goes directly\nto the developer ({DEVELOPER_NAME}).",
            font=(MODERN_FONT, 11), text_color=self._SUP_DLG_MUTED,
            anchor="w", justify="left").pack(
            anchor="w", padx=16, pady=(0, 6))

        # ---- Check for Updates bar ----
        update_bar = ctk.CTkFrame(dlg, fg_color="transparent")
        update_bar.pack(fill="x", padx=16, pady=(0, 6))

        check_btn = ctk.CTkButton(
            update_bar, text="Check for Updates",
            command=lambda: None,  # set below
            fg_color=self._SUP_DLG_ACCENT, hover_color=self._SUP_DLG_ACCENT_H,
            width=140, height=26, corner_radius=4,
            font=(MODERN_FONT, 10, "bold"))
        check_btn.pack(side="left")

        update_status = ctk.CTkLabel(
            update_bar, text="",
            font=(MODERN_FONT, 10), text_color=self._SUP_DLG_MUTED,
            anchor="w")
        update_status.pack(side="left", padx=(10, 0))

        # Spinner animation state
        _spin = {"frame": 0, "running": False}

        def _spin_step():
            if not _spin["running"]:
                return
            _spin["frame"] = (_spin["frame"] + 1) % 8
            dots = "." * ((_spin["frame"] % 4) + 1)
            update_status.configure(text=f"Checking{dots}")
            dlg.after(200, _spin_step)

        def _do_check():
            _spin["running"] = True
            check_btn.configure(state="disabled", text="Checking...")
            _spin_step()

            def _worker():
                result = _check_for_update()
                dlg.after(0, lambda: _check_done(result))

            threading.Thread(target=_worker, daemon=True).start()

        def _check_done(result):
            _spin["running"] = False
            check_btn.configure(state="normal", text="Check for Updates")
            if result:
                ver = result.get("version", "?")
                self._pending_update = result
                update_status.configure(
                    text=f"v{ver} available!",
                    text_color=self._SUP_DLG_LIGHT)
                # Replace the check button with an "Update Now" button
                check_btn.configure(
                    text=f"Update Now \u2014 v{ver}",
                    command=self._apply_fixes_dialog,
                    fg_color="#b8860b", hover_color="#daa520",
                    width=180)
            else:
                update_status.configure(
                    text=f"You're up to date (v{APP_VERSION})",
                    text_color=self._SUP_DLG_MUTED)

        check_btn.configure(command=_do_check)

        # Category selector
        cat_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        cat_frame.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(cat_frame, text="Category:",
                     font=(MODERN_FONT, 11, "bold"),
                     text_color=self._SUP_DLG_LIGHT,
                     anchor="w").pack(anchor="w", pady=(0, 2))
        cat_var = ctk.StringVar(value="Bug Fix")
        cat_row = ctk.CTkFrame(cat_frame, fg_color="transparent")
        cat_row.pack(fill="x")
        for cat in ("Bug Fix", "Feature Request",
                     "Environmental Change", "Other"):
            ctk.CTkRadioButton(
                cat_row, text=cat, variable=cat_var, value=cat,
                font=(MODERN_FONT, 10), text_color=self._SUP_DLG_LIGHT,
                fg_color=self._SUP_DLG_ACCENT,
                hover_color=self._SUP_DLG_ACCENT_H,
                border_color=self._SUP_DLG_BORDER).pack(
                side="left", padx=(0, 12))

        box = ctk.CTkTextbox(
            dlg, height=140, fg_color=self._SUP_DLG_INPUT,
            border_color=self._SUP_DLG_BORDER, border_width=1,
            corner_radius=4, text_color=self._SUP_DLG_LIGHT,
            font=(MODERN_FONT, 11))
        box.pack(fill="both", expand=True, padx=16)
        box.focus_set()

        # Optional email field
        email_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        email_frame.pack(fill="x", padx=16, pady=(4, 0))
        ctk.CTkLabel(
            email_frame,
            text="Your email (optional \u2014 for updates on this report)",
            font=(MODERN_FONT, 10), text_color=self._SUP_DLG_MUTED,
            anchor="w").pack(anchor="w")
        email_var = ctk.StringVar(value="")
        email_entry = ctk.CTkEntry(
            email_frame, textvariable=email_var, height=28,
            fg_color=self._SUP_DLG_INPUT, border_color=self._SUP_DLG_BORDER,
            border_width=1, corner_radius=4, text_color=self._SUP_DLG_LIGHT,
            font=(MODERN_FONT, 11))
        email_entry.pack(fill="x", pady=(1, 0))

        # "What to expect" info box
        info = ctk.CTkFrame(dlg, fg_color=self._SUP_DLG_INFO_BG, corner_radius=6)
        info.pack(fill="x", padx=16, pady=(6, 8))
        ctk.CTkLabel(info, text="What to expect",
                     font=(MODERN_FONT, 11, "bold"),
                     text_color=self._SUP_DLG_INFO_HDR,
                     anchor="w").pack(anchor="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(info,
            text=f"\u2022 {DEVELOPER_NAME} will review your report within "
                 f"24-48 hours\n"
                 f"\u2022 When a fix is ready, click 'Check for Updates' above\n"
                 f"   or relaunch the console \u2014 the bug icon will show "
                 f"'Apply Fixes'\n"
                 f"\u2022 {DEVELOPER_NAME} will coordinate with management if "
                 f"the fix\n"
                 f"   requires substantial work\n"
                 f"\u2022 Write down your case number for follow-up: contact "
                 f"{DEVELOPER_EMAIL}",
            font=(MODERN_FONT, 10), text_color=self._SUP_DLG_INFO_TXT,
            anchor="w", justify="left").pack(
            anchor="w", padx=10, pady=(0, 8))

        def _next():
            desc = box.get("0.0", "end").strip()
            if not desc:
                messagebox.showwarning("Empty",
                                       "Please describe the bug first.")
                return
            email = email_var.get().strip()
            category = cat_var.get()
            dlg.destroy()
            self._show_attach_files_dialog(
                desc, email, category, window_name)

        ctk.CTkButton(btns, text="Next", command=_next,
                      fg_color=self._SUP_DLG_GREEN,
                      hover_color=self._SUP_DLG_GREEN_H, width=100,
                      height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left")
        ctk.CTkButton(btns, text="Cancel", command=dlg.destroy,
                      fg_color="#667788", hover_color="#556677", width=90,
                      height=30, corner_radius=5,
                      font=(MODERN_FONT, 11)).pack(side="left", padx=(8, 0))

    # ---- Attach files dialog -------------------------------------------
    def _show_attach_files_dialog(self, description, reporter_email="",
                                  category="Bug Fix", window_name=""):
        """After the bug description, offer to attach sample files.
        The actual Discord send happens here \u2014 text + files in ONE message."""
        case_num = _generate_case_number()
        dlg = ctk.CTkToplevel(self.win)
        dlg.title(f"Attach Files? (Case {case_num})")
        dlg.configure(fg_color=self._SUP_DLG_BG)
        dlg.transient(self.win)
        dlg.grab_set()
        dlg.resizable(False, False)
        w, h = 460, 420
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(dlg, text="Attach Sample Files?",
                     font=(MODERN_FONT, 15, "bold"),
                     text_color=self._SUP_DLG_LIGHT).pack(
            anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(dlg,
            text=f"Based on your description, {DEVELOPER_NAME} may need "
                 f"sample\nfiles to reproduce the issue in a test "
                 f"environment.\n\n"
                 f"The console script itself is always attached "
                 f"automatically\nso {DEVELOPER_NAME} gets your exact version "
                 f"for testing.\n\n"
                 f"If you'd like, attach additional Excel or CSV files "
                 f"below.\nThis is completely optional and files are sent "
                 f"privately\nalongside your bug report.",
            font=(MODERN_FONT, 11), text_color=self._SUP_DLG_MUTED,
            anchor="w", justify="left").pack(
            anchor="w", padx=16, pady=(0, 8))

        file_list_frame = ctk.CTkFrame(
            dlg, fg_color=self._SUP_DLG_INPUT, corner_radius=4,
            border_width=1, border_color=self._SUP_DLG_BORDER)
        file_list_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        file_list_label = ctk.CTkLabel(
            file_list_frame, text="No files selected",
            font=(MODERN_FONT, 11), text_color=self._SUP_DLG_MUTED,
            anchor="w", justify="left")
        file_list_label.pack(anchor="w", padx=10, pady=10)

        selected_files = []

        def _pick_files():
            paths = filedialog.askopenfilenames(
                title="Select files to attach",
                filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"),
                           ("All files", "*.*")])
            if paths:
                selected_files.clear()
                selected_files.extend(paths)
                names = [Path(p).name for p in paths]
                display = "\n".join(names[:6])
                if len(names) > 6:
                    display += f"\n...and {len(names)-6} more"
                file_list_label.configure(
                    text=display, text_color=self._SUP_DLG_LIGHT)

        def _submit():
            submit_btn.configure(state="disabled", text="Sending...")
            skip_btn.configure(state="disabled")

            def _worker():
                all_files = [str(SCRIPT_PATH)] + list(selected_files)
                ok, err = _post_bug_report_with_files(
                    description, case_num, all_files,
                    reporter_email, category, window_name)
                self.win.after(0, lambda: _done(ok, err))
            threading.Thread(target=_worker, daemon=True).start()

        def _done(ok, err):
            if ok:
                dlg.destroy()
                n_files = len(selected_files)
                if n_files:
                    messagebox.showinfo("Bug Reported",
                        f"Case {case_num} submitted with {n_files} "
                        f"file(s).\n\n"
                        f"Your report and files have been sent to "
                        f"{DEVELOPER_NAME}.\n"
                        "When a fix is ready, you'll see 'Apply Fixes' "
                        "here.")
                else:
                    messagebox.showinfo("Bug Reported",
                        f"Case {case_num} submitted.\n\n"
                        f"Your report has been sent to "
                        f"{DEVELOPER_NAME}.\n"
                        "When a fix is ready, you'll see 'Apply Fixes' "
                        "here.")
            else:
                submit_btn.configure(state="normal", text="Submit Report")
                skip_btn.configure(state="normal")
                messagebox.showerror("Could Not Send",
                    f"The report could not be sent:\n{err}\n\n"
                    "Check your internet connection and try again.")

        ctk.CTkButton(btns, text="Browse Files", command=_pick_files,
                      fg_color=self._SUP_DLG_ACCENT,
                      hover_color=self._SUP_DLG_ACCENT_H, width=120,
                      height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left")
        submit_btn = ctk.CTkButton(
            btns, text="Submit Report", command=_submit,
            fg_color=self._SUP_DLG_GREEN, hover_color=self._SUP_DLG_GREEN_H,
            width=130, height=30, corner_radius=5,
            font=(MODERN_FONT, 11, "bold"))
        submit_btn.pack(side="left", padx=(8, 0))
        skip_btn = ctk.CTkButton(
            btns, text="Skip & Send", command=_submit,
            fg_color="#667788", hover_color="#556677", width=110,
            height=30, corner_radius=5, font=(MODERN_FONT, 11))
        skip_btn.pack(side="left", padx=(8, 0))

    # ---- Apply fixes dialog --------------------------------------------
    def _apply_fixes_dialog(self):
        upd = self._pending_update
        if not upd:
            return

        dlg = ctk.CTkToplevel(self.win)
        dlg.title("Apply Fixes")
        dlg.configure(fg_color=self._SUP_DLG_BG)
        dlg.transient(self.win)
        dlg.grab_set()
        dlg.resizable(False, False)
        w, h = 460, 380
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=16, pady=(0, 14))

        ver = upd.get("version", "?")
        ctk.CTkLabel(dlg, text=f"Fixes Available  (v{ver})",
                     font=(MODERN_FONT, 15, "bold"),
                     text_color=self._SUP_DLG_LIGHT).pack(
            anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(
            dlg,
            text=f"From {DEVELOPER_NAME}. Click Apply Fixes to download,\n"
                 f"then choose how to restart.",
            font=(MODERN_FONT, 11), text_color=self._SUP_DLG_MUTED,
            anchor="w").pack(anchor="w", padx=16, pady=(0, 8))

        box = ctk.CTkTextbox(
            dlg, height=170, fg_color=self._SUP_DLG_INPUT,
            border_color=self._SUP_DLG_BORDER, border_width=1,
            corner_radius=4, text_color=self._SUP_DLG_LIGHT,
            font=(MODERN_FONT, 11), wrap="word")
        box.pack(fill="both", expand=True, padx=16)
        box.insert("0.0", upd.get("changelog", "(no description provided)"))
        box.configure(state="disabled")

        def _apply():
            apply_btn.configure(state="disabled", text="Applying...")

            def _worker():
                ok, err = _download_and_apply_update(upd.get("url", ""))
                self.win.after(0, lambda: _done(ok, err))
            threading.Thread(target=_worker, daemon=True).start()

        def _done(ok, err):
            if ok:
                old_ver = APP_VERSION
                new_ver = upd.get("version", "?")
                threading.Thread(
                    target=lambda: _post_update_applied(old_ver, new_ver),
                    daemon=True).start()
                # Clear the pending update so the bug icon goes back
                # to 🐞 — the update is on disk, they just need to
                # restart.  This way they can still report a new bug
                # if they click "Later" instead of restarting.
                self._pending_update = None
                try:
                    self.win.after(0, self._refresh_support_icon)
                except Exception:
                    pass
                dlg.destroy()
                # Show a restart-choice dialog with two options:
                # "Restart Only" and "Restart + Restore Progress".
                self._show_restart_choice(new_ver)
            else:
                apply_btn.configure(state="normal", text="Apply Fixes")
                messagebox.showerror("Update Failed",
                    f"Could not apply the update:\n{err}\n\n"
                    "Check your internet connection and try again.")

        apply_btn = ctk.CTkButton(
            btns, text="Apply Fixes", command=_apply,
            fg_color="#b8860b", hover_color="#daa520",
            width=130, height=30, corner_radius=5,
            font=(MODERN_FONT, 11, "bold"))
        apply_btn.pack(side="left")
        ctk.CTkButton(btns, text="Later", command=dlg.destroy,
                      fg_color="#667788", hover_color="#556677", width=90,
                      height=30, corner_radius=5,
                      font=(MODERN_FONT, 11)).pack(side="left", padx=(8, 0))

    def _show_restart_choice(self, new_ver):
        """After an update is downloaded, ask the user how to restart.
        Offers two options: plain restart, or restart with session
        preservation (saves UploadWindow treeview + decl numbers)."""
        dlg = ctk.CTkToplevel(self.win)
        dlg.title("Restart to Apply")
        dlg.configure(fg_color=self._SUP_DLG_BG)
        dlg.transient(self.win)
        dlg.grab_set()
        dlg.resizable(False, False)
        w, h = 480, 220
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        ctk.CTkLabel(dlg, text=f"Updated to v{new_ver}",
                     font=(MODERN_FONT, 15, "bold"),
                     text_color=self._SUP_DLG_LIGHT).pack(
            anchor="w", padx=20, pady=(18, 4))
        ctk.CTkLabel(dlg,
            text="The console needs to restart to apply the changes.\n"
                 "Choose how you'd like to restart:",
            font=(MODERN_FONT, 11), text_color=self._SUP_DLG_MUTED,
            anchor="w", justify="left").pack(anchor="w", padx=20, pady=(0, 12))

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=20, pady=(0, 16))

        def _restart_only():
            dlg.destroy()
            self._restart_app(preserve_session=False)

        def _restart_preserve():
            dlg.destroy()
            self._restart_app(preserve_session=True)

        ctk.CTkButton(btns, text="Restart Only",
                      command=_restart_only,
                      fg_color="#2471a3", hover_color="#3498db",
                      width=120, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left")
        ctk.CTkButton(btns, text="Restart + Restore Progress",
                      command=_restart_preserve,
                      fg_color="#b8860b", hover_color="#daa520",
                      width=200, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=(8, 0))
        ctk.CTkButton(btns, text="Later",
                      command=dlg.destroy,
                      fg_color="#444444", hover_color="#555555",
                      width=70, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11)).pack(side="right")

    # ---- Session save / restore (for live bug-fix restarts) ------------
    # When the user clicks "Restart + Restore Progress", we serialize the
    # UploadWindow's treeview rows + declaration numbers to a temp JSON
    # file with "used": false.  On the next launch, UploadWindow.__init__
    # calls _restore_session() which:
    #   1. Reads the JSON
    #   2. Checks "used" is false and the file is < 24 hours old
    #   3. Restores the data into the treeview + dicts
    #   4. Marks "used": true and writes it back
    #   5. Tries to delete the file
    #
    # The "used" flag is the primary one-time-use guarantee: even if the
    # file can't be deleted (locked temp dir, antivirus, strict perms),
    # it will never be restored again because "used" is true.
    #
    # The 24-hour timestamp is a safety net for edge cases like a crash
    # during restart — the file survives with "used": false, and the
    # next successful launch picks it up.  But after 24 hours it's
    # considered stale and ignored, so it can't haunt someone days later.

    _SESSION_FILE = "mbe_session_restore.json"
    _SESSION_EXPIRY_SECONDS = 86400  # 24 hours (safety net only)

    def _save_session_for_restore(self):
        """Serialize the UploadWindow's treeview + decl data to a temp
        JSON file so it can be restored after a restart-restart update.
        Cross-platform: uses tempfile.gettempdir()."""
        import tempfile, os, json, datetime

        # Find the UploadWindow via the launcher — it holds the
        # declaration numbers and treeview that we want to preserve.
        upload_win = None
        try:
            upload_win = self.launcher._upload_win
        except Exception:
            pass
        if not upload_win:
            return False

        try:
            rows = []
            for item in upload_win._tree.get_children():
                vals = upload_win._tree.set(item)
                rows.append({
                    "file":   vals.get("file", ""),
                    "cby":    vals.get("cby", ""),
                    "status": vals.get("status", ""),
                    "docs":   vals.get("docs", ""),
                    "decl":   vals.get("decl", ""),
                })

            session = {
                "written_at": datetime.datetime.now().isoformat(),
                "used": False,
                "rows": rows,
                "decl_numbers": dict(upload_win._decl_numbers),
                "master_decl": upload_win._master_decl,
                "xml_folder": str(upload_win._xml_folder) if upload_win._xml_folder else None,
            }

            path = os.path.join(tempfile.gettempdir(), self._SESSION_FILE)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(session, f)
            return True
        except Exception:
            return False

    def _restore_session(self):
        """Check for a session-restore JSON in the temp folder.
        If it exists, is unused ("used": false), and is < 24 hours old,
        restore treeview rows and declaration data.  Then mark it as
        used and try to delete the file.
        Returns True if data was restored, False otherwise."""
        import tempfile, os, json, datetime

        path = os.path.join(tempfile.gettempdir(), self._SESSION_FILE)
        if not os.path.exists(path):
            return False

        try:
            with open(path, "r", encoding="utf-8") as f:
                session = json.load(f)
        except Exception:
            # Corrupted file — try to delete, move on.
            try:
                os.remove(path)
            except Exception:
                pass
            return False

        # Check "used" flag — primary one-time-use guarantee.
        # If already used, this file can never be restored again,
        # even if it couldn't be deleted last time.
        if session.get("used", False):
            try:
                os.remove(path)
            except Exception:
                pass
            return False

        # Safety net: reject anything older than 24 hours.  This handles
        # the case where a file persists for days (locked temp dir) —
        # we don't want stale data from last week reappearing.
        try:
            written = datetime.datetime.fromisoformat(session["written_at"])
            age = (datetime.datetime.now() - written).total_seconds()
            if age > self._SESSION_EXPIRY_SECONDS:
                try:
                    os.remove(path)
                except Exception:
                    pass
                return False
        except Exception:
            try:
                os.remove(path)
            except Exception:
                pass
            return False

        # Restore treeview rows
        restored_count = 0
        try:
            for row in session.get("rows", []):
                self._tree.insert("", "end", values=(
                    row.get("file", ""),
                    row.get("cby", ""),
                    row.get("status", ""),
                    row.get("docs", ""),
                    row.get("decl", ""),
                ))
                restored_count += 1
        except Exception:
            pass

        # Restore declaration data
        try:
            self._decl_numbers = dict(session.get("decl_numbers", {}))
        except Exception:
            self._decl_numbers = {}
        try:
            self._master_decl = session.get("master_decl", "")
        except Exception:
            self._master_decl = ""
        try:
            folder = session.get("xml_folder")
            if folder:
                from pathlib import Path
                self._xml_folder = Path(folder)
        except Exception:
            pass

        # Mark as used and write back — this is the one-time-use lock.
        # Even if the delete fails, the "used": true flag prevents any
        # future session from restoring the same data again.
        try:
            session["used"] = True
            with open(path, "w", encoding="utf-8") as f:
                json.dump(session, f)
        except Exception:
            pass

        # Now try to delete the file.  If it fails, the "used" flag
        # above guarantees it won't be used again.
        try:
            os.remove(path)
        except Exception:
            pass

        return restored_count > 0

    # ---- Restart --------------------------------------------------------
    def _restart_app(self, preserve_session=False):
        if preserve_session:
            self._save_session_for_restore()
        try:
            subprocess.Popen([sys.executable, str(SCRIPT_PATH)])
        except Exception:
            pass
        try:
            self.win.destroy()
        except Exception:
            pass
        # Also destroy the launcher root if accessible
        try:
            if hasattr(self, 'launcher') and self.launcher:
                self.launcher.root.destroy()
        except Exception:
            pass
        sys.exit(0)


class Launcher:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.report_callback_exception = _tk_global_exception_handler
        self.root.title("XML Declaration Builder")
        self.root.configure(fg_color=LAUNCHER_BG)
        _register_window_name(self.root, "Launcher", {"bg": LAUNCHER_BG, "panel": LAUNCHER_PANEL, "accent": LAUNCHER_ACCENT, "accent_hover": "#1a4a7a", "text": "#e8e8e8"})
        w, h = 420, 370
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{int((sw-w)/2)}+{int((sh-h)/2)}")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Logo
        try:
            _img = Image.open(io.BytesIO(base64.b64decode(MBE_LOGO_B64)))
            _lw, _lh = _img.size
            _tw = 90
            _th = max(1, int(_lh * _tw / _lw))
            _logo = ctk.CTkImage(light_image=_img, dark_image=_img, size=(_tw, _th))
            ctk.CTkLabel(self.root, image=_logo, text="").pack(pady=(20, 4))
        except Exception:
            pass

        ctk.CTkLabel(self.root, text="XML Declaration Builder",
                     font=(MODERN_FONT, 16, "bold"), text_color=LAUNCHER_TEXT).pack(pady=(4, 2))
        ctk.CTkLabel(self.root, text="Choose a shipment type to begin",
                     font=(MODERN_FONT, 11), text_color="#888888").pack(pady=(0, 16))

        btn_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))

        ctk.CTkButton(btn_frame, text="Build XML files for Ocean",
                      command=lambda: self._open("ocean"),
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=170, height=42, corner_radius=8,
                      font=(MODERN_FONT, 13, "bold")).pack(side="left", padx=10)

        ctk.CTkButton(btn_frame, text="Build XML files for Air",
                      command=lambda: self._open("air"),
                      fg_color="#c2185b", hover_color="#e91e63",
                      width=170, height=42, corner_radius=8,
                      font=(MODERN_FONT, 13, "bold")).pack(side="left", padx=10)

        # TIN Numbers + Codes buttons
        btn_row = ctk.CTkFrame(self.root, fg_color="transparent")
        btn_row.pack(pady=(0, 20))
        ctk.CTkButton(btn_row, text="TIN Numbers",
                      command=self._open_tin,
                      fg_color="#0f3460", hover_color="#1a4a7a",
                      width=120, height=28, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=5)
        ctk.CTkButton(btn_row, text="Item Codes",
                      command=self._open_codes,
                      fg_color="#0f3460", hover_color="#1a4a7a",
                      width=120, height=28, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=5)

        # Divider — separates the build/lookup tools from the standalone
        # upload action. corner_radius=0 so the bar renders as a crisp line
        # (rounded corners collapse a short bar and make it invisible), and a
        # taller/brighter bar so it stays visible under display scaling.
        ctk.CTkFrame(self.root, fg_color="#9aa0d4", height=3, width=300,
                     corner_radius=0).pack(pady=(18, 16))

        # Upload Declarations button (standalone)
        ctk.CTkButton(self.root, text="Upload Declarations to COLS",
                      command=self._open_upload,
                      fg_color="#6c3483", hover_color="#8e44ad",
                      width=280, height=36, corner_radius=8,
                      font=(MODERN_FONT, 12, "bold")).pack(pady=(0, 22))

        self._console = None
        self._codes_win = None
        self._tin_win = None
        self._upload_win = None

    def _open_tin(self):
        if self._tin_win is not None:
            try:
                self._tin_win.win.deiconify()
                self._tin_win.win.lift()
                return
            except Exception:
                pass
        self._tin_win = TINWindow(self)

    def _open_codes(self):
        if self._codes_win is not None:
            try:
                self._codes_win.win.deiconify()
                self._codes_win.win.lift()
                return
            except Exception:
                pass
        self._codes_win = CodesWindow(self)

    def _open_upload(self):
        if self._upload_win is not None:
            try:
                self._upload_win.win.deiconify()
                self._upload_win.win.lift()
                return
            except Exception:
                pass
        self.root.withdraw()
        self._upload_win = UploadWindow(self)

    def _open(self, mode):
        self.root.withdraw()
        cfg = CONFIGS[mode]
        self._console = ConsoleWindow(self, cfg)

    def show(self):
        self.root.deiconify()

    def _on_close(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ==============================================================================
# CONSOLE WINDOW  (full editor – parameterised by mode config)
# ==============================================================================
class ConsoleWindow(SupportMixin):
    def __init__(self, launcher, cfg):
        self.launcher = launcher
        self.cfg = cfg
        c = cfg["colors"]

        # Support-mixin state
        self._pending_update = None
        self._support_tooltip = "Report a Bug"
        self._tooltip_win = None
        self._support_bg = c["bg"]
        self._window_name = f"XML Builder - {cfg['mode'].title()}"
        self._set_support_palette(c)

        self.win = ctk.CTkToplevel(launcher.root)
        self.win.title(cfg["title"])
        self.win.configure(fg_color=c["bg"])
        _register_window_name(self.win, self._window_name, c)
        WIN_W, WIN_H = 1290, 760
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry(f"{WIN_W}x{WIN_H}+{int((sw-WIN_W)/2)}+{max(0, int((sh-WIN_H)/2))}")
        self.win.minsize(1100, 560)
        self.win.protocol("WM_DELETE_WINDOW", self._close_all)

        self.header_entries = {}
        self._desc_to_code = {}
        self._code_to_unit = {}
        self._code_to_proc = {}
        self._manifest_dir = None  # set when Build from Manifest is used

        # Vessel display helpers
        self._vessel_display = [f"{name} ({code})" for name, code in cfg["vessels"]]
        self._code_to_display = {code.upper(): f"{name} ({code})" for name, code in cfg["vessels"]}

        self._build_ui()
        self._show_placeholder()
        self.win.after(100, self.win.focus_force)

        # Check for updates in the background
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    # -- UI construction --------------------------------------------------
    def _build_ui(self):
        c = self.cfg["colors"]
        win = self.win

        # ---- Title bar ----
        title_frame = ctk.CTkFrame(win, fg_color="transparent")
        title_frame.pack(fill="x", padx=20, pady=(16, 6))

        # Back button (left)
        ctk.CTkButton(title_frame, text="\u2190 Back",
                      command=self._go_back,
                      fg_color=c["accent"], hover_color=c["accent_hover"],
                      width=80, height=28, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold"),
                      text_color=c["btn_text"]).pack(side="left")

        # Logo (right)
        try:
            _img = Image.open(io.BytesIO(base64.b64decode(MBE_LOGO_B64)))
            _lw, _lh = _img.size
            _tw = 130
            _th = max(1, int(_lh * _tw / _lw))
            _logo = ctk.CTkImage(light_image=_img, dark_image=_img, size=(_tw, _th))
            ctk.CTkLabel(title_frame, image=_logo, text="").pack(side="right")
        except Exception:
            pass

        ctk.CTkLabel(title_frame, text=self.cfg["title"],
                     font=(MODERN_FONT, 20, "bold"), text_color=c["light"]).pack(side="left", padx=(16, 0))
        ctk.CTkLabel(title_frame, text=self.cfg["subtitle"],
                     font=(MODERN_FONT, 12), text_color=c["text"]).pack(side="left", padx=(14, 0))

        # ---- Header fields ----
        header_panel = ctk.CTkFrame(win, fg_color=c["panel"], corner_radius=8)
        header_panel.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(header_panel, text="Shipment Header",
                     font=(MODERN_FONT, 14, "bold"), text_color=c["light"]).pack(anchor="w", padx=14, pady=(10, 4))

        header_grid = ctk.CTkFrame(header_panel, fg_color="transparent")
        header_grid.pack(fill="x", padx=10, pady=(0, 12))

        for idx, (key, label) in enumerate(self.cfg["header_fields"]):
            r, col = divmod(idx, 3)
            block = ctk.CTkFrame(header_grid, fg_color="transparent")
            block.grid(row=r, column=col, padx=8, pady=6, sticky="w")
            ctk.CTkLabel(block, text=label, width=170, anchor="w",
                         font=(MODERN_FONT, 11), text_color=c["text"]).pack(side="left")
            if key == "vessel" and self._vessel_display:
                widget = ctk.CTkComboBox(block, width=190, height=28, values=self._vessel_display,
                                         fg_color=c["input"], border_color=c["border"], border_width=1,
                                         corner_radius=5, text_color=c["text"],
                                         button_color=c["accent"], button_hover_color=c["accent_hover"],
                                         dropdown_fg_color=c["panel"], dropdown_text_color=c["text"],
                                         dropdown_hover_color=c["accent"])
                widget.set("")
            else:
                widget = ctk.CTkEntry(block, width=190, height=28, fg_color=c["input"],
                                      border_color=c["border"], border_width=1, corner_radius=5,
                                      text_color=c["text"])
            widget.pack(side="left")
            self.header_entries[key] = widget

        # Sensible defaults
        self._set_header("pkg_count", "1")
        self._set_header("weight_lb", "0.00")
        self._set_header("weight_cf", "0.00")
        self._set_header("master_importer", MASTER_IMPORTER_DEFAULT)
        if self.cfg.get("vessel_fixed") and "vessel" in self.header_entries:
            self._set_header("vessel", self.cfg["vessel_fixed"])
        if self.cfg["mode"] == "air" and "voyage" in self.header_entries:
            self._set_header("voyage", "909")

        # ---- Action bar ----
        action_frame = ctk.CTkFrame(win, fg_color="transparent")
        action_frame.pack(fill="x", padx=20, pady=(0, 8))

        self._make_btn(action_frame, "Build from Manifest", self._build_from_manifest, width=170).pack(side="left", padx=(0, 8))
        self._make_btn(action_frame, "+ Add Row", lambda: self._add_row(),
                       fg="#1f5e3a" if self.cfg["mode"] == "ocean" else "#a8630a",
                       hover="#2e8b57" if self.cfg["mode"] == "ocean" else "#cc7a14",
                       width=110).pack(side="left", padx=8)
        self._make_btn(action_frame, "Clear All",
                       lambda: (self._clear_all_rows(), self._show_placeholder()),
                       fg=c["danger"], hover=c["danger_hover"], width=100).pack(side="left", padx=8)
        self._make_btn(action_frame, "Delete Row",
                       self._remove_selected_row,
                       fg=c["danger"], hover=c["danger_hover"], width=100).pack(side="left", padx=8)
        self._make_btn(action_frame, "Generate XML Files", self._generate_xml, width=190).pack(side="right", padx=(8, 0))

        # ---- Items table header ----
        ctk.CTkLabel(win, text="Line Items  (grouped by CBY on output)  -  double-click a cell to edit",
                     font=(MODERN_FONT, 14, "bold"), text_color=c["light"]).pack(anchor="w", padx=22, pady=(4, 2))

        # ---- Treeview table ----
        tree_frame = ctk.CTkFrame(win, fg_color=c["panel"], corner_radius=6)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(4, 8))

        col_keys = [k for k, _, _ in self.cfg["col_defs"]]

        self._tree = ttk.Treeview(tree_frame, columns=col_keys, show="headings",
                                  selectmode="browse")
        for key, label, w in self.cfg["col_defs"]:
            self._tree.heading(key, text=label, anchor="w")
            # Description column stretches to fill remaining width
            self._tree.column(key, width=w, minwidth=w, anchor="w",
                              stretch=(key == "desc"))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=c["input"], foreground=c["text"],
                        fieldbackground=c["input"], bordercolor=c["border"],
                        rowheight=28, font=(MODERN_FONT, 11))
        style.configure("Treeview.Heading",
                        background=c["panel"], foreground=c["light"],
                        font=(MODERN_FONT, 11, "bold"), relief="flat")
        style.map("Treeview",
                  background=[("selected", c["accent"])],
                  foreground=[("selected", c["btn_text"])])
        style.map("Treeview.Heading", background=[("active", c["accent"])])
        # Tag for rows with blank descriptions (red highlight)
        self._tree.tag_configure("blank_desc", background="#5c1a1a", foreground="#ff6b6b")

        tree_scroll = ctk.CTkScrollbar(tree_frame, command=self._tree.yview)
        self._tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y", padx=(2, 4), pady=4)
        self._tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)

        self._tree.bind("<Double-1>", self._on_tree_double_click)
        self._tree.bind("<Delete>", lambda e: self._remove_selected_row())
        # Tab navigation between cells
        self._tree.bind("<Tab>", self._on_tab_next)
        self._tree.bind("<Shift-Tab>", self._on_tab_prev)
        self._editing_row_id = None
        self._editing_key = None

        self._placeholder = None
        self._editing_window = None

        # ---- Status bar (with support icon in the corner) ----
        status_frame = ctk.CTkFrame(win, fg_color="transparent")
        status_frame.pack(fill="x", padx=22, pady=(0, 12))

        self.status_label = ctk.CTkLabel(
            status_frame, text="Standalone mode. Use 'Build from Manifest' to load from a manifest Excel, "
                      "or enter line items manually.",
            font=(MODERN_FONT, 11), text_color=c["text"], anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)

        self._attach_support_icon(status_frame)

    # -- Small helpers ----------------------------------------------------
    def _make_btn(self, parent, text, command, fg=None, hover=None, width=150):
        c = self.cfg["colors"]
        return ctk.CTkButton(parent, text=text, command=command, width=width, height=34,
                             fg_color=fg or c["accent"], hover_color=hover or c["accent_hover"],
                             text_color=c["btn_text"],
                             font=(MODERN_FONT, 12, "bold"), corner_radius=6)

    def _set_header(self, key, value):
        w = self.header_entries[key]
        w.delete(0, "end")
        w.insert(0, value)

    def _set_vessel(self, code):
        code = str(code).strip()
        if self.cfg.get("vessel_fixed"):
            if "vessel" in self.header_entries:
                self._set_header("vessel", self.cfg["vessel_fixed"])
        elif "vessel" in self.header_entries:
            self.header_entries["vessel"].set(self._code_to_display.get(code.upper(), code))

    def _set_status(self, text):
        self.status_label.configure(text=text)

    # -- Row management (Treeview-based) ---------------------------------
    def _show_placeholder(self):
        """Placeholder disabled — empty space is cleaner."""
        pass

    def _hide_placeholder(self):
        """Placeholder disabled."""
        pass

    def _on_tree_double_click(self, event):
        if self._editing_window is not None:
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None
        self._autocomplete_list = None
        region = self._tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self._tree.identify_column(event.x)
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        col_idx = int(col.replace("#", "")) - 1
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        if col_idx < 0 or col_idx >= len(col_keys):
            return
        key = col_keys[col_idx]
        self._start_edit(row_id, key)

    def _on_tab_next(self, event):
        """Tab moves to the next cell to the right, or first cell of next row."""
        if self._editing_window is not None:
            return  # let the editing widget handle Tab
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        # Find current column
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        # Try next column
        next_idx = cur_idx + 1
        if next_idx >= len(col_keys):
            # Wrap to next row, first column
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return  # last row, last col - nowhere to go
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._start_edit(row_id, col_keys[next_idx])
        return "break"

    def _on_tab_prev(self, event):
        """Shift+Tab moves to the previous cell to the left, or last cell of prev row."""
        if self._editing_window is not None:
            return  # let the editing widget handle Tab
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        prev_idx = cur_idx - 1
        if prev_idx < 0:
            # Wrap to previous row, last column
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    prev_idx = len(col_keys) - 1
                else:
                    return
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._start_edit(row_id, col_keys[prev_idx])
        return "break"

    def _start_edit(self, row_id, key):
        """Open an edit popup on the given cell. Shared by double-click and Tab nav."""
        if self._editing_window is not None:
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None
        if self._autocomplete_list is not None:
            self._autocomplete_list.place_forget()
            self._autocomplete_list.destroy()
            self._autocomplete_list = None

        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        col_idx = col_keys.index(key)
        col = f"#{col_idx + 1}"
        bbox = self._tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        cx = self._tree.winfo_rootx() + x
        cy = self._tree.winfo_rooty() + y
        c = self.cfg["colors"]
        current_val = self._tree.set(row_id, key)
        self._editing_window = tk.Toplevel(self.win)
        self._editing_window.overrideredirect(True)
        self._editing_window.geometry(f"{w}x{h}+{cx}+{cy}")
        self._editing_window.attributes("-topmost", True)
        self._editing_row_id = row_id
        self._editing_key = key
        # Convert cell position to self.win-relative coordinates.
        # bbox gives coords relative to the tree widget. Use winfo_rootx/y
        # (reliable at click time since the window is already rendered) to
        # walk the full widget hierarchy offset back to self.win.
        tree_rel_x = self._tree.winfo_rootx() - self.win.winfo_rootx()
        tree_rel_y = self._tree.winfo_rooty() - self.win.winfo_rooty()
        self._edit_cell_x = x + tree_rel_x
        self._edit_cell_y = y + tree_rel_y
        self._edit_cell_h = h
        # Select the row so Tab navigation knows where we are
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)

        # Helper: save current and move to next cell (Tab)
        def tab_to_next():
            self._save_current_edit()
            self.win.after(10, self._tab_to_next_cell)

        # Helper: save current and move to prev cell (Shift+Tab)
        def tab_to_prev():
            self._save_current_edit()
            self.win.after(10, self._tab_to_prev_cell)

        if key == "proc":
            widget = ctk.CTkComboBox(self._editing_window, values=PROCEDURE_OPTIONS,
                                     width=w, height=h, fg_color=c["input"],
                                     border_color=c["border"], border_width=1,
                                     corner_radius=2, text_color=c["text"],
                                     button_color=c["accent"],
                                     button_hover_color=c["accent_hover"],
                                     dropdown_fg_color=c["panel"],
                                     dropdown_text_color=c["text"],
                                     dropdown_hover_color=c["accent"])
            widget.set(current_val)
            widget.pack(fill="both", expand=True)
            widget.bind("<<ComboboxSelected>>",
                        lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<FocusOut>",
                        lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_to_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_to_prev(), "break"))
            widget.focus_set()
        elif key == "desc":
            # Description column gets autocomplete from BUILTIN_CODES
            widget = ctk.CTkEntry(self._editing_window, width=w, height=h,
                                  fg_color=c["input"], border_color=c["accent"],
                                  border_width=2, corner_radius=2, text_color=c["text"],
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            self._editing_row_id = row_id
            self._editing_key = key
            self._desc_entry = widget
            # Build autocomplete list popup
            self._autocomplete_list = tk.Listbox(self.win, height=8,
                                                 bg=c["input"], fg=c["text"],
                                                 selectbackground=c["accent"],
                                                 selectforeground=c["btn_text"],
                                                 borderwidth=1, relief="solid",
                                                 font=(MODERN_FONT, 11),
                                                 activestyle="none")
            self._autocomplete_list.place_forget()
            self._autocomplete_list.bind("<<ListboxSelect>>",
                                         lambda e: self._on_autocomplete_select(row_id))
            widget.bind("<KeyRelease>", lambda e: self._update_autocomplete(widget, row_id))
            widget.bind("<Return>", lambda e: (self._finish_desc_edit(row_id, widget.get()), self._tab_to_next_cell()))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>",
                        lambda e: self.win.after(150, lambda: self._finish_desc_edit(row_id, widget.get())))
            widget.bind("<Down>", lambda e: self._autocomplete_focus(widget))
            widget.bind("<Tab>", lambda e: (tab_to_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_to_prev(), "break"))
            # No initial dropdown — suggestions appear as the user types
        elif key == "cby":
            # CBY column gets autocomplete from BUILTIN_TIN_NUMBERS
            widget = ctk.CTkEntry(self._editing_window, width=w, height=h,
                                  fg_color=c["input"], border_color=c["accent"],
                                  border_width=2, corner_radius=2, text_color=c["text"],
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            self._editing_row_id = row_id
            self._editing_key = key
            # Build autocomplete list popup showing "CBY# - Name (TIN)"
            self._autocomplete_list = tk.Listbox(self.win, height=8,
                                                 bg=c["input"], fg=c["text"],
                                                 selectbackground=c["accent"],
                                                 selectforeground=c["btn_text"],
                                                 borderwidth=1, relief="solid",
                                                 font=(MODERN_FONT, 11),
                                                 activestyle="none")
            self._autocomplete_list.place_forget()
            self._autocomplete_list.bind("<<ListboxSelect>>",
                                         lambda e: self._on_cby_autocomplete_select(row_id))
            widget.bind("<KeyRelease>", lambda e: self._update_cby_autocomplete(widget))
            widget.bind("<Return>", lambda e: (self._finish_cby_edit(row_id, widget.get()), self._tab_to_next_cell()))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>",
                        lambda e: self.win.after(150, lambda: self._finish_cby_edit(row_id, widget.get())))
            widget.bind("<Down>", lambda e: self._autocomplete_focus(widget))
            widget.bind("<Tab>", lambda e: (tab_to_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_to_prev(), "break"))
            # No initial dropdown — suggestions appear as the user types
        elif key == "importer":
            # Importer column - suggest TIN based on the CBY in this row
            widget = ctk.CTkEntry(self._editing_window, width=w, height=h,
                                  fg_color=c["input"], border_color=c["accent"],
                                  border_width=2, corner_radius=2, text_color=c["text"],
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            # Look up the CBY for this row and show a hint
            cby_val = self._tree.set(row_id, "cby").strip()
            hint = ""
            if cby_val:
                tin_entry = _get_tin_entry(cby_val)
                if tin_entry:
                    name, tin = tin_entry[0], tin_entry[1]
                    if tin:
                        hint = f"  ({name}: {tin})"
                    else:
                        hint = f"  ({name}: no TIN on file)"
                else:
                    hint = f"  (CBY {cby_val} not in database)"
            widget.bind("<Return>", lambda e: (self._finish_importer_edit(row_id, widget.get(), cby_val), self._tab_to_next_cell()))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>",
                        lambda e: self.win.after(150, lambda: self._finish_importer_edit(row_id, widget.get(), cby_val)))
            widget.bind("<Tab>", lambda e: (tab_to_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_to_prev(), "break"))
        else:
            widget = ctk.CTkEntry(self._editing_window, width=w, height=h,
                                  fg_color=c["input"], border_color=c["accent"],
                                  border_width=2, corner_radius=2, text_color=c["text"],
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            widget.bind("<Return>", lambda e: (self._finish_edit(row_id, key, widget.get()), self._tab_to_next_cell()))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_to_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_to_prev(), "break"))
        self._editing_row_id = row_id
        self._editing_key = key

    def _update_cby_autocomplete(self, entry):
        """Update the CBY autocomplete listbox based on what's typed."""
        if self._autocomplete_list is None:
            return
        typed = entry.get().strip().lower()
        matches = []
        if typed:
            for cby, entry in BUILTIN_TIN_NUMBERS.items():
                name = entry[0] if entry else ""
                tin = entry[1] if len(entry) > 1 else ""
                if typed in cby or typed in name.lower() or typed in f"{cby} {name.lower()}":
                    matches.append((cby, name, tin))
                    if len(matches) >= 20:
                        break
        # Sort by CBY number
        matches.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 0)

        self._autocomplete_list.delete(0, "end")
        for cby, name, tin in matches:
            display = f"#{cby} - {name}"
            if tin:
                display += f" ({tin})"
            self._autocomplete_list.insert("end", display)

        if matches:
            entry_x = self._edit_cell_x
            entry_y = self._edit_cell_y
            entry_h = self._edit_cell_h
            list_h = min(len(matches), 8) * 22
            win_h = self.win.winfo_height()
            if entry_y + entry_h + list_h > win_h:
                place_y = max(0, entry_y - list_h)
            else:
                place_y = entry_y + entry_h
            self._autocomplete_list.configure(height=min(len(matches), 8))
            self._autocomplete_list.place(x=entry_x, y=place_y,
                                          width=400, height=list_h)
            self._autocomplete_list.lift()
        else:
            self._autocomplete_list.place_forget()

    def _on_cby_autocomplete_select(self, row_id):
        """When a CBY suggestion is clicked, fill in CBY + importer number."""
        if self._autocomplete_list is None:
            return
        sel = self._autocomplete_list.curselection()
        if not sel:
            return
        selected = self._autocomplete_list.get(sel[0])
        # Parse "#cby - name (tin)" format
        # Extract CBY number
        cby = selected.split(" - ")[0].lstrip("#").strip()
        self._finish_cby_edit(row_id, cby)

    def _finish_cby_edit(self, row_id, value):
        """Finish editing CBY and auto-fill importer number from TIN database."""
        if self._autocomplete_list is not None:
            self._autocomplete_list.place_forget()
            self._autocomplete_list.destroy()
            self._autocomplete_list = None
        if self._editing_window is not None:
            try:
                self._tree.set(row_id, "cby", value)
            except Exception:
                pass
            # Auto-fill importer number from TIN database
            tin_entry = _get_tin_entry(value)
            if tin_entry:
                name, tin = tin_entry[0], tin_entry[1]
                if tin:
                    try:
                        self._tree.set(row_id, "importer", tin)
                    except Exception:
                        pass
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None

    def _finish_importer_edit(self, row_id, value, cby_val):
        """Finish editing importer number. If empty and CBY has a TIN, auto-fill it."""
        if self._editing_window is not None:
            value = value.strip()
            if not value and cby_val:
                # Field was left empty - auto-fill from TIN database
                tin_entry = _get_tin_entry(cby_val)
                if tin_entry and tin_entry[1]:
                    value = tin_entry[1]
            try:
                self._tree.set(row_id, "importer", value)
            except Exception:
                pass
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None

    def _update_autocomplete(self, entry, row_id):
        """Update the autocomplete listbox based on what's typed in the entry."""
        if self._autocomplete_list is None:
            return
        typed = entry.get().strip().lower()
        matches = []
        if typed:
            for desc, code, unit, proc in BUILTIN_CODES:
                if typed in desc.lower():
                    matches.append((desc, code, unit, proc))
                    if len(matches) >= 20:
                        break
        else:
            # Show first 20 if nothing typed
            for desc, code, unit, proc in BUILTIN_CODES[:20]:
                matches.append((desc, code, unit, proc))

        self._autocomplete_list.delete(0, "end")
        for desc, code, unit, proc in matches:
            self._autocomplete_list.insert("end", desc)

        if matches:
            # Position the listbox below the editing entry, or above if
            # there isn't enough room below (e.g. first row near top)
            entry_x = self._edit_cell_x
            entry_y = self._edit_cell_y
            entry_h = self._edit_cell_h
            list_h = min(len(matches), 8) * 22
            win_h = self.win.winfo_height()
            if entry_y + entry_h + list_h > win_h:
                # Not enough room below — place above the entry
                place_y = max(0, entry_y - list_h)
            else:
                place_y = entry_y + entry_h
            self._autocomplete_list.configure(height=min(len(matches), 8))
            self._autocomplete_list.place(x=entry_x, y=place_y,
                                          width=350, height=list_h)
            self._autocomplete_list.lift()
        else:
            self._autocomplete_list.place_forget()

    def _autocomplete_focus(self, entry):
        """Move focus to the autocomplete listbox."""
        if self._autocomplete_list is not None and self._autocomplete_list.size() > 0:
            self._autocomplete_list.focus_set()
            self._autocomplete_list.selection_set(0)
            self._autocomplete_list.activate(0)

    def _on_autocomplete_select(self, row_id):
        """When a suggestion is clicked, fill in desc + code + unit + proc."""
        if self._autocomplete_list is None:
            return
        sel = self._autocomplete_list.curselection()
        if not sel:
            return
        selected_desc = self._autocomplete_list.get(sel[0])
        self._finish_desc_edit(row_id, selected_desc)

    def _finish_desc_edit(self, row_id, value):
        """Finish editing the description and auto-fill code/unit/proc if matched."""
        if self._autocomplete_list is not None:
            self._autocomplete_list.place_forget()
            self._autocomplete_list.destroy()
            self._autocomplete_list = None
        if self._editing_window is not None:
            try:
                self._tree.set(row_id, "desc", value)
            except Exception:
                pass
            # Look up code/unit/proc from BUILTIN_CODES
            match = _BUILTIN_DESC_MAP.get(clean_text(value))
            if match:
                code, unit, proc = match
                try:
                    self._tree.set(row_id, "code", code)
                    self._tree.set(row_id, "unit", unit)
                    self._tree.set(row_id, "proc", proc)
                except Exception:
                    pass
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None

    def _finish_edit(self, row_id, key, value):
        if self._editing_window is not None:
            try:
                self._tree.set(row_id, key, value)
            except Exception:
                pass
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None
        # Update red tag when description is edited
        if key == "desc":
            if value.strip():
                self._tree.item(row_id, tags=())
            else:
                self._tree.item(row_id, tags=("blank_desc",))

    def _cancel_edit(self):
        if self._autocomplete_list is not None:
            self._autocomplete_list.place_forget()
            self._autocomplete_list.destroy()
            self._autocomplete_list = None
        if self._editing_window is not None:
            try:
                self._editing_window.destroy()
            except Exception:
                pass
            self._editing_window = None
        self._editing_row_id = None
        self._editing_key = None

    def _save_current_edit(self):
        """Save whatever is in the current editing widget without closing the nav flow."""
        if self._editing_window is None or self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        key = self._editing_key
        # Find the widget inside the editing window
        for child in self._editing_window.winfo_children():
            try:
                value = child.get()
                break
            except Exception:
                continue
        else:
            return
        # Save using the appropriate finish method
        if key == "desc":
            self._finish_desc_edit(row_id, value)
        elif key == "cby":
            self._finish_cby_edit(row_id, value)
        elif key == "importer":
            cby_val = self._tree.set(row_id, "cby").strip()
            self._finish_importer_edit(row_id, value, cby_val)
        else:
            self._finish_edit(row_id, key, value)

    def _tab_to_next_cell(self):
        """After saving current cell, move to the next cell to the right."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        key = self._editing_key
        if key not in col_keys:
            return
        cur_idx = col_keys.index(key)
        next_idx = cur_idx + 1
        if next_idx >= len(col_keys):
            # Wrap to next row
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return  # last cell
            except ValueError:
                return
        self._start_edit(row_id, col_keys[next_idx])

    def _tab_to_prev_cell(self):
        """After saving current cell, move to the previous cell to the left."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        key = self._editing_key
        if key not in col_keys:
            return
        cur_idx = col_keys.index(key)
        prev_idx = cur_idx - 1
        if prev_idx < 0:
            # Wrap to previous row, last column
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    prev_idx = len(col_keys) - 1
                else:
                    return
            except ValueError:
                return
        self._start_edit(row_id, col_keys[prev_idx])

    def _remove_selected_row(self):
        selected = self._tree.selection()
        if not selected:
            return
        for item in selected:
            self._tree.delete(item)
        if not self._tree.get_children():
            self._show_placeholder()

    def _remove_row(self, item_id):
        try:
            self._tree.delete(item_id)
        except Exception:
            pass
        if not self._tree.get_children():
            self._show_placeholder()

    def _add_row(self, data=None):
        self._hide_placeholder()
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        values = []
        for key in col_keys:
            val = ""
            if data and data.get(key) not in (None, ""):
                val = str(data[key])
            elif key == "qty":
                val = "1"
            elif key == "origin":
                val = "USA"
            elif key == "proc":
                val = "HOME"
            values.append(val)
        # Check if description is blank — highlight the row red
        tags = ()
        if data and not str(data.get("desc", "")).strip():
            tags = ("blank_desc",)
        item_id = self._tree.insert("", "end", values=values, tags=tags)
        return item_id

    def _clear_all_rows(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._hide_placeholder()

    def _get_all_rows(self):
        col_keys = [k for k, _, _ in self.cfg["col_defs"]]
        rows = []
        for item in self._tree.get_children():
            vals = self._tree.set(item)
            row = {}
            for key in col_keys:
                row[key] = _TreeValue(vals.get(key, ""))
            rows.append(row)
        return rows

    # -- Build from Manifest (reads COLS sheet) ---------------------------
    def _build_from_manifest(self):
        """Open a file picker, then parse the COLS sheet in a background thread
        with a loading popup so the UI doesn't freeze."""
        initial_dir = str(self.cfg["excel_file"].parent) if self.cfg["excel_file"].parent.exists() else str(Path.home())
        file_path = filedialog.askopenfilename(
            title="Select " + self.cfg["mode"].title() + " Manifest Excel File",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not file_path:
            return

        # --- Loading popup (show BEFORE importing pandas to avoid freeze) ---
        loader = ctk.CTkToplevel(self.win)
        loader.title("Building from Manifest")
        loader.configure(fg_color=self.cfg["colors"]["bg"])
        loader.geometry("340x140")
        loader.transient(self.win)
        loader.grab_set()
        sw, sh = loader.winfo_screenwidth(), loader.winfo_screenheight()
        loader.geometry(f"+{int((sw-340)/2)}+{int((sh-140)/2)}")
        loader.resizable(False, False)
        loader.overrideredirect(False)

        c = self.cfg["colors"]
        ctk.CTkLabel(loader, text="Building from Manifest...",
                     font=(MODERN_FONT, 15, "bold"), text_color=c["light"]).pack(pady=(24, 8))
        ctk.CTkLabel(loader, text="Loading, please wait...",
                     font=(MODERN_FONT, 11), text_color=c["text"]).pack(pady=(0, 12))
        progress = ctk.CTkProgressBar(loader, width=280, height=14,
                                      progress_color=c["accent"],
                                      fg_color=c["input"])
        progress.pack(pady=(0, 20))
        progress.set(0.05)
        loader.update()

        # Now import pandas (this is the slow part — popup is already showing)
        try:
            import pandas as pd
        except ImportError:
            loader.destroy()
            messagebox.showerror("Missing Dependency",
                                 "pandas is required.\n\n"
                                 "Install with:\n    pip install pandas openpyxl")
            return

        # --- Parse in background thread ---
        result = {"data": None, "error": None}

        def parse_worker():
            try:
                cfg = self.cfg
                xlsx = Path(file_path)
                sheet = cfg["cols_sheet"]
                cm = cfg["cols_column_map"]
                hl = cfg["cols_header_labels"]
                hdr_row = cfg["cols_header_row"]
                data_start = cfg["cols_data_start_row"]

                # Read COLS sheet - try the configured name, then variations
                # (some manifests have " COLS" with a leading space, others "COLS")
                xl = pd.ExcelFile(xlsx)
                sheet_names = [s.strip() for s in xl.sheet_names]
                sheet_found = None
                for sn in xl.sheet_names:
                    if sn.strip().lower() == sheet.strip().lower():
                        sheet_found = sn
                        break
                if sheet_found is None:
                    # Try case-insensitive match on "cols"
                    for sn in xl.sheet_names:
                        if sn.strip().lower() == "cols":
                            sheet_found = sn
                            break
                if sheet_found is None:
                    raise ValueError(f"Could not find a COLS sheet. Available sheets: {xl.sheet_names}")
                raw_df = pd.read_excel(xlsx, sheet_name=sheet_found, header=None)

                # Parse header overview cells
                hv = {"manifest_date": "", "arrival_date": "", "master_bl": "",
                      "pkg_count": "1", "weight_cf": "0.00", "weight_lb": "0.00",
                      "vessel_code": "", "voyage_no": ""}
                for i in range(min(hdr_row, len(raw_df))):
                    row_vals = [str(v).strip() if not pd.isna(v) else "" for v in raw_df.iloc[i].tolist()]
                    for key, label in hl.items():
                        if label and label in row_vals:
                            col_idx = row_vals.index(label)
                            val_idx = col_idx + 1
                            if val_idx < len(raw_df.columns):
                                val = raw_df.iloc[i].tolist()[val_idx]
                                if pd.isna(val) or str(val).strip().lower() in ["", "nan"]:
                                    continue
                                if key in ("manifest_date", "arrival_date"):
                                    try:
                                        hv[key] = pd.to_datetime(val).strftime("%Y-%m-%d")
                                    except Exception:
                                        hv[key] = str(val).strip()
                                elif key == "pkg_count":
                                    try:
                                        hv[key] = str(int(float(val)))
                                    except Exception:
                                        hv[key] = str(val).strip()
                                elif key in ("weight_cf", "weight_lb"):
                                    hv[key] = money_str(val)
                                else:
                                    hv[key] = str(val).strip()

                # Use built-in codes database (not from manifest)
                desc_to_code = dict(_BUILTIN_DESC_MAP)
                code_to_unit = {code: unit for code, (_, unit, _) in _BUILTIN_DESC_MAP.values() for code in [code]} if False else {}
                code_to_proc = {}
                for _clean_desc, (_code, _unit, _proc) in _BUILTIN_DESC_MAP.items():
                    code_to_unit[_code] = _unit
                    code_to_proc[_code] = _proc
                # Also map by original desc for exact matching
                for _desc, _code, _unit, _proc in BUILTIN_CODES:
                    desc_to_code[clean_text(_desc)] = _code

                def find_best_match(description):
                    cleaned = clean_text(description)
                    if cleaned in desc_to_code:
                        return desc_to_code[cleaned]
                    for k, v in desc_to_code.items():
                        if cleaned in k or k in cleaned:
                            return v
                    words = cleaned.split()
                    if len(words) >= 2:
                        first_two = " ".join(words[:2])
                        for k, v in desc_to_code.items():
                            if first_two in k:
                                return v
                    matches = get_close_matches(cleaned, desc_to_code.keys(), n=1, cutoff=0.60)
                    if matches:
                        return desc_to_code[matches[0]]
                    return DEFAULT_COMMODITY_CODE

                # Build item rows data
                items = []
                for i in range(data_start, len(raw_df)):
                    row_data = raw_df.iloc[i].tolist()

                    cby_val = ""
                    if cm["cby"] is not None and cm["cby"] < len(row_data):
                        cby_val = clean_dock_receipt(row_data[cm["cby"]])
                    if not cby_val or str(cby_val).lower() in ("", "nan", "0", "0.0", "none", "store"):
                        continue

                    desc_str = ""
                    if cm["desc"] is not None and cm["desc"] < len(row_data):
                        desc_str = str(row_data[cm["desc"]]).strip()
                    # Treat 0/0.0/NaN as blank descriptions (empty Excel cells)
                    if desc_str.lower() in ("nan", "0", "0.0", "shipment total"):
                        desc_str = ""
                    # Don't skip rows with blank descriptions — include them
                    # so the user sees the gap and can fix it.

                    dock_val = ""
                    if cm["dock"] is not None and cm["dock"] < len(row_data):
                        dock_val = clean_dock_receipt(row_data[cm["dock"]])

                    importer_number = ""
                    if cm["importer"] is not None and cm["importer"] < len(row_data):
                        raw_imp = row_data[cm["importer"]]
                        try:
                            importer_number = str(int(float(raw_imp)))
                        except Exception:
                            importer_number = str(raw_imp).strip()
                        if importer_number.lower() in ("", "nan", "0", "none"):
                            importer_number = ""
                    # If no importer from manifest, look up by CBY in built-in TIN list
                    if not importer_number and cby_val:
                        tin_entry = _get_tin_entry(str(cby_val))
                        if tin_entry:
                            importer_number = tin_entry[1]

                    qty_val = "1"
                    if cm["qty"] is not None and cm["qty"] < len(row_data):
                        try:
                            qty_val = str(int(float(row_data[cm["qty"]])))
                        except Exception:
                            qty_val = "1"

                    value_val = money_str(row_data[cm["value"]]) if cm["value"] is not None and cm["value"] < len(row_data) else "0.00"
                    freight_val = money_str(row_data[cm["freight"]]) if cm["freight"] is not None and cm["freight"] < len(row_data) else "0.00"
                    insurance_val = money_str(row_data[cm["insurance"]]) if cm["insurance"] is not None and cm["insurance"] < len(row_data) else "0.00"

                    if desc_str and desc_to_code:
                        code = find_best_match(desc_str)
                    else:
                        code = ""  # No code for blank descriptions

                    items.append({
                        "cby": cby_val,
                        "dock": dock_val,
                        "desc": desc_str,
                        "code": code,
                        "unit": code_to_unit.get(code, "NO"),
                        "proc": _apply_special_proc(cby_val, code_to_proc.get(code, "HOME")),
                        "qty": qty_val,
                        "value": value_val,
                        "freight": freight_val,
                        "insurance": insurance_val,
                        "importer": importer_number,
                        "origin": "USA",
                    })

                result["data"] = {"hv": hv, "items": items,
                                   "desc_to_code": desc_to_code,
                                   "code_to_unit": code_to_unit,
                                   "code_to_proc": code_to_proc,
                                   "filename": xlsx.name,
                                   "manifest_dir": xlsx.parent}

            except Exception as e:
                result["error"] = str(e)

        def on_progress(value):
            try:
                progress.set(value)
                loader.update_idletasks()
            except Exception:
                pass

        def run_thread():
            # Animate the progress bar gently while parsing
            import time
            pv = [0.1]
            stop = [False]

            def animate():
                while not stop[0]:
                    pv[0] = min(0.9, pv[0] + 0.03)
                    on_progress(pv[0])
                    time.sleep(0.08)

            anim_thread = threading.Thread(target=animate, daemon=True)
            anim_thread.start()

            parse_worker()
            stop[0] = True
            on_progress(1.0)

            # Schedule UI update on main thread
            self.win.after(100, lambda: finish_ui())

        def finish_ui():
            try:
                loader.destroy()
            except Exception:
                pass

            if result["error"]:
                messagebox.showerror("Load Error", f"Failed to read the manifest:\n\n{result['error']}")
                return

            data = result["data"]
            hv = data["hv"]
            items = data["items"]

            # Store commodity lookup tables
            self._desc_to_code = data["desc_to_code"]
            self._code_to_unit = data["code_to_unit"]
            self._code_to_proc = data["code_to_proc"]
            self._manifest_dir = data.get("manifest_dir")

            # Populate header widgets
            cfg = self.cfg
            self._set_header("awb", hv["master_bl"])
            if hv.get("vessel_code") and "vessel" in self.header_entries:
                self._set_vessel(hv["vessel_code"])
            elif cfg.get("vessel_fixed") and "vessel" in self.header_entries:
                self._set_header("vessel", cfg["vessel_fixed"])
            if hv.get("voyage_no"):
                self._set_header("voyage", hv["voyage_no"])
            if hv.get("manifest_date"):
                self._set_header("departure", hv["manifest_date"])
            if hv.get("arrival_date"):
                self._set_header("arrival", hv["arrival_date"])
            if hv.get("weight_lb"):
                self._set_header("weight_lb", hv["weight_lb"])
            if hv.get("weight_cf"):
                self._set_header("weight_cf", hv["weight_cf"])
            if hv.get("pkg_count"):
                self._set_header("pkg_count", hv["pkg_count"])

            # Populate treeview - instant with native widget
            self._clear_all_rows()
            for item in items:
                self._add_row(item)

            self._set_status(f"Loaded {len(items)} line item(s) from {data['filename']} (COLS sheet).")
            messagebox.showinfo("Loaded", f"Loaded {len(items)} line item(s) from the COLS sheet.\n\n"
                                           "Some header fields may need manual entry.")

        # Start the background thread
        work_thread = threading.Thread(target=run_thread, daemon=True)
        work_thread.start()



    # -- Collect header ---------------------------------------------------
    def _collect_header(self):
        vessel_raw = self.header_entries["vessel"].get().strip() if "vessel" in self.header_entries else ""
        m = re.search(r"\((.*?)\)", vessel_raw)
        vessel_code = m.group(1).strip() if m else vessel_raw
        return {
            "awb": self.header_entries["awb"].get().strip(),
            "vessel": vessel_code,
            "voyage": self.header_entries["voyage"].get().strip(),
            "departure": self.header_entries["departure"].get().strip(),
            "arrival": self.header_entries["arrival"].get().strip(),
            "weight_lb": self.header_entries["weight_lb"].get().strip() or "0.00",
            "weight_cf": self.header_entries["weight_cf"].get().strip() or "0.00",
            "pkg_count": self.header_entries["pkg_count"].get().strip() or "1",
            "master_importer": self.header_entries["master_importer"].get().strip() or MASTER_IMPORTER_DEFAULT,
        }

    # -- Generate XML -----------------------------------------------------
    def _generate_xml(self):
        cfg = self.cfg
        hdr = self._collect_header()
        if not hdr["awb"]:
            messagebox.showerror("Missing AWB", "Please enter the AWB / Master BL # before generating.")
            return

        # Read all rows from the treeview
        all_rows = self._get_all_rows()
        active_rows = []
        for row in all_rows:
            cby = row["cby"].get().strip()
            desc = row["desc"].get().strip()
            if not cby and not desc:
                continue
            if not cby:
                continue
            active_rows.append(row)

        if not active_rows:
            messagebox.showerror("No Line Items", "There are no line items with a CBY to generate.")
            return

        # Check for blank descriptions
        blank_desc_cby = []
        for row in active_rows:
            desc = row["desc"].get().strip()
            cby = row["cby"].get().strip()
            if not desc or desc.lower() in ("0", "0.0", "nan"):
                blank_desc_cby.append(cby)

        if blank_desc_cby:
            # Count occurrences per CBY so we show each once with a count
            counts = Counter(blank_desc_cby)
            cby_list = ", ".join(
                f"{cby} ({n})" if n > 1 else cby
                for cby, n in counts.items())
            # Custom dialog with proper button labels
            dialog = ctk.CTkToplevel(self.win)
            dialog.title("Description Box Blank")
            dialog.configure(fg_color="#1a1a2e")
            dialog.geometry("460x240")
            dialog.resizable(False, False)
            dialog.transient(self.win)
            dialog.grab_set()
            dialog.update_idletasks()
            px = self.win.winfo_rootx() + (self.win.winfo_width() - 460) // 2
            py = self.win.winfo_rooty() + (self.win.winfo_height() - 240) // 2
            dialog.geometry(f"+{px}+{py}")
            ctk.CTkLabel(dialog, text="Description Box Blank",
                         font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(20, 4))
            ctk.CTkLabel(dialog,
                         text=f"The following CBY(s) have blank descriptions:\n{cby_list}",
                         font=(MODERN_FONT, 12), text_color="#ff6b6b",
                         justify="center").pack(pady=(0, 16))
            choice = [None]
            def on_cancel():
                choice[0] = False
                dialog.destroy()
            def on_proceed():
                choice[0] = True
                dialog.destroy()
            btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            btn_frame.pack(pady=(0, 20))
            ctk.CTkButton(btn_frame, text="Cancel Export", command=on_cancel,
                         fg_color="#c0392b", hover_color="#e74c3c",
                         width=140, height=32, corner_radius=6,
                         text_color="#ffffff", font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=(30, 10))
            ctk.CTkButton(btn_frame, text="Proceed - with Issues", command=on_proceed,
                         fg_color="#e8820c", hover_color="#ffa726",
                         width=160, height=32, corner_radius=6,
                         text_color="#ffffff", font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=(10, 30))
            dialog.bind("<Escape>", lambda e: on_cancel())
            dialog.bind("<Return>", lambda e: on_proceed())
            dialog.wait_window()
            if not choice[0]:
                return  # User chose to cancel and fix

        try:
            # If we loaded from a manifest, save XML files next to it.
            # Otherwise fall back to the default output folder.
            if self._manifest_dir:
                # Use the departure/manifest date for the subfolder name
                date_str = hdr.get("departure") or hdr.get("arrival") or "Unknown Date"
                # Clean the date for use as a folder name (replace any invalid chars)
                safe_date = re.sub(r'[<>:"/\|?*]', '-', date_str)
                out = self._manifest_dir / f"XML files for {safe_date}"
            else:
                out = cfg["output_folder"]
            out.mkdir(parents=True, exist_ok=True)
            created = []

            # ---- MASTER ----
            master_wt = hdr["weight_lb"]
            master_vol = hdr["weight_cf"]
            master_root = build_sad_structure(cfg, hdr, "MASTER", hdr["awb"],
                                              "Consolidated Shipment", hdr["pkg_count"],
                                              pkg_type="BG",
                                              gross_wt=master_wt, gross_vol=master_vol)
            mi = ET.Element("Importer")
            ET.SubElement(mi, "Number").text = hdr["master_importer"]
            master_root.insert(2, mi)
            ET.SubElement(master_root, "MoneyDeclaredFlag").text = "N"
            master_fn = f"HBL-Master {hdr['awb']}.xml"
            ET.ElementTree(master_root).write(out / master_fn, encoding="utf-8", xml_declaration=True)
            created.append(master_fn)

            # ---- GROUP BY CBY ----
            groups = {}
            order = []
            for row in active_rows:
                cby = row["cby"].get().strip()
                if cby not in groups:
                    groups[cby] = []
                    order.append(cby)
                groups[cby].append(row)

            skipped = 0
            for cby in order:
                rows = groups[cby]
                try:
                    first_row = rows[0]
                    first_dock = clean_dock_receipt(first_row["dock"].get().strip())
                    importer_number = first_row["importer"].get().strip() or FALLBACK_IMPORTER_NUMBER

                    total_value = Decimal("0.00")
                    total_freight = Decimal("0.00")
                    total_insurance = Decimal("0.00")
                    unique_descs = set()
                    for row in rows:
                        total_value += money(row["value"].get())
                        total_freight += money(row["freight"].get())
                        total_insurance += money(row["insurance"].get())
                        d = row["desc"].get().strip()
                        if d:
                            unique_descs.add(d)

                    if len(unique_descs) == 1:
                        contents = truncate_desc(list(unique_descs)[0]).upper()
                    elif len(unique_descs) > 1:
                        contents = "MIXED GOODS"
                    else:
                        contents = "GOODS"

                    # Air house XMLs use 0.00 for weight/volume; ocean uses header values
                    if cfg["house_weight_zero"]:
                        h_wt, h_vol = "0.00", "0.00"
                    else:
                        h_wt, h_vol = hdr["weight_lb"], hdr["weight_cf"]

                    house = build_sad_structure(cfg, hdr, "HOUSE", first_dock, contents,
                                                len(rows), pkg_type="PK",
                                                gross_wt=h_wt, gross_vol=h_vol)

                    imp_node = ET.Element("Importer")
                    ET.SubElement(imp_node, "Number").text = importer_number
                    house.insert(2, imp_node)

                    valuation = ET.SubElement(house, "Valuation")
                    ET.SubElement(valuation, "Currency").text = "USD"
                    ET.SubElement(valuation, "NetCost").text = money_str(total_value)
                    ET.SubElement(valuation, "NetInsurance").text = money_str(total_insurance)
                    ET.SubElement(valuation, "NetFreight").text = money_str(total_freight)
                    ET.SubElement(valuation, "TermsOfDelivery").text = "FOB"

                    row_idx = 1
                    for row in rows:
                        comm_code = row["code"].get().strip() or DEFAULT_COMMODITY_CODE
                        item_node = ET.SubElement(house, "Items", row=str(row_idx))
                        ET.SubElement(item_node, "Code").text = comm_code
                        ET.SubElement(item_node, "Desc").text = truncate_desc(row["desc"].get())
                        ET.SubElement(item_node, "Origin").text = row["origin"].get().strip() or "USA"
                        ET.SubElement(item_node, "Qty").text = row["qty"].get().strip() or "1"
                        ET.SubElement(item_node, "QtyUnit").text = row["unit"].get().strip() or "NO"
                        ET.SubElement(item_node, "Cost").text = money_str(row["value"].get())
                        ET.SubElement(item_node, "Insurance").text = money_str(row["insurance"].get())
                        ET.SubElement(item_node, "Freight").text = money_str(row["freight"].get())
                        ET.SubElement(item_node, "InvNumber").text = clean_dock_receipt(row["dock"].get().strip())

                        proc_node = ET.SubElement(item_node, "Procedure")
                        _proc_code = row["proc"].get().strip() or "HOME"
                        _proc_code = _apply_special_proc(cby, _proc_code)
                        ET.SubElement(proc_node, "Code").text = _proc_code
                        ET.SubElement(proc_node, "ImporterNumber").text = importer_number
                        row_idx += 1

                    ET.SubElement(house, "MoneyDeclaredFlag").text = "N"
                    h_fn = f"HBL-CBY {cby} {first_dock}.xml"
                    ET.ElementTree(house).write(out / h_fn, encoding="utf-8", xml_declaration=True)
                    created.append(h_fn)
                except Exception as e:
                    skipped += 1
                    print(f"Skipping CBY {cby}: {e}")

            self._set_status(f"Generated {len(created)} XML file(s) into '{out.name}'.")
            summary = (f"Master entries:  1\n"
                       f"House entries:   {len(created) - 1}\n"
                       f"Groups skipped:  {skipped}\n\n"
                       f"XML files saved in:\n{out}")
            messagebox.showinfo("XML Generation Complete", summary)

        except Exception as e:
            messagebox.showerror("Generation Error", f"An error occurred while generating XML:\n\n{e}")

    # -- Navigation -------------------------------------------------------
    def _close_all(self):
        """X button: close everything (console + launcher)."""
        self.win.destroy()
        self.launcher.root.destroy()

    def _go_back(self):
        """Back button: return to the launcher."""
        self.win.destroy()
        self.launcher.show()


# ==============================================================================
# CODES MANAGEMENT WINDOW
# ==============================================================================
class CodesWindow(SupportMixin):
    """Searchable, editable table of commodity codes, units, and procedures.
    Changes are saved back to the script file so they persist across machines
    (via Dropbox sync)."""

    def __init__(self, launcher):
        self.launcher = launcher

        # Support-mixin state
        self._pending_update = None
        self._support_tooltip = "Report a Bug"
        self._tooltip_win = None
        self._support_bg = "#0f1117"
        self._window_name = "Item Codes"
        self._set_support_palette({
            "bg": "#0f1117", "panel": "#141a2a", "input": "#0a0e16",
            "border": "#1a2a4a", "text": "#c8d6e5", "light": "#5a7a9a",
            "accent": "#0f3460", "accent_hover": "#1a4a7a",
        })

        self.win = ctk.CTkToplevel()
        self.win.title("Commodity Codes Management")
        self.win.configure(fg_color="#0f1117")
        _register_window_name(self.win, self._window_name, {"bg": "#0f1117", "panel": "#141a2a", "accent": "#0f3460", "accent_hover": "#1a4a7a", "text": "#c8d6e5"})
        self.win.geometry("900x600")
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry(f"+{int((sw-900)/2)}+{int((sh-600)/2)}")
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)
        self.win.transient(launcher.root)

        # Title bar with Back button
        title_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        title_frame.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkButton(title_frame, text="\u2190 Back",
                      command=self._on_close,
                      fg_color="#333", hover_color="#444",
                      width=80, height=28, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold"),
                      text_color="#e8e8e8").pack(side="left")
        ctk.CTkLabel(title_frame, text="Commodity Codes Database",
                     font=(MODERN_FONT, 18, "bold"), text_color="#e8e8e8").pack(side="left", padx=(16, 0))

        # Subtitle
        ctk.CTkLabel(self.win, text=f"{len(BUILTIN_CODES)} codes loaded  -  double-click a cell to edit  -  changes save to the script file",
                     font=(MODERN_FONT, 11), text_color="#888").pack(pady=(8, 12))

        # Search bar
        search_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(search_frame, text="Search:", font=(MODERN_FONT, 12),
                     text_color="#ccc").pack(side="left", padx=(0, 8))
        self._search_var = ctk.StringVar()
        self._search_var.trace("w", lambda *a: self._filter_tree())
        search_entry = ctk.CTkEntry(search_frame, textvariable=self._search_var,
                                    width=300, height=30, fg_color="#1a1a2e",
                                    border_color="#333", border_width=1, corner_radius=5,
                                    text_color="#e8e8e8", font=(MODERN_FONT, 12))
        search_entry.pack(side="left", padx=(0, 12))

        # Add button
        ctk.CTkButton(search_frame, text="+ Add Code", command=self._add_code,
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=110, height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=4)
        # Delete button
        ctk.CTkButton(search_frame, text="Delete Selected", command=self._delete_selected,
                      fg_color="#c0392b", hover_color="#e74c3c",
                      width=130, height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=4)
        # Save button
        ctk.CTkButton(search_frame, text="Save Changes", command=self._save_to_script,
                      fg_color="#e8820c", hover_color="#ffa726",
                      width=130, height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="right", padx=4)

        # Treeview
        tree_frame = ctk.CTkFrame(self.win, fg_color="#1a1a2e", corner_radius=6)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        cols = ["desc", "code", "unit", "proc"]
        labels = ["Description", "Code", "Unit", "Procedure"]
        widths = [400, 120, 80, 120]
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        for col, label, w in zip(cols, labels, widths):
            self._tree.heading(col, text=label, anchor="w")
            self._tree.column(col, width=w, minwidth=w, anchor="w",
                              stretch=(col == "desc"))

        style = ttk.Style()
        style.configure("Treeview", background="#1a1a2e", foreground="#e8e8e8",
                        fieldbackground="#1a1a2e", bordercolor="#333",
                        rowheight=26, font=(MODERN_FONT, 11))
        style.configure("Treeview.Heading", background="#0f1117", foreground="#ffa726",
                        font=(MODERN_FONT, 11, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#2e8b57")],
                  foreground=[("selected", "#fff")])

        tree_scroll = ctk.CTkScrollbar(tree_frame, command=self._tree.yview)
        self._tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y", padx=(2, 4), pady=4)
        self._tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)

        self._tree.bind("<Double-1>", self._on_double_click)
        self._tree.bind("<Tab>", self._codes_tab_next)
        self._tree.bind("<Shift-Tab>", self._codes_tab_prev)
        self._editing_win = None
        self._editing_row_id = None
        self._editing_key = None
        self._deleted_desc_set = set()
        self._dirty = False

        # Populate
        self._populate_tree()

        # Status (with support icon in the corner)
        status_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        status_frame.pack(fill="x", padx=22, pady=(0, 8))
        self._status = ctk.CTkLabel(status_frame, text="Ready", font=(MODERN_FONT, 11),
                                    text_color="#888", anchor="w")
        self._status.pack(side="left", fill="x", expand=True)
        self._attach_support_icon(status_frame)

        # Check for updates in the background
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    def _populate_tree(self, filter_text=""):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for desc, code, unit, proc in BUILTIN_CODES:
            if filter_text:
                combined = f"{desc} {code} {unit} {proc}".lower()
                if filter_text.lower() not in combined:
                    continue
            self._tree.insert("", "end", values=(desc, code, unit, proc))
        # Clear deletion tracking when full list is reloaded
        if not filter_text:
            self._deleted_desc_set = set()

    def _filter_tree(self):
        self._populate_tree(self._search_var.get().strip())

    def _on_double_click(self, event):
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        region = self._tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        self._add_code(existing_item=row_id)

    def _codes_tab_next(self, event):
        if self._editing_win is not None:
            return
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        cols = ["desc", "code", "unit", "proc"]
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        next_idx = cur_idx + 1
        if next_idx >= len(cols):
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._codes_start_edit(row_id, cols[next_idx])
        return "break"

    def _codes_tab_prev(self, event):
        if self._editing_win is not None:
            return
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        cols = ["desc", "code", "unit", "proc"]
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        prev_idx = cur_idx - 1
        if prev_idx < 0:
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    prev_idx = len(cols) - 1
                else:
                    return
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._codes_start_edit(row_id, cols[prev_idx])
        return "break"

    def _codes_start_edit(self, row_id, key):
        """Open edit popup on a cell. Shared by double-click and Tab nav."""
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        cols = ["desc", "code", "unit", "proc"]
        col_idx = cols.index(key)
        col = f"#{col_idx + 1}"
        bbox = self._tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        cx = self._tree.winfo_rootx() + x
        cy = self._tree.winfo_rooty() + y
        current_val = self._tree.set(row_id, key)
        self._editing_win = tk.Toplevel(self.win)
        self._editing_win.overrideredirect(True)
        self._editing_win.geometry(f"{w}x{h}+{cx}+{cy}")
        self._editing_win.attributes("-topmost", True)
        self._editing_row_id = row_id
        self._editing_key = key
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)

        def tab_next():
            self._codes_save_current_and_tab(1)
        def tab_prev():
            self._codes_save_current_and_tab(-1)

        if key == "unit":
            widget = ctk.CTkComboBox(self._editing_win, values=UNIT_OPTIONS,
                                     width=w, height=h, fg_color="#1a1a2e",
                                     border_color="#333", border_width=1, corner_radius=2,
                                     text_color="#e8e8e8", button_color="#e8820c",
                                     button_hover_color="#ffa726",
                                     dropdown_fg_color="#1a1a2e",
                                     dropdown_text_color="#e8e8e8",
                                     dropdown_hover_color="#2e8b57")
            widget.set(current_val)
            widget.pack(fill="both", expand=True)
            widget.bind("<<ComboboxSelected>>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<FocusOut>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_prev(), "break"))
            widget.focus_set()
        elif key == "proc":
            widget = ctk.CTkComboBox(self._editing_win, values=PROCEDURE_OPTIONS,
                                     width=w, height=h, fg_color="#1a1a2e",
                                     border_color="#333", border_width=1, corner_radius=2,
                                     text_color="#e8e8e8", button_color="#e8820c",
                                     button_hover_color="#ffa726",
                                     dropdown_fg_color="#1a1a2e",
                                     dropdown_text_color="#e8e8e8",
                                     dropdown_hover_color="#2e8b57")
            widget.set(current_val)
            widget.pack(fill="both", expand=True)
            widget.bind("<<ComboboxSelected>>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<FocusOut>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_prev(), "break"))
            widget.focus_set()
        else:
            widget = ctk.CTkEntry(self._editing_win, width=w, height=h,
                                  fg_color="#1a1a2e", border_color="#e8820c",
                                  border_width=2, corner_radius=2, text_color="#e8e8e8",
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            widget.bind("<Return>", lambda e: (self._finish_edit(row_id, key, widget.get()), self._codes_tab_to_cell(1)))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_prev(), "break"))

    def _codes_save_current_and_tab(self, direction):
        """Save current edit and move to next/prev cell."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        key = self._editing_key
        # Get value from the editing widget
        for child in self._editing_win.winfo_children():
            try:
                value = child.get()
                break
            except:
                continue
        else:
            return
        self._finish_edit(row_id, key, value)
        self.win.after(10, lambda: self._codes_tab_to_cell(direction))

    def _codes_tab_to_cell(self, direction):
        """Move to next (1) or prev (-1) cell after saving."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        cols = ["desc", "code", "unit", "proc"]
        key = self._editing_key
        if key not in cols:
            return
        cur_idx = cols.index(key)
        next_idx = cur_idx + direction
        if next_idx >= len(cols):
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return
            except ValueError:
                return
        elif next_idx < 0:
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    next_idx = len(cols) - 1
                else:
                    return
            except ValueError:
                return
        self._codes_start_edit(row_id, cols[next_idx])

    def _finish_edit(self, row_id, key, value):
        if self._editing_win is not None:
            try:
                self._tree.set(row_id, key, value)
            except Exception:
                pass
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
            self._dirty = True
            self._status.configure(text="Edited - click 'Save Changes' to persist")

    def _cancel_edit(self):
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        self._editing_row_id = None
        self._editing_key = None

    def _add_code(self, existing_item=None):
        """Open a popup dialog to enter (or edit) a commodity code.
        If existing_item is provided, the dialog is pre-filled with that
        row's values and saves update the row instead of inserting."""
        is_edit = existing_item is not None
        dlg = ctk.CTkToplevel(self.win)
        dlg.title("Edit Commodity Code" if is_edit else "Add Commodity Code")
        dlg.configure(fg_color="#0f1117")
        dlg.geometry("460x520")
        dlg.transient(self.win)
        dlg.attributes("-topmost", True)
        # Center over the parent window
        self.win.update_idletasks()
        px = self.win.winfo_rootx() + (self.win.winfo_width() - 460) // 2
        py = self.win.winfo_rooty() + (self.win.winfo_height() - 520) // 2
        dlg.geometry(f"+{px}+{py}")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Edit Commodity Code" if is_edit else "Add Commodity Code",
                     font=(MODERN_FONT, 16, "bold"),
                     text_color="#e8e8e8").pack(pady=(16, 4))
        ctk.CTkLabel(dlg, text="Fill in the details and click Save." if is_edit
                     else "Fill in the details and click Add.",
                     font=(MODERN_FONT, 11), text_color="#888").pack(pady=(0, 12))

        form = ctk.CTkFrame(dlg, fg_color="transparent")
        form.pack(fill="x", padx=24, pady=(0, 8))

        def add_row(label_text, widget):
            ctk.CTkLabel(form, text=label_text, anchor="w",
                         font=(MODERN_FONT, 12), text_color="#ccc").pack(fill="x", pady=(8, 2))
            widget.pack(fill="x")

        # Pre-fill values if editing
        cur_desc, cur_code, cur_unit, cur_proc = "", "", "NO", "HOME"
        if is_edit:
            vals = self._tree.set(existing_item)
            cur_desc = vals.get("desc", "")
            cur_code = vals.get("code", "")
            cur_unit = vals.get("unit", "NO") or "NO"
            cur_proc = vals.get("proc", "HOME") or "HOME"

        desc_entry = ctk.CTkEntry(form, height=30, fg_color="#1a1a2e",
                                  border_color="#333", border_width=1, corner_radius=5,
                                  text_color="#e8e8e8", font=(MODERN_FONT, 12))
        desc_entry.insert(0, cur_desc)
        add_row("Description:", desc_entry)

        code_entry = ctk.CTkEntry(form, height=30, fg_color="#1a1a2e",
                                  border_color="#333", border_width=1, corner_radius=5,
                                  text_color="#e8e8e8", font=(MODERN_FONT, 12))
        code_entry.insert(0, cur_code)
        add_row("Code:", code_entry)
        # Numeric-only warning label (hidden by default)
        code_warn = ctk.CTkLabel(form, text="Numbers only please.",
                                 font=(MODERN_FONT, 10), text_color="#e74c3c", anchor="w")
        code_warn.pack_forget()

        def _validate_code(char):
            if char.isdigit() or char == "":
                code_warn.pack_forget()
                return True
            code_warn.pack(fill="x", pady=(2, 0))
            self.win.after(2000, lambda: code_warn.pack_forget())
            return False
        code_entry.configure(validate="key",
                             validatecommand=(self.win.register(_validate_code), "%S"))

        unit_combo = ctk.CTkComboBox(form, values=UNIT_OPTIONS, height=30,
                                     fg_color="#1a1a2e", border_color="#333",
                                     border_width=1, corner_radius=5,
                                     text_color="#e8e8e8", button_color="#e8820c",
                                     button_hover_color="#ffa726",
                                     dropdown_fg_color="#1a1a2e",
                                     dropdown_text_color="#e8e8e8",
                                     dropdown_hover_color="#2e8b57")
        unit_combo.set(cur_unit)
        add_row("Unit:", unit_combo)

        # Item codes only use HOME/BLD MAT/SCHOOL procedures; RETAILER and
        # SPCL ECO ZONE are customer-level (TIN) procedures, not item-level.
        _ITEM_PROC_OPTIONS = [p for p in PROCEDURE_OPTIONS
                              if p not in ("RETAILER", "SPCL ECO ZONE")]
        proc_combo = ctk.CTkComboBox(form, values=_ITEM_PROC_OPTIONS, height=30,
                                     fg_color="#1a1a2e", border_color="#333",
                                     border_width=1, corner_radius=5,
                                     text_color="#e8e8e8", button_color="#e8820c",
                                     button_hover_color="#ffa726",
                                     dropdown_fg_color="#1a1a2e",
                                     dropdown_text_color="#e8e8e8",
                                     dropdown_hover_color="#2e8b57")
        # If the existing proc is RETAILER/SPCL ECO ZONE (legacy data),
        # default to HOME since those aren't valid item-level procedures.
        if cur_proc in ("RETAILER", "SPCL ECO ZONE"):
            cur_proc = "HOME"
        proc_combo.set(cur_proc)
        add_row("Procedure:", proc_combo)

        status_lbl = ctk.CTkLabel(dlg, text="", font=(MODERN_FONT, 11),
                                  text_color="#e74c3c", anchor="w")
        status_lbl.pack(fill="x", padx=28, pady=(4, 0))

        def confirm():
            desc = desc_entry.get().strip()
            code = code_entry.get().strip()
            unit = unit_combo.get().strip() or "NO"
            proc = proc_combo.get().strip() or "HOME"
            if not desc:
                status_lbl.configure(text="Description is required.")
                desc_entry.focus_set()
                return
            if not code:
                status_lbl.configure(text="Code is required.")
                code_entry.focus_set()
                return
            # When editing, skip duplicate check if the code hasn't changed
            if not (is_edit and code.lower() == cur_code.strip().lower()):
                # Check for duplicate code in the tree AND the full in-memory
                # database (the tree may be filtered, hiding the duplicate).
                dup_item = None
                dup_desc_from_mem = None
                for item in self._tree.get_children():
                    if item == existing_item:
                        continue
                    if self._tree.set(item, "code").strip().lower() == code.lower():
                        dup_item = item
                        break
                if dup_item is None:
                    for mem_desc, mem_code, _mem_unit, _mem_proc in BUILTIN_CODES:
                        if mem_code.strip().lower() == code.lower():
                            dup_desc_from_mem = mem_desc
                            break
                if dup_item is not None or dup_desc_from_mem is not None:
                    if dup_item is not None:
                        dup_desc = self._tree.set(dup_item, "desc").strip()
                    else:
                        dup_desc = dup_desc_from_mem
                    if not messagebox.askyesno(
                            "Duplicate Code",
                            f"Code '{code}' already exists for:\n\n"
                            f"  {dup_desc}\n\n"
                            f"Do you want to override it with the new details?",
                            parent=dlg):
                        return  # user chose No - stay on the dialog
                    # Override the duplicate row (in tree if visible, else memory)
                    if dup_item is not None:
                        self._tree.set(dup_item, "desc", desc)
                        self._tree.set(dup_item, "unit", unit)
                        self._tree.set(dup_item, "proc", proc)
                        self._tree.see(dup_item)
                        self._tree.selection_set(dup_item)
                    else:
                        for i, (mem_desc, mem_code, _u, _p) in enumerate(BUILTIN_CODES):
                            if mem_code.strip().lower() == code.lower():
                                BUILTIN_CODES[i] = (desc, code, unit, proc)
                                break
                    self._dirty = True
                    self._status.configure(
                        text=f"Updated '{desc}' - click 'Save Changes' to persist")
                    dlg.grab_release()
                    dlg.destroy()
                    return
            # Save: update existing row or insert new
            if is_edit:
                self._tree.set(existing_item, "desc", desc)
                self._tree.set(existing_item, "code", code)
                self._tree.set(existing_item, "unit", unit)
                self._tree.set(existing_item, "proc", proc)
                self._tree.see(existing_item)
                self._tree.selection_set(existing_item)
            else:
                item_id = self._tree.insert("", "end", values=(desc, code, unit, proc))
                self._tree.see(item_id)
                self._tree.selection_set(item_id)
            self._dirty = True
            self._status.configure(
                text=f"{'Updated' if is_edit else 'Added'} '{desc}' - click 'Save Changes' to persist")
            dlg.grab_release()
            dlg.destroy()

        def cancel():
            dlg.grab_release()
            dlg.destroy()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(8, 16))
        ctk.CTkButton(btn_frame, text="Cancel", command=cancel,
                      fg_color="#333", hover_color="#444",
                      width=100, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold"),
                      text_color="#e8e8e8").pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text="Save" if is_edit else "Add", command=confirm,
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=100, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="right")

        # Enter key confirms, Escape cancels
        dlg.bind("<Return>", lambda e: confirm())
        dlg.bind("<Escape>", lambda e: cancel())
        desc_entry.focus_set()

    def _delete_selected(self):
        selected = self._tree.selection()
        if not selected:
            return
        if not messagebox.askyesno("Confirm", "Delete the selected code(s)?"):
            return
        for item in selected:
            desc = self._tree.set(item, "desc").strip().lower()
            if desc:
                self._deleted_desc_set.add(desc)
            self._tree.delete(item)
        self._dirty = True
        self._status.configure(text="Deleted - click 'Save Changes' to persist")

    def _save_to_script(self):
        """Save the current tree contents back to the BUILTIN_CODES list in the script file.
        Merges with any changes saved by other users since this window was opened."""
        # Start with the full in-memory list (has ALL codes, not just visible ones)
        my_codes = list(BUILTIN_CODES)
        # Override with any visible tree edits
        visible_codes = {}
        for item in self._tree.get_children():
            vals = self._tree.set(item)
            desc = vals.get("desc", "").strip()
            code = vals.get("code", "").strip()
            unit = vals.get("unit", "NO").strip() or "NO"
            proc = vals.get("proc", "HOME").strip() or "HOME"
            if desc and code:
                visible_codes[desc.lower()] = (desc, code, unit, proc)
        # Merge visible edits into the full list
        if visible_codes:
            my_codes_dict = {desc.lower(): (desc, code, unit, proc) for desc, code, unit, proc in my_codes}
            my_codes_dict.update(visible_codes)
            my_codes = list(my_codes_dict.values())

        # Apply explicit deletions
        deleted_set = getattr(self, '_deleted_desc_set', set())
        my_codes = [(desc, code, unit, proc) for desc, code, unit, proc in my_codes
                    if desc.lower() not in deleted_set]

        if not my_codes and not deleted_set:
            messagebox.showerror("Error", "No valid codes to save.")
            return

        # Re-read the script file to get the latest version from disk
        # (someone else may have saved changes via Dropbox sync)
        script_path = Path(__file__).resolve()
        content = script_path.read_text(encoding="utf-8")

        # Parse the current BUILTIN_CODES block from the file on disk
        import re as _re
        disk_codes = []
        pattern = r'BUILTIN_CODES = \[.*?\n\]'
        m = _re.search(pattern, content, flags=_re.DOTALL)
        if m:
            block_text = m.group(0)
            # Extract tuples from the block
            for line in block_text.split("\n"):
                line = line.strip()
                if line.startswith("(") and line.endswith("),"):
                    try:
                        # Parse: ("desc", "code", "unit", "proc"),
                        inner = line[1:-2]  # strip ( and ),
                        parts = inner.split('", "')
                        if len(parts) == 4:
                            desc = parts[0].lstrip('"')
                            code = parts[1]
                            unit = parts[2]
                            proc = parts[3].rstrip('"')
                            disk_codes.append((desc, code, unit, proc))
                    except Exception:
                        continue

        # Merge: disk entries + my entries, apply explicit deletions
        merged = {}
        for desc, code, unit, proc in disk_codes:
            merged[desc.lower()] = (desc, code, unit, proc)
        for desc, code, unit, proc in my_codes:
            merged[desc.lower()] = (desc, code, unit, proc)
        # Remove explicitly deleted entries
        deleted_set = getattr(self, '_deleted_desc_set', set())
        for key in list(merged.keys()):
            if key in deleted_set:
                del merged[key]

        final_codes = list(merged.values())
        # Sort alphabetically by description
        final_codes.sort(key=lambda x: x[0].lower())

        # Build the new BUILTIN_CODES block
        lines = ["BUILTIN_CODES = ["]
        for desc, code, unit, proc in final_codes:
            safe_desc = desc.replace('\\', '\\\\').replace('"', '\\"')
            lines.append(f'    ("{safe_desc}", "{code}", "{unit}", "{proc}"),')
        lines.append("]")
        new_block = "\n".join(lines)

        new_content = _re.sub(pattern, lambda m: new_block, content, count=1, flags=_re.DOTALL)

        if new_content == content:
            messagebox.showwarning("No Change", "No changes to save.")
            return

        try:
            script_path.write_text(new_content, encoding="utf-8")
            # Update the in-memory list
            BUILTIN_CODES.clear()
            BUILTIN_CODES.extend(final_codes)
            # Rebuild lookup maps
            _BUILTIN_DESC_MAP.clear()
            _BUILTIN_CODE_MAP.clear()
            for _desc, _code, _unit, _proc in final_codes:
                _clean = clean_text(_desc)
                _BUILTIN_DESC_MAP[_clean] = (_code, _unit, _proc)
                _BUILTIN_CODE_MAP[_code] = (_desc, _unit, _proc)
            # Clear deletions and refresh the tree
            self._deleted_desc_set = set()
            self._dirty = False
            self._search_var.set("")  # clear filter to show full list
            self._populate_tree()
            self._status.configure(text=f"Saved {len(final_codes)} codes (merged with disk)")
            messagebox.showinfo("Saved", f"Saved {len(final_codes)} codes to the script file.\n\n"
                                          "Changes merged with any updates from other users\n"
                                          "and will sync via Dropbox.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save:\n\n{e}")

    def _on_close(self):
        if self._dirty:
            answer = messagebox.askyesnocancel("Unsaved Changes",
                                                "You have unsaved changes.\n\n"
                                                "Click Yes to save and close.\n"
                                                "Click No to discard and close.\n"
                                                "Click Cancel to go back.")
            if answer is None:
                return  # cancel - don't close
            if answer:
                self._save_to_script()
        self.launcher._codes_win = None
        self.win.destroy()



def _get_tin_entry(cby):
    """Get TIN entry as a 3-tuple (name, tin, special_proc).
    Handles both old 2-tuple and new 3-tuple formats."""
    entry = BUILTIN_TIN_NUMBERS.get(str(cby))
    if entry is None:
        return None
    if len(entry) >= 3:
        return entry
    elif len(entry) == 2:
        return (entry[0], entry[1], "")
    else:
        return (str(entry[0]) if entry else "", "", "")

def _apply_special_proc(cby, current_proc):
    """Apply special procedure override from TIN database.
    RETAILER overrides HOME only (not SCHOOL or BLD MAT).
    SPCL ECO ZONE overrides ALL procedures.
    Returns the (possibly overridden) procedure code."""
    entry = _get_tin_entry(cby)
    if not entry:
        return current_proc
    special = entry[2] if len(entry) > 2 else ""
    if special == "SPCL ECO ZONE":
        return "SPCL ECO ZONE"
    elif special == "RETAILER" and current_proc == "HOME":
        return "RETAILER"
    return current_proc

# ==============================================================================
# TIN NUMBERS MANAGEMENT WINDOW
# ==============================================================================
class TINWindow(SupportMixin):
    """Searchable, editable table of customer TIN/importer numbers.
    Supports Quick Sync from the AoA Log Master Excel file.
    Changes save back to the script file (syncs via Dropbox)."""

    def __init__(self, launcher):
        self.launcher = launcher

        # Support-mixin state
        self._pending_update = None
        self._support_tooltip = "Report a Bug"
        self._tooltip_win = None
        self._support_bg = "#0f1117"
        self._window_name = "TIN Numbers"
        self._set_support_palette({
            "bg": "#0f1117", "panel": "#141a2a", "input": "#0a0e16",
            "border": "#1a2a4a", "text": "#c8d6e5", "light": "#5a7a9a",
            "accent": "#0f3460", "accent_hover": "#1a4a7a",
        })

        self.win = ctk.CTkToplevel()
        self.win.title("TIN Numbers Management")
        self.win.configure(fg_color="#0f1117")
        _register_window_name(self.win, self._window_name, {"bg": "#0f1117", "panel": "#141a2a", "accent": "#0f3460", "accent_hover": "#1a4a7a", "text": "#c8d6e5"})
        self.win.geometry("960x600")
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry(f"+{int((sw-960)/2)}+{int((sh-600)/2)}")
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)
        self.win.transient(launcher.root)

        # Title bar with Back button
        title_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        title_frame.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkButton(title_frame, text="\u2190 Back",
                      command=self._on_close,
                      fg_color="#333", hover_color="#444",
                      width=80, height=28, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold"),
                      text_color="#e8e8e8").pack(side="left")
        ctk.CTkLabel(title_frame, text="TIN Numbers Database",
                     font=(MODERN_FONT, 18, "bold"), text_color="#e8e8e8").pack(side="left", padx=(16, 0))

        # Subtitle
        ctk.CTkLabel(self.win,
                     text=f"{len(BUILTIN_TIN_NUMBERS)} customers loaded  -  double-click a cell to edit  -  Special Procedure requires CBC approval  -  fallback TIN: {FALLBACK_IMPORTER_NUMBER}",
                     font=(MODERN_FONT, 11), text_color="#888").pack(pady=(8, 12))

        # Search + action bar
        search_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(search_frame, text="Search:", font=(MODERN_FONT, 12),
                     text_color="#ccc").pack(side="left", padx=(0, 8))
        self._search_var = ctk.StringVar()
        self._search_var.trace("w", lambda *a: self._filter_tree())
        search_entry = ctk.CTkEntry(search_frame, textvariable=self._search_var,
                                    width=250, height=30, fg_color="#1a1a2e",
                                    border_color="#333", border_width=1, corner_radius=5,
                                    text_color="#e8e8e8", font=(MODERN_FONT, 12))
        search_entry.pack(side="left", padx=(0, 8))

        ctk.CTkButton(search_frame, text="Quick Sync with AoA List",
                      command=self._quick_sync,
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=180, height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=4)
        ctk.CTkButton(search_frame, text="Save Changes",
                      command=self._save_to_script,
                      fg_color="#e8820c", hover_color="#ffa726",
                      width=120, height=30, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="right", padx=4)

        # Treeview
        tree_frame = ctk.CTkFrame(self.win, fg_color="#1a1a2e", corner_radius=6)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        cols = ["cby", "name", "tin", "sproc"]
        labels = ["CBY #", "Customer Name", "TIN #", "Special Procedure"]
        widths = [80, 350, 120, 160]
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        for col, label, w in zip(cols, labels, widths):
            self._tree.heading(col, text=label, anchor="w")
            self._tree.column(col, width=w, minwidth=w, anchor="w",
                              stretch=(col == "name"))

        style = ttk.Style()
        style.configure("Treeview", background="#1a1a2e", foreground="#e8e8e8",
                        fieldbackground="#1a1a2e", bordercolor="#333",
                        rowheight=26, font=(MODERN_FONT, 11))
        style.configure("Treeview.Heading", background="#0f1117", foreground="#ffa726",
                        font=(MODERN_FONT, 11, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#2e8b57")],
                  foreground=[("selected", "#fff")])

        tree_scroll = ctk.CTkScrollbar(tree_frame, command=self._tree.yview)
        self._tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y", padx=(2, 4), pady=4)
        self._tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)

        self._tree.bind("<Double-1>", self._on_double_click)
        self._tree.bind("<Tab>", self._tin_tab_next)
        self._tree.bind("<Shift-Tab>", self._tin_tab_prev)
        self._editing_win = None
        self._editing_row_id = None
        self._editing_key = None
        self._deleted_cby_set = set()
        self._dirty = False

        # Delete button
        del_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        del_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkButton(del_frame, text="Delete Selected",
                      command=self._delete_selected,
                      fg_color="#c0392b", hover_color="#e74c3c",
                      width=130, height=28, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=4)
        ctk.CTkButton(del_frame, text="+ Add Customer",
                      command=self._add_entry,
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=130, height=28, corner_radius=5,
                      font=(MODERN_FONT, 11, "bold")).pack(side="left", padx=4)

        # Status (with support icon in the corner)
        status_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        status_frame.pack(fill="x", padx=22, pady=(0, 8))
        self._status = ctk.CTkLabel(status_frame, text="Ready", font=(MODERN_FONT, 11),
                                    text_color="#888", anchor="w")
        self._status.pack(side="left", fill="x", expand=True)
        self._attach_support_icon(status_frame)

        # Populate
        self._populate_tree()

        # Check for updates in the background
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    def _sync_tree_to_memory(self):
        """Sync current treeview values back to BUILTIN_TIN_NUMBERS in memory.
        This preserves edits made in filtered mode when the tree is repopulated."""
        _DISPLAY_TO_CODE = {"": "", "Retailer": "RETAILER", "Special Economic Zone": "SPCL ECO ZONE"}
        for item in self._tree.get_children():
            vals = self._tree.set(item)
            cby = vals.get("cby", "").strip()
            if not cby:
                continue
            name = vals.get("name", "").strip()
            tin = vals.get("tin", "").strip()
            sproc_display = vals.get("sproc", "").strip()
            sproc = _DISPLAY_TO_CODE.get(sproc_display, sproc_display)
            BUILTIN_TIN_NUMBERS[cby] = (name, tin, sproc)

    def _populate_tree(self, filter_text="", skip_memory_sync=False):
        # Sync any tree edits back to memory before repopulating.
        # skip_memory_sync=True is used after Quick Sync, which already
        # updated BUILTIN_TIN_NUMBERS directly — syncing the stale tree
        # values back would clobber those updates.
        if not skip_memory_sync:
            self._sync_tree_to_memory()
        for item in self._tree.get_children():
            self._tree.delete(item)
        _SPROC_DISPLAY = {"": "", "RETAILER": "Retailer", "SPCL ECO ZONE": "Special Economic Zone"}
        for cby, entry in sorted(BUILTIN_TIN_NUMBERS.items(),
                                        key=lambda x: int(x[0]) if x[0].isdigit() else 0):
            name = entry[0] if entry else ""
            tin = entry[1] if len(entry) > 1 else ""
            sproc = entry[2] if len(entry) > 2 else ""
            sproc_display = _SPROC_DISPLAY.get(sproc, sproc)
            if filter_text:
                combined = f"{cby} {name} {tin} {sproc_display}".lower()
                if filter_text.lower() not in combined:
                    continue
            self._tree.insert("", "end", values=(cby, name, tin, sproc_display))
        # Clear deletion tracking when full list is reloaded
        if not filter_text:
            self._deleted_cby_set = set()

    def _filter_tree(self, skip_memory_sync=False):
        self._populate_tree(self._search_var.get().strip(),
                            skip_memory_sync=skip_memory_sync)

    def _on_double_click(self, event):
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        region = self._tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        self._add_entry(existing_item=row_id)

    def _tin_tab_next(self, event):
        if self._editing_win is not None:
            return
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        cols = ["cby", "name", "tin", "sproc"]
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        next_idx = cur_idx + 1
        if next_idx >= len(cols):
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._tin_start_edit(row_id, cols[next_idx])
        return "break"

    def _tin_tab_prev(self, event):
        if self._editing_win is not None:
            return
        sel = self._tree.selection()
        if not sel:
            return
        row_id = sel[0]
        cols = ["cby", "name", "tin", "sproc"]
        cur_col = self._tree.identify_column(event.x)
        cur_idx = int(cur_col.replace("#", "")) - 1 if cur_col.startswith("#") else 0
        prev_idx = cur_idx - 1
        if prev_idx < 0:
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    prev_idx = len(cols) - 1
                else:
                    return
            except ValueError:
                return
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)
        self._tree.see(row_id)
        self._tin_start_edit(row_id, cols[prev_idx])
        return "break"

    def _tin_start_edit(self, row_id, key):
        """Open edit popup on a cell. Shared by double-click and Tab nav."""
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        cols = ["cby", "name", "tin", "sproc"]
        col_idx = cols.index(key)
        col = f"#{col_idx + 1}"
        bbox = self._tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        cx = self._tree.winfo_rootx() + x
        cy = self._tree.winfo_rooty() + y
        current_val = self._tree.set(row_id, key)
        self._editing_win = tk.Toplevel(self.win)
        self._editing_win.overrideredirect(True)
        self._editing_win.geometry(f"{w}x{h}+{cx}+{cy}")
        self._editing_win.attributes("-topmost", True)
        self._editing_row_id = row_id
        self._editing_key = key
        self._tree.selection_set(row_id)
        self._tree.focus(row_id)

        def tab_next():
            self._tin_save_current_and_tab(1)
        def tab_prev():
            self._tin_save_current_and_tab(-1)

        if key == "sproc":
            # Dropdown for Special Procedure
            _SPROC_OPTIONS = ["", "Retailer", "Special Economic Zone"]
            widget = ctk.CTkComboBox(self._editing_win, values=_SPROC_OPTIONS,
                                     width=w, height=h, fg_color="#1a1a2e",
                                     border_color="#e8820c", border_width=2, corner_radius=2,
                                     text_color="#e8e8e8", font=(MODERN_FONT, 11),
                                     button_color="#e8820c", button_hover_color="#ffa726",
                                     dropdown_fg_color="#1a1a2e",
                                     dropdown_text_color="#e8e8e8",
                                     dropdown_hover_color="#2a2a4e")
            widget.set(current_val)
            widget.pack(fill="both", expand=True)
            widget.focus_set()

            _sproc_resolved = [False]
            def _on_sproc_select(value=None):
                if _sproc_resolved[0]:
                    return
                val = value if value is not None else widget.get()
                # Check if user selected a special procedure that needs confirmation
                if val and val != current_val:
                    _sproc_resolved[0] = True
                    if not self._confirm_special_proc(val):
                        # User cancelled - revert
                        widget.set(current_val)
                        self._cancel_edit()
                        return
                    _sproc_resolved[0] = False  # allow further edits
                self._finish_edit(row_id, key, val)
                _sproc_resolved[0] = True

            widget.bind("<<ComboboxSelected>>", lambda e: _on_sproc_select(widget.get()))
            widget.bind("<Return>", lambda e: _on_sproc_select(widget.get()))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>", lambda e: _on_sproc_select(widget.get()))
            widget.bind("<Tab>", lambda e: (tab_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_prev(), "break"))
        else:
            widget = ctk.CTkEntry(self._editing_win, width=w, height=h,
                                  fg_color="#1a1a2e", border_color="#e8820c",
                                  border_width=2, corner_radius=2, text_color="#e8e8e8",
                                  font=(MODERN_FONT, 11))
            widget.pack(fill="both", expand=True)
            widget.insert(0, current_val)
            widget.focus_set()
            widget.select_range(0, "end")
            widget.bind("<Return>", lambda e: (self._finish_edit(row_id, key, widget.get()), self._tin_tab_to_cell(1)))
            widget.bind("<Escape>", lambda e: self._cancel_edit())
            widget.bind("<FocusOut>", lambda e: self._finish_edit(row_id, key, widget.get()))
            widget.bind("<Tab>", lambda e: (tab_next(), "break"))
            widget.bind("<Shift-Tab>", lambda e: (tab_prev(), "break"))

    def _tin_save_current_and_tab(self, direction):
        """Save current edit and move to next/prev cell."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        key = self._editing_key
        for child in self._editing_win.winfo_children():
            try:
                value = child.get()
                break
            except:
                continue
        else:
            return
        self._finish_edit(row_id, key, value)
        self.win.after(10, lambda: self._tin_tab_to_cell(direction))

    def _tin_tab_to_cell(self, direction):
        """Move to next (1) or prev (-1) cell after saving."""
        if self._editing_row_id is None:
            return
        row_id = self._editing_row_id
        cols = ["cby", "name", "tin", "sproc"]
        key = self._editing_key
        if key not in cols:
            return
        cur_idx = cols.index(key)
        next_idx = cur_idx + direction
        if next_idx >= len(cols):
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx + 1 < len(children):
                    row_id = children[cur_row_idx + 1]
                    next_idx = 0
                else:
                    return
            except ValueError:
                return
        elif next_idx < 0:
            children = self._tree.get_children()
            try:
                cur_row_idx = children.index(row_id)
                if cur_row_idx > 0:
                    row_id = children[cur_row_idx - 1]
                    next_idx = len(cols) - 1
                else:
                    return
            except ValueError:
                return
        self._tin_start_edit(row_id, cols[next_idx])

    def _finish_edit(self, row_id, key, value):
        if self._editing_win is not None:
            try:
                self._tree.set(row_id, key, value)
            except Exception:
                pass
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
            self._dirty = True
            self._status.configure(text="Edited - click 'Save Changes' to persist")

    def _cancel_edit(self):
        if self._editing_win is not None:
            try:
                self._editing_win.destroy()
            except Exception:
                pass
            self._editing_win = None
        self._editing_row_id = None
        self._editing_key = None

    def _confirm_special_proc(self, display_name, parent=None):
        """Show a confirmation popup when assigning a special procedure.
        Returns True if user confirms, False otherwise.
        If parent is provided, the dialog is transient to that widget
        (so it appears on top of an already-grabbed dialog)."""
        _DISPLAY_TO_CODE = {"Retailer": "RETAILER", "Special Economic Zone": "SPCL ECO ZONE"}
        code = _DISPLAY_TO_CODE.get(display_name, display_name)
        owner = parent if parent is not None else self.win
        dialog = ctk.CTkToplevel(owner)
        dialog.title("Special Procedure Confirmation")
        dialog.configure(fg_color="#1a1a2e")
        if display_name == "Retailer":
            dialog.geometry("520x320")
            dw, dh = 520, 320
        else:
            dialog.geometry("480x260")
            dw, dh = 480, 260
        dialog.resizable(False, False)
        dialog.transient(owner)
        dialog.attributes("-topmost", True)
        dialog.grab_set()
        dialog.update_idletasks()
        px = owner.winfo_rootx() + (owner.winfo_width() - dw) // 2
        py = owner.winfo_rooty() + (owner.winfo_height() - dh) // 2
        dialog.geometry(f"+{px}+{py}")
        ctk.CTkLabel(dialog, text="Special Procedure Assignment",
                     font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(20, 4))
        if display_name == "Retailer":
            ctk.CTkLabel(dialog,
                         text=f"You are assigning: {display_name}\n\n"
                              f"Retailer will only work if ONE of the following is true:\n\n"
                              f"1. The customer has RETAILER granted by CBC\n"
                              f"   on their own account (permanent).\n\n"
                              f"2. We have delegated Retailer to the customer,\n"
                              f"   which must be accepted first and expires after 60 days.\n\n"
                              f"Otherwise the declaration will be rejected by COLS.\n"
                              f"Has this been confirmed?",
                         font=(MODERN_FONT, 12), text_color="#ff6b6b",
                         justify="center").pack(pady=(0, 16))
        else:
            ctk.CTkLabel(dialog,
                         text=f"You are assigning: {display_name}\n\n"
                              f"This procedure requires prior approval from CBC.\n"
                              f"The customer must already be granted this permission\n"
                              f"in COLS, otherwise the declaration will be rejected.\n\n"
                              f"Has CBC approved this special procedure for this customer?",
                         font=(MODERN_FONT, 12), text_color="#ff6b6b",
                         justify="center").pack(pady=(0, 16))
        choice = [None]
        def on_confirm():
            choice[0] = True
            dialog.destroy()
        def on_cancel():
            choice[0] = False
            dialog.destroy()
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))
        ctk.CTkButton(btn_frame, text="Cancel - Not Yet Assigned", command=on_cancel,
                     fg_color="#c0392b", hover_color="#e74c3c",
                     width=180, height=32, corner_radius=6,
                     text_color="#ffffff", font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=(20, 10))
        ctk.CTkButton(btn_frame, text="Yes - Already Approved by CBC", command=on_confirm,
                     fg_color="#2e8b57", hover_color="#3cb371",
                     width=210, height=32, corner_radius=6,
                     text_color="#ffffff", font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=(10, 20))
        dialog.bind("<Escape>", lambda e: on_cancel())
        dialog.wait_window()
        return choice[0] is True

    def _add_entry(self, existing_item=None):
        """Open a popup dialog to enter (or edit) a customer.
        If existing_item is provided, the dialog is pre-filled with that
        row's values and saves update the row instead of inserting.
        If the name AND TIN match an existing entry (when adding new),
        the user is asked whether to override it."""
        _SPROC_OPTIONS = ["", "Retailer", "Special Economic Zone"]
        is_edit = existing_item is not None
        dlg = ctk.CTkToplevel(self.win)
        dlg.title("Edit Customer" if is_edit else "Add Customer")
        dlg.configure(fg_color="#0f1117")
        dlg.geometry("460x560")
        dlg.transient(self.win)
        dlg.attributes("-topmost", True)
        # Center over the parent window
        self.win.update_idletasks()
        px = self.win.winfo_rootx() + (self.win.winfo_width() - 460) // 2
        py = self.win.winfo_rooty() + (self.win.winfo_height() - 560) // 2
        dlg.geometry(f"+{px}+{py}")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Edit Customer" if is_edit else "Add Customer",
                     font=(MODERN_FONT, 16, "bold"),
                     text_color="#e8e8e8").pack(pady=(16, 4))
        ctk.CTkLabel(dlg, text="Fill in the details and click Save." if is_edit
                     else "Fill in the details and click Add.",
                     font=(MODERN_FONT, 11), text_color="#888").pack(pady=(0, 12))

        form = ctk.CTkFrame(dlg, fg_color="transparent")
        form.pack(fill="x", padx=24, pady=(0, 8))

        def add_row(label_text, widget):
            ctk.CTkLabel(form, text=label_text, anchor="w",
                         font=(MODERN_FONT, 12), text_color="#ccc").pack(fill="x", pady=(8, 2))
            widget.pack(fill="x")

        # Pre-fill values if editing
        cur_cby, cur_name, cur_tin, cur_sproc = "", "", "", ""
        if is_edit:
            vals = self._tree.set(existing_item)
            cur_cby = vals.get("cby", "")
            cur_name = vals.get("name", "")
            cur_tin = vals.get("tin", "")
            cur_sproc = vals.get("sproc", "")

        cby_entry = ctk.CTkEntry(form, height=30, fg_color="#1a1a2e",
                                 border_color="#333", border_width=1, corner_radius=5,
                                 text_color="#e8e8e8", font=(MODERN_FONT, 12))
        cby_entry.insert(0, cur_cby)
        add_row("CBY #:", cby_entry)

        name_entry = ctk.CTkEntry(form, height=30, fg_color="#1a1a2e",
                                  border_color="#333", border_width=1, corner_radius=5,
                                  text_color="#e8e8e8", font=(MODERN_FONT, 12))
        name_entry.insert(0, cur_name)
        add_row("Customer Name:", name_entry)

        tin_entry = ctk.CTkEntry(form, height=30, fg_color="#1a1a2e",
                                 border_color="#333", border_width=1, corner_radius=5,
                                 text_color="#e8e8e8", font=(MODERN_FONT, 12))
        tin_entry.insert(0, cur_tin)
        add_row("TIN #:", tin_entry)
        # Numeric-only warning label (hidden by default)
        tin_warn = ctk.CTkLabel(form, text="Numbers only please.",
                                font=(MODERN_FONT, 10), text_color="#e74c3c", anchor="w")
        tin_warn.pack_forget()

        def _validate_tin(char):
            if char.isdigit() or char == "":
                tin_warn.pack_forget()
                return True
            tin_warn.pack(fill="x", pady=(2, 0))
            self.win.after(2000, lambda: tin_warn.pack_forget())
            return False
        tin_entry.configure(validate="key",
                            validatecommand=(self.win.register(_validate_tin), "%S"))

        sproc_combo = ctk.CTkComboBox(form, values=_SPROC_OPTIONS, height=30,
                                      fg_color="#1a1a2e", border_color="#333",
                                      border_width=1, corner_radius=5,
                                      text_color="#e8e8e8", button_color="#e8820c",
                                      button_hover_color="#ffa726",
                                      dropdown_fg_color="#1a1a2e",
                                      dropdown_text_color="#e8e8e8",
                                      dropdown_hover_color="#2e8b57")
        sproc_combo.set(cur_sproc)
        add_row("Special Procedure:", sproc_combo)

        status_lbl = ctk.CTkLabel(dlg, text="", font=(MODERN_FONT, 11),
                                  text_color="#e74c3c", anchor="w")
        status_lbl.pack(fill="x", padx=28, pady=(4, 0))

        def confirm():
            cby = cby_entry.get().strip()
            name = name_entry.get().strip()
            tin = tin_entry.get().strip()
            sproc_display = sproc_combo.get().strip()
            if not cby:
                status_lbl.configure(text="CBY # is required.")
                cby_entry.focus_set()
                return
            if not name:
                status_lbl.configure(text="Customer Name is required.")
                name_entry.focus_set()
                return
            # Confirm special procedure assignment if it's Retailer or
            # Special Economic Zone (and has changed from the current value)
            if sproc_display in ("Retailer", "Special Economic Zone"):
                if sproc_display != cur_sproc.strip():
                    if not self._confirm_special_proc(sproc_display, parent=dlg):
                        return  # user cancelled - stay on the dialog
            # When editing, skip duplicate check if name+TIN haven't changed
            unchanged = (is_edit and name.lower() == cur_name.strip().lower()
                         and tin.lower() == cur_tin.strip().lower())
            if not unchanged:
                # Check for duplicate: name AND TIN match an existing entry.
                # Search both the visible tree AND the full in-memory database
                # (the tree may be filtered, hiding the duplicate).
                self._sync_tree_to_memory()
                dup_item = None
                # First check the visible tree
                for item in self._tree.get_children():
                    if item == existing_item:
                        continue
                    ex_name = self._tree.set(item, "name").strip().lower()
                    ex_tin = self._tree.set(item, "tin").strip().lower()
                    if name and name.lower() == ex_name and tin.lower() == ex_tin:
                        dup_item = item
                        break
                # If not found in tree, check the full in-memory database
                dup_cby_from_mem = None
                if dup_item is None:
                    for mem_cby, (mem_name, mem_tin, _mem_sproc) in BUILTIN_TIN_NUMBERS.items():
                        if is_edit and mem_cby == cur_cby.strip():
                            continue
                        if (name and name.lower() == mem_name.strip().lower()
                                and tin.lower() == mem_tin.strip().lower()):
                            dup_cby_from_mem = mem_cby
                            break
                if dup_item is not None or dup_cby_from_mem is not None:
                    if dup_item is not None:
                        ex_cby = self._tree.set(dup_item, "cby").strip()
                    else:
                        ex_cby = dup_cby_from_mem
                    if not messagebox.askyesno(
                            "Duplicate Customer",
                            f"A customer '{name}' with TIN '{tin or '(none)'}' already exists"
                            f" (CBY #{ex_cby}).\n\n"
                            f"Do you want to override it with the new details?",
                            parent=dlg):
                        return  # user chose No - stay on the dialog
                    # Override the duplicate row (in memory if not visible in tree)
                    if dup_item is not None:
                        self._tree.set(dup_item, "cby", cby)
                        self._tree.set(dup_item, "name", name)
                        self._tree.set(dup_item, "tin", tin)
                        self._tree.set(dup_item, "sproc", sproc_display)
                        self._tree.see(dup_item)
                        self._tree.selection_set(dup_item)
                    else:
                        # Entry exists in memory but not in filtered tree -
                        # update memory directly and mark old CBY for deletion
                        self._deleted_cby_set.add(dup_cby_from_mem)
                        BUILTIN_TIN_NUMBERS[cby] = (name, tin,
                                                    {"": "", "Retailer": "RETAILER",
                                                     "Special Economic Zone": "SPCL ECO ZONE"}.get(sproc_display, sproc_display))
                    self._dirty = True
                    self._status.configure(
                        text=f"Updated '{name}' - click 'Save Changes' to persist")
                    dlg.grab_release()
                    dlg.destroy()
                    return
            # Save: update existing row or insert new
            if is_edit:
                self._tree.set(existing_item, "cby", cby)
                self._tree.set(existing_item, "name", name)
                self._tree.set(existing_item, "tin", tin)
                self._tree.set(existing_item, "sproc", sproc_display)
                self._tree.see(existing_item)
                self._tree.selection_set(existing_item)
            else:
                item_id = self._tree.insert("", "end", values=(cby, name, tin, sproc_display))
                self._tree.see(item_id)
                self._tree.selection_set(item_id)
            self._dirty = True
            self._status.configure(
                text=f"{'Updated' if is_edit else 'Added'} '{name}' - click 'Save Changes' to persist")
            dlg.grab_release()
            dlg.destroy()

        def cancel():
            dlg.grab_release()
            dlg.destroy()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(8, 16))
        ctk.CTkButton(btn_frame, text="Cancel", command=cancel,
                      fg_color="#333", hover_color="#444",
                      width=100, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold"),
                      text_color="#e8e8e8").pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text="Save" if is_edit else "Add", command=confirm,
                      fg_color="#2e8b57", hover_color="#3cb371",
                      width=100, height=32, corner_radius=6,
                      font=(MODERN_FONT, 11, "bold")).pack(side="right")

        # Enter key confirms, Escape cancels
        dlg.bind("<Return>", lambda e: confirm())
        dlg.bind("<Escape>", lambda e: cancel())
        cby_entry.focus_set()

    def _delete_selected(self):
        selected = self._tree.selection()
        if not selected:
            return
        if not messagebox.askyesno("Confirm", "Delete the selected entry(s)?"):
            return
        for item in selected:
            cby = self._tree.set(item, "cby").strip()
            if cby:
                self._deleted_cby_set.add(cby)
            self._tree.delete(item)
        self._dirty = True
        self._status.configure(text="Deleted - click 'Save Changes' to persist")

    def _quick_sync(self):
        """Open a file picker to select an AoA Log Master Excel file,
        then update TIN numbers from it with a progress bar."""
        file_path = filedialog.askopenfilename(
            title="Select AoA Log Master Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not file_path:
            return

        # Show progress bar window IMMEDIATELY (before file read)
        prog_win = ctk.CTkToplevel(self.win)
        prog_win.title("Syncing")
        prog_win.configure(fg_color="#0f1117")
        prog_win.geometry("400x120")
        sw, sh = prog_win.winfo_screenwidth(), prog_win.winfo_screenheight()
        prog_win.geometry(f"+{int((sw-400)/2)}+{int((sh-120)/2)}")
        prog_win.transient(self.win)
        prog_win.grab_set()
        prog_win.protocol("WM_DELETE_WINDOW", lambda: None)  # prevent closing

        ctk.CTkLabel(prog_win, text="Syncing with AoA List...",
                     font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(16, 8))
        self._sync_progress = ctk.CTkProgressBar(prog_win, width=340, height=16,
                                                 progress_color="#2e8b57",
                                                 fg_color="#1a1a2e")
        self._sync_progress.set(0)
        self._sync_progress.pack(pady=(0, 8))
        self._sync_status_label = ctk.CTkLabel(prog_win, text="Opening file...",
                                               font=(MODERN_FONT, 11), text_color="#888")
        self._sync_status_label.pack()
        prog_win.update_idletasks()

        # Run everything in background thread (openpyxl file read + sync)
        def do_sync():
            # Use openpyxl directly — it's already imported elsewhere in the
            # app, so there's no heavy import delay (unlike pandas).
            try:
                import openpyxl
            except Exception as e:
                self.win.after(0, lambda: self._sync_error(prog_win, f"Could not load Excel engine:\n\n{e}"))
                return

            self.win.after(0, lambda: self._update_sync_progress(0, "Reading file..."))

            # Open the workbook (read-only, data_only for cached values)
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            except Exception as e:
                self.win.after(0, lambda: self._sync_error(prog_win, f"Could not open the file:\n\n{e}"))
                return

            # Check for the AOA Log sheet (case-insensitive, space-insensitive)
            sheet_found = None
            for sn in wb.sheetnames:
                if sn.strip().lower() == "aoa log":
                    sheet_found = sn
                    break
            if sheet_found is None:
                self.win.after(0, lambda: self._sync_error(prog_win,
                    f"This does not appear to be an AoA Log Master file.\n\n"
                    f"It must contain a sheet named 'AOA Log' with columns:\n"
                    f"  CBY#, Full Name, Tin #\n\n"
                    f"Sheets found in this file: {wb.sheetnames}"))
                wb.close()
                return

            ws = wb[sheet_found]

            # Read all rows into a list (read_only mode requires iteration)
            all_rows = list(ws.iter_rows(values_only=True))
            wb.close()

            if len(all_rows) < 1:
                self.win.after(0, lambda: self._sync_error(prog_win, "The AOA Log sheet is empty."))
                return

            # Validate the header row looks right
            header = [str(v).strip().lower() if v is not None else '' for v in all_rows[0]]
            if 'cby' not in header[0].lower() or 'name' not in header[1].lower():
                self.win.after(0, lambda: self._sync_error(prog_win,
                    f"The AOA Log sheet doesn't have the expected columns.\n\n"
                    f"Expected: CBY#, Full Name, Tin #\n"
                    f"Found: {header[:4]}"))
                return

            # Process rows
            added = 0
            updated = 0
            unchanged = 0
            total = len(all_rows) - 1
            for i in range(1, len(all_rows)):
                row = all_rows[i]
                cby = str(row[0]).strip() if row[0] is not None else ''
                name = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
                tin = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ''
                if not cby or cby.lower() == 'nan':
                    continue
                if not name or name.lower() in ('nan', '', 'none', '#n/a', '#n/a!', '#ref!', '#value!'):
                    continue
                try:
                    cby_int = str(int(float(cby)))
                except:
                    cby_int = cby
                if not tin or tin.lower() in ('nan', '', '0', 'none'):
                    tin_int = ''
                else:
                    try:
                        tin_int = str(int(float(tin)))
                    except:
                        # TIN contains letters (e.g. "rejected", "pending approval")
                        # — ignore it so the fallback TIN is used instead
                        tin_int = ''

                existing = _get_tin_entry(cby_int)
                if existing is None:
                    BUILTIN_TIN_NUMBERS[cby_int] = (name, tin_int, "")
                    added += 1
                elif tin_int and existing[1] != tin_int:
                    new_name = name if name and name.lower() != 'nan' else existing[0]
                    sproc = existing[2] if len(existing) > 2 else ""
                    BUILTIN_TIN_NUMBERS[cby_int] = (new_name, tin_int, sproc)
                    updated += 1
                elif not existing[0] and name and name.lower() != 'nan':
                    sproc = existing[2] if len(existing) > 2 else ""
                    BUILTIN_TIN_NUMBERS[cby_int] = (name, existing[1], sproc)
                    updated += 1
                else:
                    unchanged += 1

                # Update progress every 50 rows
                if i % 50 == 0 or i == total:
                    progress = i / total
                    self.win.after(0, lambda p=progress, r=i, t=total:
                                   self._update_sync_progress(p, f"Processing row {r} of {t}..."))

            # Done - update UI from main thread
            self.win.after(0, lambda: self._finish_sync(prog_win, added, updated, unchanged))

        threading.Thread(target=do_sync, daemon=True).start()

    def _sync_error(self, prog_win, message):
        """Close progress window and show error (called from main thread)."""
        prog_win.destroy()
        self._sync_progress = None
        self._sync_status_label = None
        messagebox.showerror("Sync Error", message)

    def _update_sync_progress(self, progress, status_text):
        """Update the sync progress bar (called from main thread via .after)."""
        if hasattr(self, '_sync_progress') and self._sync_progress:
            self._sync_progress.set(progress)
        if hasattr(self, '_sync_status_label') and self._sync_status_label:
            self._sync_status_label.configure(text=status_text)

    def _finish_sync(self, prog_win, added, updated, unchanged):
        """Called when sync is complete - close progress bar and show results."""
        prog_win.destroy()
        self._sync_progress = None
        self._sync_status_label = None
        if added > 0 or updated > 0:
            self._dirty = True
        self._filter_tree(skip_memory_sync=True)
        self._status.configure(text=f"Sync complete: {added} new, {updated} updated, {unchanged} unchanged")
        messagebox.showinfo("Sync Complete",
                            f"Added: {added} new entries\n"
                            f"Updated: {updated} entries\n"
                            f"Unchanged: {unchanged} entries\n\n"
                            f"Click 'Save Changes' to persist to the script file.")

    def _save_to_script(self):
        """Save the current tree contents back to BUILTIN_TIN_NUMBERS in the script file.
        Merges with any changes saved by other users since this window was opened.
        Works correctly even when the tree is filtered."""
        # Start with the full in-memory dict (has ALL entries, not just visible ones)
        my_data = dict(BUILTIN_TIN_NUMBERS)
        # Override with any visible tree edits (the tree may be filtered, so only
        # override what's actually shown - those are the ones the user may have edited)
        for item in self._tree.get_children():
            vals = self._tree.set(item)
            cby = vals.get("cby", "").strip()
            name = vals.get("name", "").strip()
            tin = vals.get("tin", "").strip()
            sproc_display = vals.get("sproc", "").strip()
            _DISPLAY_TO_CODE = {"": "", "Retailer": "RETAILER", "Special Economic Zone": "SPCL ECO ZONE"}
            sproc = _DISPLAY_TO_CODE.get(sproc_display, sproc_display)
            if cby:
                my_data[cby] = (name, tin, sproc)

        # Apply explicit deletions
        deleted_set = getattr(self, '_deleted_cby_set', set())
        for cby in deleted_set:
            my_data.pop(cby, None)

        if not my_data and not deleted_set:
            messagebox.showerror("Error", "No valid entries to save.")
            return

        # Re-read the script file to get the latest version from disk
        script_path = Path(__file__).resolve()
        content = script_path.read_text(encoding="utf-8")

        # Parse the current BUILTIN_TIN_NUMBERS block from the file on disk
        import re as _re
        disk_data = {}
        pattern = r'BUILTIN_TIN_NUMBERS = \{.*?\n\}'
        m = _re.search(pattern, content, flags=_re.DOTALL)
        if m:
            block_text = m.group(0)
            for line in block_text.split("\n"):
                line = line.strip()
                if line.startswith('"') and ": (" in line:
                    try:
                        cby_part, rest = line.split('": (', 1)
                        cby = cby_part.lstrip('"')
                        rest = rest.rstrip(',').rstrip(')').rstrip(')')
                        inner = rest.strip()
                        if inner.startswith('('):
                            inner = inner[1:]
                        parts = inner.split('", "')
                        if len(parts) == 2:
                            name = parts[0].lstrip('"')
                            tin = parts[1].rstrip('"')
                            disk_data[cby] = (name, tin, "")
                        elif len(parts) == 3:
                            name = parts[0].lstrip('"')
                            tin = parts[1]
                            sproc = parts[2].rstrip('"')
                            disk_data[cby] = (name, tin, sproc)
                    except Exception:
                        continue

        # Merge: disk entries that aren't in my_data and weren't deleted = added by someone else
        merged = dict(disk_data)
        merged.update(my_data)  # my edits override disk
        # Remove explicitly deleted entries
        for cby in deleted_set:
            merged.pop(cby, None)

        # Build the new BUILTIN_TIN_NUMBERS block
        lines = ["BUILTIN_TIN_NUMBERS = {"]
        for cby in sorted(merged.keys(), key=lambda x: int(x) if x.isdigit() else 0):
            entry = merged[cby]
            name = entry[0] if entry else ""
            tin = entry[1] if len(entry) > 1 else ""
            sproc = entry[2] if len(entry) > 2 else ""
            safe_name = name.replace('\\', '\\\\').replace('"', '\\"')
            lines.append(f'    "{cby}": ("{safe_name}", "{tin}", "{sproc}"),')
        lines.append("}")
        new_block = "\n".join(lines)

        new_content = _re.sub(pattern, lambda m: new_block, content, count=1, flags=_re.DOTALL)

        if new_content == content:
            messagebox.showwarning("No Change", "No changes to save.")
            return

        try:
            script_path.write_text(new_content, encoding="utf-8")
            # Update the in-memory dict
            BUILTIN_TIN_NUMBERS.clear()
            BUILTIN_TIN_NUMBERS.update(merged)
            # Clear deletions and refresh the tree
            self._deleted_cby_set = set()
            self._dirty = False
            self._search_var.set("")  # clear filter to show full list
            self._populate_tree()
            self._status.configure(text=f"Saved {len(merged)} entries (merged with disk)")
            messagebox.showinfo("Saved", f"Saved {len(merged)} TIN entries to the script file.\n\n"
                                          "Changes merged with any updates from other users\n"
                                          "and will sync via Dropbox.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save:\n\n{e}")

    def _on_close(self):
        if self._dirty:
            answer = messagebox.askyesnocancel("Unsaved Changes",
                                                "You have unsaved changes.\n\n"
                                                "Click Yes to save and close.\n"
                                                "Click No to discard and close.\n"
                                                "Click Cancel to go back.")
            if answer is None:
                return  # cancel - don't close
            if answer:
                self._save_to_script()
        self.launcher._tin_win = None
        self.win.destroy()


# ==============================================================================
# UPLOAD DECLARATIONS TO COLS WINDOW
# ==============================================================================
class AutoPilotProgressBar:
    """A small always-on-top floating progress window for auto-pilot mode.
    Has a title bar with close/minimize, is resizable, and matches the
    console UI. If closed or minimized, it pops back up when the session
    completes. Persists on screen until manually closed after completion.
    Cross-platform (Win/Mac)."""
    def __init__(self, parent_win=None):
        self._parent = parent_win
        self._win = ctk.CTkToplevel()
        self._win.title("Auto-Pilot Progress")
        self._win.configure(fg_color="#0f0f1a")
        # Keep on top but allow normal window controls
        try:
            self._win.attributes("-topmost", True)
        except Exception:
            pass

        self._error_flash = False
        self._issue_count = 0
        self._total = 0
        self._current = 0
        self._phase = ""
        self._completed = False  # track if session is done
        self._user_closed = False  # track if user manually closed

        # Size: matches console proportions, resizable
        self._w = 340
        self._h = 120
        self._min_w = 280
        self._min_h = 100
        sw = self._win.winfo_screenwidth()
        sh = self._win.winfo_screenheight()
        x = sw - self._w - 30
        y = sh - self._h - 80
        self._win.geometry(f"{self._w}x{self._h}+{x}+{y}")
        self._win.minsize(self._min_w, self._min_h)
        self._win.resizable(True, True)

        # Handle X button — mark as user-closed, don't destroy yet
        # (we'll destroy in close() or reopen on complete)
        self._win.protocol("WM_DELETE_WINDOW", self._on_user_close)

        # --- Phase label ---
        self._phase_label = ctk.CTkLabel(
            self._win, text="Uploading House Declarations...",  # default phase
            font=(MODERN_FONT, 13, "bold"), text_color="#e8e8e8",
            anchor="w")
        self._phase_label.pack(fill="x", padx=12, pady=(8, 2))

        # --- Progress bar (tkinter Canvas for custom colors) ---
        self._bar_canvas = tk.Canvas(
            self._win, height=16,
            bg="#0f0f1a", highlightthickness=1,
            highlightbackground="#333333")
        self._bar_canvas.pack(fill="x", padx=12, pady=(0, 2))

        # --- Detail label ---
        self._detail_label = ctk.CTkLabel(
            self._win, text="",
            font=(MODERN_FONT, 10), text_color="#888888",
            anchor="w")
        self._detail_label.pack(fill="x", padx=12, pady=(0, 2))

        # --- Activity log (real-time status from console) ---
        self._activity_label = ctk.CTkLabel(
            self._win, text="",
            font=(MODERN_FONT, 9), text_color="#666666",
            anchor="w", wraplength=300)
        self._activity_label.pack(fill="x", padx=12, pady=(0, 6))

        self._win.update_idletasks()

        # Flash animation state
        self._flash_on = False
        self._flash_after_id = None

    def _on_user_close(self):
        """User clicked X. If session not complete, just withdraw (hide).
        If complete, actually destroy."""
        if self._completed:
            self._user_closed = True
            self._win.destroy()
        else:
            # Session still running — hide but keep alive
            # It will pop back up when complete() is called
            self._win.withdraw()

    def set_phase(self, phase, detail=""):
        """Update the phase label and detail text."""
        self._phase = phase
        self._phase_label.configure(text=phase)
        if detail:
            self._detail_label.configure(text=detail)

    def set_activity(self, text):
        """Update the real-time activity line (mirrors console status bar)."""
        # Truncate if too long to fit
        if len(text) > 80:
            text = text[:77] + "..."
        self._activity_label.configure(text=text)

    def set_progress(self, current, total):
        """Update the progress bar. current/total determines fill."""
        self._current = current
        self._total = total
        self._draw_bar()
        if total > 0:
            pct = int(current / total * 100)
            self._detail_label.configure(
                text=f"{current} / {total}  ({pct}%)  |  Issues: {self._issue_count}")

    def add_issue(self):
        """Increment the issue count and flash red."""
        self._issue_count += 1
        self._error_flash = True
        self._start_flash()
        self._detail_label.configure(
            text=f"{self._current} / {self._total}  |  Issues: {self._issue_count}")

    def _start_flash(self):
        """Flash the bar red briefly on error."""
        if self._flash_after_id:
            try:
                self._win.after_cancel(self._flash_after_id)
            except Exception:
                pass
        self._flash_on = True
        self._draw_bar()
        self._flash_after_id = self._win.after(600, self._stop_flash)

    def _stop_flash(self):
        self._flash_on = False
        self._flash_after_id = None
        self._draw_bar()

    def _draw_bar(self):
        """Draw the progress bar on the canvas."""
        self._bar_canvas.delete("all")
        w = self._bar_canvas.winfo_width()
        h = self._bar_canvas.winfo_height()
        if w <= 1:
            w = self._w - 32
        if h <= 1:
            h = 16

        # Background (empty bar)
        self._bar_canvas.create_rectangle(0, 0, w, h, fill="#1a1a2e", outline="")

        # Fill
        if self._total > 0:
            fill_w = int(w * (self._current / self._total))
        else:
            fill_w = 0

        if self._flash_on:
            color = "#b71c1c"  # red flash
        elif "declaration session complete" in self._phase.lower():
            color = "#1b5e20"  # green - complete
        elif "attach" in self._phase.lower():
            color = "#e8820c"  # orange - attaching docs
        elif "upload" in self._phase.lower() or "retry" in self._phase.lower():
            color = "#1a237e"  # blue - uploading
        else:
            color = "#e8820c"  # default orange

        if fill_w > 0:
            self._bar_canvas.create_rectangle(0, 0, fill_w, h, fill=color, outline="")

    def _play_done_sound(self):
        """Play a notification sound when session is complete.
        Cross-platform: uses system bell + platform-specific sound."""
        try:
            self._win.bell()
        except Exception:
            pass
        try:
            if sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["afplay", "/System/Library/Sounds/Glass.aiff"],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            elif sys.platform == "win32":
                import winsound
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            pass

    def complete(self, uploaded, errors, missed, docs_attached, docs_issues):
        """Show the final complete state with green bar and summary.
        If the window was minimized or closed by the user, pop it back up."""
        self._completed = True
        total_issues = errors + missed + docs_issues
        self._phase = "Declaration Session Complete"
        self._phase_label.configure(
            text="Declaration Session Complete",
            text_color="#4caf50")
        self._current = self._total if self._total > 0 else 1
        self._error_flash = False
        self._flash_on = False
        self._draw_bar()
        detail = (f"Uploaded: {uploaded}  |  Docs: {docs_attached}  |  "
                  f"Issues: {total_issues}")
        if total_issues > 0:
            detail += f"  ({errors} errors, {missed} missed, {docs_issues} docs)"
            self._detail_label.configure(text=detail, text_color="#ff6b6b")
        else:
            self._detail_label.configure(text=detail, text_color="#4caf50")
        self._activity_label.configure(text="Declarations complete. Review COLS and submit when ready.")

        # If user minimized or closed the window, pop it back up
        if not self._user_closed:
            try:
                self._win.deiconify()  # un-minimize / un-hide
                self._win.lift()
                try:
                    self._win.attributes("-topmost", True)
                except Exception:
                    pass
                # Flash the title bar to get attention (Windows)
                if sys.platform == "win32":
                    try:
                        self._win.attributes("-topmost", False)
                        self._win.after(100, lambda: self._win.attributes("-topmost", True))
                    except Exception:
                        pass
            except Exception:
                pass

        # Play notification sound
        self._play_done_sound()

    def close(self):
        """Destroy the floating bar. Called by the main script on shutdown."""
        if self._flash_after_id:
            try:
                self._win.after_cancel(self._flash_after_id)
            except Exception:
                pass
        try:
            self._win.destroy()
        except Exception:
            pass

    def is_alive(self):
        """Check if the window still exists and hasn't been user-closed."""
        if self._user_closed:
            return False
        try:
            self._win.winfo_exists()
            return True
        except Exception:
            return False

class UploadWindow(SupportMixin):
    """Window for uploading XML declaration files to the COLS website via Selenium."""

    UPLOAD_URL = "https://online.gov.ky/cols/faces/uploaddeclaration"
    SHORT_DELAY = 1.0
    UPLOAD_DELAY = 2.0
    WAIT_TIMEOUT = 15

    # Status colors for treeview rows
    STATUS_COLORS = {
        "pending":   ("#1a1a2e", "#e8e8e8"),    # dark bg, light text
        "uploading": ("#1a237e", "#ffffff"),    # blue
        "uploaded":  ("#1b5e20", "#ffffff"),    # green
        "error":     ("#b71c1c", "#ffffff"),    # red - failed validation
        "missed":    ("#6a1b9a", "#ffffff"),    # purple - skipped/missed by worker
        "skipped":   ("#424242", "#999999"),    # grey
        "docs_attach": ("#1a237e", "#ffffff"),   # blue - docs being attached
        "docs_ok":     ("#1b5e20", "#ffffff"),   # green - docs attached successfully
        "docs_manual": ("#b71c1c", "#ffffff"),   # red - docs need manual attention
    }

    def __init__(self, launcher):
        self.launcher = launcher

        # Support-mixin state
        self._pending_update = None
        self._support_tooltip = "Report a Bug"
        self._tooltip_win = None
        self._support_bg = "#0f0f1a"
        self._window_name = "Upload Declarations"
        self._set_support_palette({
            "bg": "#0f0f1a", "panel": "#1a1520", "input": "#0a0a12",
            "border": "#e8820c", "text": "#e8e8e8", "light": "#a0a0a0",
            "accent": "#e8820c", "accent_hover": "#f0a030",
        })

        self.win = ctk.CTkToplevel()
        self.win.title("Upload Declarations to COLS")
        self.win.configure(fg_color="#0f0f1a")
        _register_window_name(self.win, self._window_name, {"bg": "#0f0f1a", "panel": "#1a1520", "accent": "#6c3483", "accent_hover": "#8e44ad", "text": "#e8e8e8"})
        w, h = 820, 610
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry(f"{w}x{h}+{int((sw-w)/2)}+{int((sh-h)/2)}")

        self._xml_folder = None
        self._xml_files = []          # list of Path objects
        self._master_files = []       # HBL-Master*.xml files
        self._house_files = []        # HBL-CBY*.xml files
        self._driver = None           # Selenium driver
        self._upload_thread = None
        self._paused = False
        self._stop_requested = False
        self._uploading = False
        self._decl_numbers = {}       # filename -> declaration number
        self._master_decl = ""        # declaration number assigned to master
        self._edge_launched = False   # True after Edge is open and user is logged in
        self._edge_ready = False      # True after Edge browser is open

        # --- Title bar ---
        title_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        title_frame.pack(fill="x", padx=16, pady=(12, 8))
        ctk.CTkButton(title_frame, text="\u2190 Back to Menu",
                      command=self._on_close,
                      fg_color="#0f3460", hover_color="#1a4a7a",
                      text_color="#e8e8e8",
                      font=(MODERN_FONT, 12, "bold"), width=120, height=30,
                      corner_radius=6).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(title_frame, text="Upload Declarations to COLS",
                     font=(MODERN_FONT, 16, "bold"), text_color="#e8e8e8").pack(side="left")

        # Logo (right)
        try:
            _img = Image.open(io.BytesIO(base64.b64decode(MBE_LOGO_B64)))
            _lw, _lh = _img.size
            _tw = 100
            _th = max(1, int(_lh * _tw / _lw))
            _logo = ctk.CTkImage(light_image=_img, dark_image=_img, size=(_tw, _th))
            ctk.CTkLabel(title_frame, image=_logo, text="").pack(side="right")
        except Exception:
            pass

        # --- Folder picker row ---
        folder_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        folder_frame.pack(fill="x", padx=16, pady=(0, 8))
        self._folder_btn = ctk.CTkButton(folder_frame, text="Choose XML Folder",
                                         command=self._pick_folder,
                                         fg_color="#e8820c", hover_color="#ffa726",
                                         text_color="#e8e8e8", font=(MODERN_FONT, 12, "bold"),
                                         width=150, height=30, corner_radius=6)
        self._folder_btn.pack(side="left", padx=(0, 8))
        self._folder_label = ctk.CTkLabel(folder_frame, text="No folder selected",
                                          font=(MODERN_FONT, 11), text_color="#888888")
        self._folder_label.pack(side="left")

        # --- Automation checkboxes ---
        auto_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        auto_frame.pack(fill="x", padx=16, pady=(0, 6))

        self._auto_retry_var = ctk.IntVar(value=0)
        self._auto_retry_cb = ctk.CTkCheckBox(auto_frame, text="Auto-retry missed",
                                              variable=self._auto_retry_var,
                                              font=(MODERN_FONT, 11), text_color="#e8e8e8",
                                              fg_color="#e8820c", hover_color="#ffa726",
                                              height=20)
        self._auto_retry_cb.pack(side="left", padx=(0, 20))

        self._auto_attach_var = ctk.IntVar(value=0)
        self._auto_attach_cb = ctk.CTkCheckBox(auto_frame, text="Auto-attach supporting docs",
                                               variable=self._auto_attach_var,
                                               font=(MODERN_FONT, 11), text_color="#e8e8e8",
                                               fg_color="#e8820c", hover_color="#ffa726",
                                               height=20)
        self._auto_attach_cb.pack(side="left")
        self._autopilot_bar = None  # floating progress bar for auto-pilot mode

        # --- Treeview ---
        tree_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        tree_style = ttk.Style()
        tree_style.theme_use("clam")
        # Configure both the named style AND the base style to ensure dark theme
        tree_style.configure("Upload.Treeview", background="#1a1a2e", foreground="#e8e8e8",
                             fieldbackground="#1a1a2e", borderwidth=0, rowheight=26,
                             font=(MODERN_FONT, 10))
        tree_style.configure("Upload.Treeview.Heading", background="#0f3460",
                             foreground="#e8e8e8", font=(MODERN_FONT, 10, "bold"), relief="flat")
        tree_style.map("Upload.Treeview",
                       background=[("selected", "#e8820c")],
                       foreground=[("selected", "#ffffff")])
        # Also set the base Treeview style in case it hasn't been set yet
        tree_style.configure("Treeview", background="#1a1a2e", foreground="#e8e8e8",
                             fieldbackground="#1a1a2e", bordercolor="#333",
                             rowheight=26, font=(MODERN_FONT, 10))
        tree_style.configure("Treeview.Heading", background="#0f1117", foreground="#ffa726",
                             font=(MODERN_FONT, 10, "bold"), relief="flat")
        tree_style.map("Treeview",
                       background=[("selected", "#e8820c")],
                       foreground=[("selected", "#ffffff")])

        tree_cols = ("file", "cby", "status", "docs", "decl")
        self._tree = ttk.Treeview(tree_frame, columns=tree_cols, show="headings",
                                  style="Upload.Treeview", selectmode="extended")
        self._tree.heading("file", text="XML File", anchor="w")
        self._tree.heading("cby", text="CBY", anchor="w")
        self._tree.heading("status", text="Status", anchor="w")
        self._tree.heading("docs", text="Docs", anchor="w")
        self._tree.heading("decl", text="Declaration #", anchor="w")
        self._tree.column("file", width=250, minwidth=180, anchor="w")
        self._tree.column("cby", width=70, minwidth=50, anchor="w")
        self._tree.column("status", width=100, minwidth=70, anchor="w")
        self._tree.column("docs", width=110, minwidth=80, anchor="w")
        self._tree.column("decl", width=120, minwidth=90, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Allow double-click to manually edit declaration number
        self._tree.bind("<Double-1>", self._on_tree_double_click)
        self._tree.bind("<Delete>", lambda e: self._remove_selected())
        self._decl_edit_win = None

        # Tag configurations for row colors
        for status, (bg, fg) in self.STATUS_COLORS.items():
            self._tree.tag_configure(status, background=bg, foreground=fg)

        # --- Control buttons ---
        ctrl_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        ctrl_frame.pack(fill="x", padx=16, pady=(0, 12))

        # Start button lives in a fixed-width slot so other buttons
        # never shift. The button itself is packed/unpacked instantly
        # when XML files are detected or the tree becomes empty.
        self._start_slot = ctk.CTkFrame(ctrl_frame, fg_color="transparent",
                                         width=150, height=30)
        self._start_slot.pack(side="left", padx=(0, 6))
        self._start_slot.pack_propagate(False)
        self._start_btn = ctk.CTkButton(self._start_slot, text="Click here to Begin",
                                        command=self._start_upload,
                                        fg_color="#229954", hover_color="#1e8449",
                                        text_color="#ffffff", font=(MODERN_FONT, 12, "bold"),
                                        width=150, height=30, corner_radius=6)
        # NOT packed initially — appears instantly when XML files are loaded
        self._start_btn_visible = False

        self._pause_btn = ctk.CTkButton(ctrl_frame, text="Pause",
                                        command=self._pause_upload,
                                        fg_color="#c0392b", hover_color="#e74c3c",
                                        text_color="#e8e8e8", font=(MODERN_FONT, 12, "bold"),
                                        width=80, height=30, corner_radius=6,
                                        state="disabled")
        self._pause_btn.pack(side="left", padx=(0, 6))

        self._attach_btn = ctk.CTkButton(ctrl_frame, text="Attach Supporting Docs",
                                         command=self._attach_supporting_docs,
                                         fg_color="#117a65", hover_color="#138d75",
                                         text_color="#e8e8e8", font=(MODERN_FONT, 12, "bold"),
                                         width=160, height=30, corner_radius=6,
                                         state="disabled")
        self._attach_btn.pack(side="left", padx=(0, 6))

        self._copy_decl_btn = ctk.CTkButton(ctrl_frame, text="Copy Decl #s from COLS",
                                            command=self._copy_decl_from_cols,
                                            fg_color="#1a5276", hover_color="#2e86c1",
                                            text_color="#e8e8e8", font=(MODERN_FONT, 12, "bold"),
                                            width=150, height=30, corner_radius=6,
                                            state="disabled")
        self._copy_decl_btn.pack(side="right", padx=(6, 0))

        self._paste_btn = ctk.CTkButton(ctrl_frame, text="Paste Declarations to Manifest",
                                        command=self._quick_paste,
                                        fg_color="#6c3483", hover_color="#8e44ad",
                                        text_color="#e8e8e8", font=(MODERN_FONT, 12, "bold"),
                                        width=175, height=30, corner_radius=6,
                                        state="disabled")
        self._paste_btn.pack(side="right", padx=(6, 0))

        # --- Advanced toggle (collapsible) ---
        self._advanced_visible = False
        self._advanced_btn = ctk.CTkButton(self.win, text="Advanced ▶",
                                           command=self._toggle_advanced,
                                           fg_color="transparent", hover_color="#1a1a2e",
                                           text_color="#888888", font=(MODERN_FONT, 10),
                                           width=80, height=20, anchor="w")
        self._advanced_btn.pack(fill="x", padx=16, pady=(2, 0))

        self._advanced_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        # Not packed initially - hidden until toggled

        self._retry_btn = ctk.CTkButton(self._advanced_frame, text="Retry Missed",
                                        command=self._retry_missed,
                                        fg_color="#6a1b9a", hover_color="#8e24aa",
                                        text_color="#e8e8e8", font=(MODERN_FONT, 11, "bold"),
                                        width=110, height=28, corner_radius=6,
                                        state="disabled")
        self._retry_btn.pack(side="left", padx=(0, 6))

        self._skip_btn = ctk.CTkButton(self._advanced_frame, text="Skip Selected",
                                       command=self._skip_selected,
                                       fg_color="#555555", hover_color="#666666",
                                       text_color="#e8e8e8", font=(MODERN_FONT, 11, "bold"),
                                       width=100, height=28, corner_radius=6,
                                       state="disabled")
        self._skip_btn.pack(side="left", padx=(0, 6))

        self._remove_btn = ctk.CTkButton(self._advanced_frame, text="Remove Selected",
                                         command=self._remove_selected,
                                         fg_color="#7b241c", hover_color="#922b21",
                                         text_color="#e8e8e8", font=(MODERN_FONT, 11, "bold"),
                                         width=120, height=28, corner_radius=6,
                                         state="disabled")
        self._remove_btn.pack(side="left", padx=(0, 6))

        self._unskip_btn = ctk.CTkButton(self._advanced_frame, text="Unskip Selected",
                                         command=self._unskip_selected,
                                         fg_color="#2c3e50", hover_color="#34495e",
                                         text_color="#e8e8e8", font=(MODERN_FONT, 11, "bold"),
                                         width=110, height=28, corner_radius=6,
                                         state="disabled")
        self._unskip_btn.pack(side="left", padx=(0, 6))

        # Logging checkbox (ticked by default for now — will be unchecked
        # by default in a future version once everything is proven stable)
        self._logging_var = ctk.IntVar(value=1)
        self._logging_cb = ctk.CTkCheckBox(self._advanced_frame, text="Enable logging",
                                            variable=self._logging_var,
                                            font=(MODERN_FONT, 10), text_color="#888888",
                                            fg_color="#e8820c", hover_color="#ffa726",
                                            height=18)
        self._logging_cb.pack(side="left", padx=(16, 0))

        # "Master Already Uploaded" button — for when you did the master
        # in a previous session and just need to upload house files now.
        self._master_done_btn = ctk.CTkButton(
            self._advanced_frame, text="Master Already Uploaded",
            command=self._mark_master_done,
            fg_color="#1a5276", hover_color="#2471a3",
            text_color="#e8e8e8", font=(MODERN_FONT, 11, "bold"),
            width=160, height=28, corner_radius=6)
        self._master_done_btn.pack(side="left", padx=(16, 0))

        # --- Status bar (with support icon in the corner) ---
        status_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        status_frame.pack(side="bottom", fill="x", padx=16, pady=(4, 10))
        self._status = ctk.CTkLabel(status_frame, text="Ready. Choose a folder to begin.",
                                    font=(MODERN_FONT, 11), text_color="#888888")
        self._status.pack(side="left", fill="x", expand=True)
        self._attach_support_icon(status_frame)

        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        # Restore session if we're restarting after a live update
        restored = self._restore_session()
        if restored:
            try:
                self._status.configure(
                    text="Restored from previous session. Ready to paste or upload.")
            except Exception:
                pass

        # Check for updates in the background
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    def _mark_master_done(self):
        """Mark master files as uploaded and ask for the declaration number.
        Used when the master was uploaded in a previous session and you
        only need to upload house files now."""
        # Mark all master files as uploaded in the tree
        marked = 0
        for item in self._tree.get_children():
            filename = self._tree.set(item, "file")
            status = self._tree.set(item, "status").lower()
            if ("HBL-Master" in filename or "MBL" in filename) and status not in ("uploaded", "error"):
                self._tree.set(item, "status", "Uploaded")
                self._tree.item(item, tags=("uploaded",))
                marked += 1

        if marked == 0 and not self._master_files:
            messagebox.showinfo("No Master Files",
                                "No master files found in the list.")
            return

        # Ask for the master declaration number
        entered_decl = [None]
        def ask_for_decl():
            dialog = ctk.CTkToplevel(self.win)
            dialog.title("Master Declaration Number")
            dialog.configure(fg_color="#1a1a2e")
            dialog.geometry("400x220")
            dialog.resizable(False, False)
            dialog.transient(self.win)
            dialog.grab_set()

            dialog.update_idletasks()
            px = self.win.winfo_rootx() + (self.win.winfo_width() - 400) // 2
            py = self.win.winfo_rooty() + (self.win.winfo_height() - 220) // 2
            dialog.geometry(f"+{px}+{py}")

            ctk.CTkLabel(dialog, text="Enter Master Declaration Number",
                         font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(24, 4))
            ctk.CTkLabel(dialog,
                         text="Enter the declaration number that was assigned\n"
                              "to the master when you uploaded it previously.",
                         font=(MODERN_FONT, 11), text_color="#888888",
                         justify="center").pack(pady=(0, 12))

            entry = ctk.CTkEntry(dialog, width=240, height=34,
                                 fg_color="#0f0f1a", border_color="#e8820c",
                                 border_width=2, corner_radius=6, text_color="#e8e8e8",
                                 font=(MODERN_FONT, 14))
            entry.pack(pady=(0, 12))
            entry.focus_set()

            def on_ok():
                entered_decl[0] = entry.get().strip()
                dialog.destroy()

            def on_cancel():
                dialog.destroy()

            ctk.CTkButton(dialog, text="OK", command=on_ok,
                         fg_color="#e8820c", hover_color="#ffa726",
                         width=100, height=30, corner_radius=6,
                         font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=(80, 8), pady=(0, 16))
            ctk.CTkButton(dialog, text="Cancel", command=on_cancel,
                         fg_color="#555555", hover_color="#666666",
                         width=100, height=30, corner_radius=6,
                         font=(MODERN_FONT, 12, "bold")).pack(side="left", pady=(0, 16))
            dialog.bind("<Return>", lambda e: on_ok())
            dialog.bind("<Escape>", lambda e: on_cancel())

        self.win.after(0, ask_for_decl)
        # Wait for the dialog to close (in the main thread)
        # We use after to check periodically
        def check_result():
            if entered_decl[0] is not None:
                if entered_decl[0]:
                    self._master_decl = entered_decl[0]
                    self._status.configure(
                        text=f"Master marked as uploaded. Declaration #{entered_decl[0]}. "
                             f"Ready to upload house files.")
                else:
                    self._status.configure(
                        text="Master marked as uploaded. No declaration number entered.")
            else:
                self.win.after(100, check_result)
        self.win.after(100, check_result)

    def _toggle_advanced(self):
        """Show/hide the advanced controls."""
        if self._advanced_visible:
            self._advanced_frame.pack_forget()
            self._advanced_btn.configure(text="Advanced ▶")
            self._advanced_visible = False
        else:
            self._advanced_frame.pack(fill="x", padx=16, pady=(2, 4),
                                      after=self._advanced_btn)
            self._advanced_btn.configure(text="Advanced ▼")
            self._advanced_visible = True

    def _log(self, message):
        """Write to the attach docs log if logging is enabled."""
        if not (getattr(self, '_logging_var', None) and self._logging_var.get()):
            return
        try:
            log_path = Path.home() / "Documents" / "attach_docs_log.txt"
            with open(log_path, "a", encoding="utf-8") as logf:
                logf.write(message)
        except Exception:
            pass

    def _upload_log(self, message):
        """Write to the XML upload log if logging is enabled. Kept in a
        separate file (upload_log.txt) so it can be diffed against the
        attach docs log to confirm every uploaded CBY also got docs."""
        if not (getattr(self, '_logging_var', None) and self._logging_var.get()):
            return
        try:
            log_path = Path.home() / "Documents" / "upload_log.txt"
            with open(log_path, "a", encoding="utf-8") as logf:
                logf.write(message)
                logf.flush()
        except Exception:
            pass

    def _skip_selected(self):
        """Mark selected rows as 'skipped' so they won't be uploaded."""
        selected = self._tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Select one or more rows first by clicking on them.")
            return
        for item in selected:
            filename = self._tree.set(item, "file")
            self._tree.set(item, "status", "Skipped")
            self._tree.item(item, tags=("skipped",))
            self._xml_files = [f for f in self._xml_files if f.name != filename]
            self._master_files = [f for f in self._master_files if f.name != filename]
            self._house_files = [f for f in self._house_files if f.name != filename]
        self._status.configure(text=f"Skipped {len(selected)} file(s). They will not be uploaded.")

    def _remove_selected(self):
        """Remove selected rows from the treeview entirely."""
        selected = self._tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Select one or more rows first by clicking on them.")
            return
        count = len(selected)
        answer = messagebox.askyesno("Remove Files",
                                     f"Remove {count} file(s) from the list?\n\n"
                                     f"This only removes them from the uploader window.\n"
                                     f"The actual XML files on disk are not deleted.")
        if not answer:
            return
        for item in selected:
            filename = self._tree.set(item, "file")
            self._tree.delete(item)
            self._xml_files = [f for f in self._xml_files if f.name != filename]
            self._master_files = [f for f in self._master_files if f.name != filename]
            self._house_files = [f for f in self._house_files if f.name != filename]
            if filename in self._decl_numbers:
                del self._decl_numbers[filename]
        self._status.configure(text=f"Removed {count} file(s) from the list.")

    def _unskip_selected(self):
        """Reset selected rows back to 'pending' so they will be uploaded."""
        selected = self._tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Select one or more rows first by clicking on them.")
            return
        restored = 0
        for item in selected:
            status = self._tree.set(item, "status").lower()
            if status == "skipped":
                filename = self._tree.set(item, "file")
                self._tree.set(item, "status", "Pending")
                self._tree.item(item, tags=("pending",))
                if self._xml_folder:
                    file_path = self._xml_folder / filename
                    if file_path.exists():
                        self._xml_files.append(file_path)
                        if "HBL-Master" in filename or "MBL" in filename:
                            self._master_files.append(file_path)
                        else:
                            self._house_files.append(file_path)
                        self._xml_files = natsorted(self._xml_files)
                        self._master_files = natsorted(self._master_files)
                        self._house_files = natsorted(self._house_files)
                        restored += 1
        self._status.configure(text=f"Unskipped {restored} file(s). They will be uploaded.")

    def _show_start_btn(self):
        """Show the Start button instantly."""
        if not self._start_btn_visible:
            self._start_btn.pack(fill="both", expand=True)
            self._start_btn_visible = True
        self._start_btn.configure(state="normal")

    def _hide_start_btn(self):
        """Hide the Start button instantly."""
        if self._start_btn_visible:
            self._start_btn.pack_forget()
            self._start_btn_visible = False

    def _pick_folder(self):
        """Open folder picker and load XML files."""
        folder = filedialog.askdirectory(title="Choose folder containing XML files")
        if not folder:
            return
        self._xml_folder = Path(folder)
        self._folder_label.configure(text=str(self._xml_folder))
        self._load_xml_files()

    def _load_xml_files(self):
        """Scan folder for XML files and populate treeview.
        Master files (HBL-Master*) are listed first, then house files (HBL-CBY*)."""
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._xml_files = []
        self._decl_numbers = {}
        self._master_files = []
        self._house_files = []

        if not self._xml_folder or not self._xml_folder.exists():
            return

        xml_files = natsorted(self._xml_folder.glob("*.xml"))
        for xf in xml_files:
            is_master = "HBL-Master" in xf.name or "MBL" in xf.name
            cby = ""
            if is_master:
                cby = "MASTER"
            else:
                m = re.search(r'CBY\s*(\d+)', xf.name)
                if m:
                    cby = m.group(1)
            self._xml_files.append(xf)
            if is_master:
                self._master_files.append(xf)
            else:
                self._house_files.append(xf)
            self._tree.insert("", "end", values=(xf.name, cby, "Pending", "", ""),
                              tags=("pending",))

        master_count = len(self._master_files)
        house_count = len(self._house_files)
        self._status.configure(
            text=f"Loaded {master_count} master + {house_count} house XML file(s)")
        if self._xml_files:
            self._show_start_btn()
            self._paste_btn.configure(state="normal")
            self._skip_btn.configure(state="normal")
            self._remove_btn.configure(state="normal")
            self._unskip_btn.configure(state="normal")
        else:
            self._hide_start_btn()

    def _set_row_docs_status(self, cby, docs_status, docs_text=""):
        """Update the Docs column text for a row matching the given CBY.
        Does NOT change row color — the row keeps its upload status color.
        The docs text itself communicates the state."""
        for item in self._tree.get_children():
            row_cby = self._tree.set(item, "cby")
            if row_cby == cby:
                self._tree.set(item, "docs", docs_text)
                break

    def _update_status(self, text):
        """Update the status bar and mirror to autopilot bar if active."""
        self._status.configure(text=text)
        if self._autopilot_bar and self._autopilot_bar.is_alive():
            self._autopilot_bar.set_activity(text)

    def _set_row_status(self, filename, status, decl_num=""):
        """Update a row's status and color in the treeview."""
        for item in self._tree.get_children():
            vals = self._tree.set(item, "file")
            if vals == filename:
                self._tree.set(item, "status", status.capitalize())
                if decl_num:
                    self._tree.set(item, "decl", decl_num)
                    self._decl_numbers[filename] = decl_num
                self._tree.item(item, tags=(status,))
                break

    def _start_upload(self):
        """Two-phase: first click launches Edge, second click starts uploading."""
        if not _SELENIUM_AVAILABLE:
            messagebox.showerror("Missing Dependency",
                                 "Selenium is not installed.\n\n"
                                 "Install it with:\n  pip install selenium natsort")
            return
        if not self._xml_files:
            messagebox.showwarning("No Files", "Please choose a folder with XML files first.")
            return

        if not self._edge_launched:
            # Phase 1: Launch Edge
            self._launch_edge()
        else:
            # Phase 2: Start uploading
            self._begin_upload()

    def _launch_edge(self):
        """Launch Edge browser and navigate to COLS. Does NOT start uploading."""
        self._edge_ready = False
        self._start_btn.configure(state="disabled", text="Launching...")
        self._status.configure(text="Starting Edge browser... please wait")
        threading.Thread(target=self._edge_launch_worker, daemon=True).start()

    def _edge_launch_worker(self):
        """Background thread: just launches Edge and navigates to COLS."""
        # Animate status while Edge is starting
        dots = ["", ".", "..", "..."]
        dot_idx = [0]
        def update_dots():
            if self._edge_ready:
                return
            self._status.configure(text=f"Starting Edge browser{dots[dot_idx[0] % 4]}... please wait")
            dot_idx[0] += 1
            self.win.after(400, update_dots)
        self.win.after(100, update_dots)

        try:
            options = Options()
            options.add_argument("--inprivate")
            options.add_experimental_option("detach", True)
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            # Performance: skip unnecessary loading for faster startup
            options.add_argument("--disable-sync")
            options.add_argument("--disable-background-networking")
            options.add_argument("--disable-default-apps")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-translate")
            options.add_argument("--disable-background-timer-throttling")
            options.add_argument("--disable-renderer-backgrounding")
            options.add_argument("--disable-backgrounding-occluded-windows")
            options.page_load_strategy = "eager"
            self._driver = webdriver.Edge(options=options)
            self._driver.get(self.UPLOAD_URL)
        except Exception as e:
            self.win.after(0, lambda: self._status.configure(
                text=f"Failed to start Edge: {e}"))
            self.win.after(0, lambda: self._start_btn.configure(
                state="normal", text="Click here to Begin"))
            return

        # Edge is ready - switch button to "Start Upload"
        self._edge_ready = True
        self._edge_launched = True
        self.win.after(0, lambda: self._status.configure(
            text="Edge is open. Log in to COLS, then click 'Start Upload'"))
        self.win.after(0, lambda: self._start_btn.configure(
            state="normal", text="Start Upload"))
        self.win.after(0, lambda: self._copy_decl_btn.configure(state="normal"))
        self.win.after(0, lambda: self._attach_btn.configure(state="normal"))

    def _begin_upload(self):
        """Phase 2: Start the actual file upload process."""
        if self._uploading:
            return
        self._uploading = True
        self._paused = False
        self._stop_requested = False
        self._start_btn.configure(state="disabled", text="Uploading...")
        self._pause_btn.configure(state="normal", text="Pause")
        self._attach_btn.configure(state="disabled")
        self._retry_btn.configure(state="disabled")
        self._status.configure(text="Starting upload...")
        threading.Thread(target=self._upload_worker, daemon=True).start()

    def _upload_worker(self):
        """Background thread: uploads files to COLS.
        Master files are uploaded first. After each master upload, we capture
        the declaration number and prompt the user to continue before uploading
        house files.
        Edge must already be launched and user logged in before calling this."""
        if not self._driver:
            self.win.after(0, lambda: self._status.configure(text="Error: Edge not launched"))
            self.win.after(0, self._reset_buttons)
            self._uploading = False
            return

        # Initialize the upload log (fresh file each run) with the full planned
        # list of CBYs. This is the authoritative list to diff against the
        # attach docs log later.
        if getattr(self, '_logging_var', None) and self._logging_var.get():
            try:
                import datetime
                _ulp = Path.home() / "Documents" / "upload_log.txt"
                with open(_ulp, "w", encoding="utf-8") as _uf:
                    _uf.write(f"upload_log — run started {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")
                    _uf.write(f"Master files: {len(self._master_files)}  House files: {len(self._house_files)}\n")
                    _uf.write("\nPlanned house CBYs:\n")
                    for hf in self._house_files:
                        m = re.search(r'CBY\s*(\d+)', hf.name)
                        cby = m.group(1) if m else "?"
                        _uf.write(f"  CBY {cby}  <- {hf.name}\n")
                    _uf.write("\n" + "=" * 60 + "\n")
            except Exception as e:
                print(f"[Upload] Log init error: {e}")

        # ---- PHASE 1: Upload master files (if any) ----
        for master_file in self._master_files:
            if self._stop_requested:
                break
            while self._paused:
                time.sleep(0.5)
                if self._stop_requested:
                    break
            if self._stop_requested:
                break

            # Check if already uploaded
            current_status = ""
            for item in self._tree.get_children():
                if self._tree.set(item, "file") == master_file.name:
                    current_status = self._tree.set(item, "status").lower()
                    break
            if current_status in ("uploaded", "skipped", "error"):
                continue

            self.win.after(0, lambda f=master_file: self._set_row_status(f.name, "uploading"))
            self.win.after(0, lambda f=master_file: self._status.configure(
                text=f"Uploading MASTER: {f.name}..."))

            try:
                wait = WebDriverWait(self._driver, self.WAIT_TIMEOUT)
                file_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
                file_input.send_keys(str(master_file.resolve()))

                # Try multiple button XPath patterns (short timeout per attempt)
                upload_clicked = False
                button_xpaths = [
                    "//button[contains(text(), 'Upload Declaration')]",
                    "//button[contains(text(), 'Upload')]",
                    "//input[@type='submit' and contains(@value, 'Upload')]",
                    "//button[@type='submit']",
                    "//a[contains(text(), 'Upload')]",
                ]
                short_wait = WebDriverWait(self._driver, 3)  # 3s per XPath, not 15s
                for xpath in button_xpaths:
                    try:
                        upload_button = short_wait.until(
                            EC.element_to_be_clickable((By.XPATH, xpath)))
                        upload_button.click()
                        upload_clicked = True
                        break
                    except Exception:
                        continue

                if not upload_clicked:
                    raise Exception("Could not find upload button on page")

                # Smart wait: poll for page to settle instead of fixed sleep.
                # For master uploads, just wait for the page to stop loading.
                # Check if the file input reappears (page ready for next upload)
                # or an error popup appears.
                # Max wait is 15s — generous for slow COLS days.
                for _mw in range(30):  # 30 * 0.5s = 15s max
                    time.sleep(0.5)
                    try:
                        body_lower = self._driver.find_element(
                            By.TAG_NAME, "body").text.lower()
                        if any(err in body_lower for err in [
                            "failed to upload", "an error occurred",
                            "invalid file", "not a valid",
                            "rejected", "please make sure"]):
                            break  # Error appeared, stop waiting
                        # Check if file input reappeared (page ready)
                        inputs = self._driver.find_elements(
                            By.XPATH, "//input[@type='file']")
                        if inputs and inputs[0].is_displayed():
                            break  # Page is ready for next upload
                    except Exception:
                        pass

                # Mark master as uploaded - declaration number entered manually in the gate
                self.win.after(0, lambda f=master_file: self._set_row_status(
                    f.name, "uploaded"))
                self.win.after(0, lambda f=master_file: self._status.configure(
                    text=f"MASTER uploaded: {f.name}. Waiting for declaration number..."))

            except Exception as e:
                self.win.after(0, lambda f=master_file, e=e: self._set_row_status(f.name, "error"))
                self.win.after(0, lambda f=master_file, e=e: self._status.configure(
                    text=f"MASTER MISSED: {f.name} - {type(e).__name__}: {e}"))

            time.sleep(self.SHORT_DELAY)

        # ---- GATE: Ask user for master declaration number ----
        # Skip if already known (e.g., auto-retry of house files only)
        if self._stop_requested:
            self.win.after(0, self._upload_finished)
            return

        if not self._master_decl:
            # Show popup to get master declaration number
            self._master_decl = ""  # clear any stale value
            entered_decl = [None]
            def ask_for_decl():
                dialog = ctk.CTkToplevel(self.win)
                dialog.title("Master Declaration Number Required")
                dialog.configure(fg_color="#1a1a2e")
                dialog.geometry("400x240")
                dialog.resizable(False, False)
                dialog.transient(self.win)
                dialog.grab_set()

                # Center on parent
                dialog.update_idletasks()
                px = self.win.winfo_rootx() + (self.win.winfo_width() - 400) // 2
                py = self.win.winfo_rooty() + (self.win.winfo_height() - 240) // 2
                dialog.geometry(f"+{px}+{py}")

                ctk.CTkLabel(dialog, text="Master Declaration Number Required",
                             font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(24, 4))
                ctk.CTkLabel(dialog,
                             text="Enter the declaration number assigned to the master.\n"
                                  "You can find it on the COLS page after submitting.\n"
                                  "House files will NOT upload until this is entered.",
                             font=(MODERN_FONT, 11), text_color="#888888",
                             justify="center").pack(pady=(0, 12))

                entry = ctk.CTkEntry(dialog, width=240, height=34,
                                     fg_color="#0f0f1a", border_color="#e8820c",
                                     border_width=2, corner_radius=6, text_color="#e8e8e8",
                                     font=(MODERN_FONT, 14))
                entry.pack(pady=(0, 12))
                entry.focus_set()

                def on_ok():
                    entered_decl[0] = entry.get().strip()
                    dialog.destroy()

                def on_cancel():
                    dialog.destroy()

                btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
                btn_frame.pack(pady=(0, 16))
                ctk.CTkButton(btn_frame, text="OK - Start House Uploads", command=on_ok,
                              fg_color="#2e8b57", hover_color="#3cb371", width=170, height=32,
                              corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)
                ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel,
                              fg_color="#555555", hover_color="#666666", width=80, height=32,
                              corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)

                entry.bind("<Return>", lambda e: on_ok())
                entry.bind("<Escape>", lambda e: on_cancel())
                dialog.protocol("WM_DELETE_WINDOW", on_cancel)

                self.win.wait_window(dialog)

            # Run the dialog on the main thread, wait for result in background thread
            import threading as _th
            evt = _th.Event()
            def run_dialog():
                ask_for_decl()
                evt.set()
            self.win.after(0, run_dialog)
            evt.wait()

            if entered_decl[0]:
                self._master_decl = entered_decl[0]
                # Update master rows in tree with the entered declaration number
                for master_file in self._master_files:
                    self.win.after(0, lambda f=master_file, d=entered_decl[0]:
                        self._set_row_status(f.name, "uploaded", d))
                self.win.after(0, lambda d=entered_decl[0]: self._status.configure(
                    text=f"Master declaration #: {d}. Starting house uploads..."))
            else:
                # User cancelled - don't proceed
                self.win.after(0, lambda: self._status.configure(
                    text="STOPPED: Master declaration number not entered. House files will not upload."))
                self.win.after(0, self._upload_finished)
                return


        # ---- PHASE 2: Upload house files ----
        # Show floating progress bar if any auto mode is enabled
        if (self._auto_retry_var.get() or self._auto_attach_var.get()) and not self._autopilot_bar:
            self._autopilot_bar = AutoPilotProgressBar(self.win)
        if self._autopilot_bar and self._autopilot_bar.is_alive():
            self._autopilot_bar.set_phase("Uploading House Declarations...")
            self._autopilot_bar.set_progress(0, len(self._house_files))
        for xml_file in self._house_files:
            if self._stop_requested:
                break
            while self._paused:
                time.sleep(0.5)
                if self._stop_requested:
                    break
            if self._stop_requested:
                break

            # Check if already uploaded
            current_status = ""
            for item in self._tree.get_children():
                if self._tree.set(item, "file") == xml_file.name:
                    current_status = self._tree.set(item, "status").lower()
                    break
            if current_status in ("uploaded", "skipped", "error"):
                continue

            # Pre-check: if this file is already on the COLS page (e.g., from
            # a previous session), mark as uploaded and skip. Only trust a FULL
            # filename match — the number-only match can hit the file-picker
            # label of a different file and cause false "uploaded" marks.
            try:
                body_check = self._driver.find_element(By.TAG_NAME, "body").text
                if xml_file.name in body_check:
                    _cby_m = re.search(r'CBY\s*(\d+)', xml_file.name)
                    self._upload_log(
                        f"CBY {_cby_m.group(1) if _cby_m else '?'}: PRE-CHECK skip "
                        f"(matched by filename)  ({xml_file.name})\n")
                    self.win.after(0, lambda fn=xml_file.name: self._set_row_status(fn, "uploaded"))
                    self.win.after(0, lambda fn=xml_file.name: self._status.configure(
                        text=f"Already on page: {fn} (skipped re-upload)"))
                    time.sleep(self.SHORT_DELAY)
                    continue
            except Exception:
                pass

            self.win.after(0, lambda f=xml_file: self._set_row_status(f.name, "uploading"))
            self.win.after(0, lambda f=xml_file: self._update_status(
                f"Uploading: {f.name}..."))

            try:
                wait = WebDriverWait(self._driver, self.WAIT_TIMEOUT)
                file_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))

                # Count "Passed" occurrences BEFORE uploading. Each successfully
                # uploaded declaration row shows "Passed" on the page. After we
                # click upload, we wait for this count to go up by 1 — proving a
                # new row genuinely appeared. This is fast (body.text includes the
                # new row's text quickly) and reliable ("Passed" only appears on
                # genuine uploads, not in the file picker label).
                # We also count "error" — if a row lands with an error status
                # instead of "Passed", the error count goes up. That also means
                # the upload settled (just with an error), so we stop waiting
                # and mark it as an error for human review.
                def _count_word(word):
                    try:
                        body = self._driver.find_element(By.TAG_NAME, "body").text
                        return body.lower().count(word)
                    except Exception:
                        return -1
                passed_before = _count_word("passed")
                error_before = _count_word("error")

                file_input.send_keys(str(xml_file.resolve()))

                # Click the Upload button. The button is always the same on COLS
                # — don't overthink it. Find it and click it immediately, no
                # visibility checks, no multiple XPath guessing. If the first
                # attempt fails, try a real Selenium .click() as fallback.
                # The "Passed" count wait after this handles verification AND
                # tells us when COLS is ready for the next file.
                upload_clicked = False
                time.sleep(0.3)  # brief moment for COLS to register the file
                try:
                    btn = self._driver.find_element(
                        By.XPATH, "//button[contains(text(), 'Upload Declaration')]")
                    self._driver.execute_script("arguments[0].click();", btn)
                    upload_clicked = True
                except Exception:
                    pass
                # Fallback: real Selenium click (fires ADF event handlers)
                if not upload_clicked:
                    try:
                        btn = self._driver.find_element(
                            By.XPATH, "//button[contains(text(), 'Upload Declaration')]")
                        btn.click()
                        upload_clicked = True
                    except Exception:
                        pass
                # Last fallback: try broader XPath patterns
                if not upload_clicked:
                    for xpath in [
                        "//button[contains(text(), 'Upload')]",
                        "//input[@type='submit' and contains(@value, 'Upload')]",
                        "//button[@type='submit']",
                        "//a[contains(text(), 'Upload')]",
                    ]:
                        try:
                            btn = self._driver.find_element(By.XPATH, xpath)
                            self._driver.execute_script("arguments[0].click();", btn)
                            upload_clicked = True
                            break
                        except Exception:
                            continue

                if not upload_clicked:
                    # Last resort: refresh and try once more
                    self.win.after(0, lambda f=xml_file: self._update_status(
                        f"Upload button not found, refreshing page for {f.name}..."))
                    try:
                        self._driver.navigate().refresh()
                        time.sleep(3)
                        retry_wait = WebDriverWait(self._driver, 10)
                        file_input2 = retry_wait.until(
                            EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
                        file_input2.send_keys(str(xml_file.resolve()))
                        time.sleep(0.3)
                        btn = retry_wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//button[contains(text(), 'Upload Declaration')]")))
                        self._driver.execute_script("arguments[0].click();", btn)
                        upload_clicked = True
                    except Exception:
                        pass
                    except Exception:
                        pass
                    if not upload_clicked:
                        raise Exception("Could not find upload button (even after refresh)")

                # Smart wait: poll for the "Passed" count OR "error" count to
                # increase — either means a new declaration row landed on the
                # page and the upload settled. "Passed" = success, "error" =
                # the row appeared with an error status (human review needed).
                # Either way we stop waiting immediately. Also check for error
                # popups (ADF dialogs with specific rejection messages).
                page_settled = False
                _row_had_error = False
                for _pw in range(60):  # 60 * 0.3s = 18s max
                    time.sleep(0.3)
                    try:
                        body_text = self._driver.find_element(
                            By.TAG_NAME, "body").text
                        body_lower = body_text.lower()
                        # Check for error popup (ADF rejection dialog)
                        if any(err in body_lower for err in [
                            "failed to upload", "an error occurred",
                            "invalid file", "not a valid",
                            "rejected", "please make sure"]):
                            page_settled = True
                            break
                        # Success: "Passed" count increased (new row appeared)
                        if passed_before >= 0 and _count_word("passed") > passed_before:
                            page_settled = True
                            break
                        # Error status: "error" count increased (row landed with
                        # error instead of "Passed" — upload settled, needs review)
                        if error_before >= 0 and _count_word("error") > error_before:
                            page_settled = True
                            _row_had_error = True
                            break
                    except Exception:
                        pass

                # Verify the file actually appeared on the COLS page.
                # Also detect error popups (ADF message dialogs) that appear
                # when COLS rejects the XML — these must be dismissed and
                # marked as errors (not missed) to avoid retry loops.
                upload_verified = False
                had_error_popup = False
                error_message = ""
                self._verify_match_kind = "none"
                try:
                    body_text = self._driver.find_element(
                        By.TAG_NAME, "body").text
                    body_lower = body_text.lower()

                    # Check for ADF error popup — the dialog has id d1_msgDlg_cancel
                    # and shows error text like "Failed to upload XML..."
                    # Also check for validation errors that appear inline.
                    error_phrases = [
                        "failed to upload xml",
                        "failed to upload",
                        "please make sure every start-tag",
                        "an error occurred",
                        "error processing",
                        "invalid file",
                        "only xml",
                        "not a valid",
                        "upload failed",
                        "rejected",
                        "validation error",
                        "must be numeric",
                        "is required",
                        "cannot be null",
                        "duplicate declaration",
                        "already exists",
                    ]
                    for phrase in error_phrases:
                        if phrase in body_lower:
                            had_error_popup = True
                            # Try to extract the actual error message
                            try:
                                # ADF error dialog text is in a TD with class x1e4
                                error_cells = self._driver.find_elements(
                                    By.XPATH, "//td[@class='x1e4']")
                                for cell in error_cells:
                                    cell_text = cell.text.strip()
                                    if cell_text and len(cell_text) > 5:
                                        error_message = cell_text[:200]
                                        break
                            except Exception:
                                pass
                            if not error_message:
                                # Fallback: find the error text in body
                                for line in body_text.split("\n"):
                                    for phrase in error_phrases:
                                        if phrase in line.lower():
                                            error_message = line.strip()[:200]
                                            break
                                    if error_message:
                                        break
                            break

                    # Verify upload: "Passed" count went up (a new declaration
                    # row genuinely appeared). This is the reliable signal —
                    # "Passed" only shows on uploaded rows, never in the picker.
                    if passed_before >= 0 and _count_word("passed") > passed_before:
                        upload_verified = True
                        self._verify_match_kind = "passed-count"

                    # If the row landed with "error" status instead of "Passed",
                    # the upload settled but the declaration has an error. Mark
                    # it as an error (not missed) so it doesn't get retried —
                    # humans need to review and fix it on COLS.
                    if _row_had_error and not upload_verified and not had_error_popup:
                        had_error_popup = True  # reuse the error-handling path
                        error_message = "Declaration uploaded with error status on COLS"
                        self._verify_match_kind = "error-row"

                    # NOTE: Inline error indicator check removed — it was scanning
                    # the ENTIRE page for any "Error" text, which matched errors
                    # from OTHER rows and caused false positives. Now, uncertain
                    # uploads (no explicit popup, no verification) fall through
                    # to "missed" status and get retried automatically.
                except Exception:
                    upload_verified = True  # Give benefit of the doubt on error

                # Dismiss error popup if present (click OK button)
                if had_error_popup:
                    try:
                        ok_btn = self._driver.find_element(
                            By.ID, "d1_msgDlg_cancel")
                        ok_btn.click()
                        time.sleep(0.5)
                    except Exception:
                        try:
                            # Fallback: find any link/button with text "OK"
                            ok_links = self._driver.find_elements(
                                By.XPATH, "//a[contains(text(), 'OK')]")
                            if ok_links:
                                ok_links[0].click()
                                time.sleep(0.5)
                        except Exception:
                            pass

                _cby_m = re.search(r'CBY\s*(\d+)', xml_file.name)
                _cby_log = _cby_m.group(1) if _cby_m else "?"

                if upload_verified and not had_error_popup:
                    self.win.after(0, lambda f=xml_file: self._set_row_status(
                        f.name, "uploaded"))
                    self.win.after(0, lambda f=xml_file: self._update_status(
                        f"Uploaded: {f.name}"))
                    self._upload_log(f"CBY {_cby_log}: UPLOADED "
                                     f"(verified by {getattr(self, '_verify_match_kind', '?')})  "
                                     f"({xml_file.name})\n")
                elif had_error_popup:
                    # Error popup on COLS — mark as error (not retried automatically)
                    self.win.after(0, lambda f=xml_file: self._set_row_status(
                        f.name, "error"))
                    if self._autopilot_bar and self._autopilot_bar.is_alive():
                        self._autopilot_bar.add_issue()
                    self.win.after(0, lambda f=xml_file, msg=error_message: self._update_status(
                        f"ERROR: {f.name} — {msg}" if msg else f"ERROR: {f.name} rejected by COLS"))
                    self._upload_log(f"CBY {_cby_log}: ERROR — {error_message or 'rejected by COLS'}  ({xml_file.name})\n")
                else:
                    self.win.after(0, lambda f=xml_file: self._set_row_status(
                        f.name, "missed"))
                    self.win.after(0, lambda f=xml_file: self._update_status(
                        f"WARNING: {f.name} not confirmed on page — will retry"))
                    if self._autopilot_bar and self._autopilot_bar.is_alive():
                        self._autopilot_bar.add_issue()
                    self._upload_log(f"CBY {_cby_log}: MISSED (not confirmed)  ({xml_file.name})\n")

            except Exception as e:
                self.win.after(0, lambda f=xml_file, e=e: self._set_row_status(f.name, "error"))
                self.win.after(0, lambda f=xml_file, e=e: self._status.configure(
                    text=f"Error: {f.name} - {type(e).__name__}: {e}"))
                _cby_m = re.search(r'CBY\s*(\d+)', xml_file.name)
                self._upload_log(f"CBY {_cby_m.group(1) if _cby_m else '?'}: EXCEPTION — {type(e).__name__}: {e}  ({xml_file.name})\n")

            # Update autopilot progress bar
            if self._autopilot_bar and self._autopilot_bar.is_alive():
                done = 0
                for item in self._tree.get_children():
                    s = self._tree.set(item, "status").lower()
                    if s in ("uploaded", "error", "skipped"):
                        done += 1
                total_files = len(self._house_files)
                self._autopilot_bar.set_progress(done, total_files)
            time.sleep(self.SHORT_DELAY)

        self.win.after(0, self._upload_finished)

    def _on_tree_double_click(self, event):
        """Allow manual editing of the Declaration # column."""
        if self._decl_edit_win is not None:
            try:
                self._decl_edit_win.destroy()
            except Exception:
                pass
            self._decl_edit_win = None

        region = self._tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self._tree.identify_column(event.x)
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        col_idx = int(col.replace("#", "")) - 1
        cols = ["file", "cby", "status", "docs", "decl"]
        if col_idx < 0 or col_idx >= len(cols):
            return
        key = cols[col_idx]
        if key != "decl":
            return  # Only allow editing the Declaration # column

        bbox = self._tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        cx = self._tree.winfo_rootx() + x
        cy = self._tree.winfo_rooty() + y
        current_val = self._tree.set(row_id, key)

        self._decl_edit_win = tk.Toplevel(self.win)
        self._decl_edit_win.overrideredirect(True)
        self._decl_edit_win.geometry(f"{w}x{h}+{cx}+{cy}")
        self._decl_edit_win.attributes("-topmost", True)

        widget = ctk.CTkEntry(self._decl_edit_win, width=w, height=h,
                              fg_color="#0f0f1a", border_color="#e8820c",
                              border_width=2, corner_radius=2, text_color="#e8e8e8",
                              font=(MODERN_FONT, 10))
        widget.pack(fill="both", expand=True)
        widget.insert(0, current_val)
        widget.focus_set()
        widget.select_range(0, "end")

        def finish(e=None):
            value = widget.get().strip()
            self._tree.set(row_id, key, value)
            filename = self._tree.set(row_id, "file")
            if value:
                self._decl_numbers[filename] = value
            elif filename in self._decl_numbers:
                del self._decl_numbers[filename]
            try:
                self._decl_edit_win.destroy()
            except Exception:
                pass
            self._decl_edit_win = None

        def cancel(e=None):
            try:
                self._decl_edit_win.destroy()
            except Exception:
                pass
            self._decl_edit_win = None

        widget.bind("<Return>", finish)
        widget.bind("<Escape>", cancel)
        widget.bind("<FocusOut>", finish)

    def _extract_decl_number(self):
        """Try to extract the declaration number assigned by COLS after upload.
        Searches both visible text and page source. May need adjustment based on
        the actual COLS page structure."""
        if not self._driver:
            return ""

        # Patterns to try - ordered from most specific to most generic
        patterns = [
            r'Declaration\s*(?:Number|#|No\.?)\s*[:]?\s*(\d{5,})',
            r'Declaration\s*(?:Number|#|No\.?)\s*[:]?\s*([A-Z0-9-]{5,})',
            r'(?:Created|Submitted|Assigned|Generated).*?(\d{5,})',
            r'Reference\s*(?:Number|#)?\s*[:]?\s*(\d{5,})',
            r'(?:Dec|Decl)\.?\s*[:#]?\s*(\d{5,})',
            r'(?:Number|No\.?)\s*[:]?\s*(\d{7,})',
            r'(\d{7,})',  # Any 7+ digit number as last resort
        ]

        # Strategy 1: Search visible body text
        try:
            body_text = self._driver.find_element(By.TAG_NAME, "body").text
            for pattern in patterns:
                m = re.search(pattern, body_text, re.IGNORECASE)
                if m:
                    return m.group(1)
        except Exception:
            pass

        # Strategy 2: Search page source (catches hidden elements, JavaScript-rendered text)
        try:
            page_source = self._driver.page_source
            # In page source, look for patterns in text nodes
            for pattern in patterns:
                m = re.search(pattern, page_source, re.IGNORECASE)
                if m:
                    return m.group(1)
        except Exception:
            pass

        # Strategy 3: Look for any element containing "declaration" text
        try:
            elements = self._driver.find_elements(By.XPATH,
                "//*[contains(text(), 'declaration') or contains(text(), 'Declaration')]")
            for el in elements:
                text = el.text
                for pattern in patterns:
                    m = re.search(pattern, text, re.IGNORECASE)
                    if m:
                        return m.group(1)
        except Exception:
            pass

        return ""

    def _pause_upload(self):
        """Toggle pause/resume."""
        if self._paused:
            self._paused = False
            self._pause_btn.configure(text="Pause", fg_color="#c0392b", hover_color="#e74c3c")
            self._attach_btn.configure(state="disabled")
            self._status.configure(text="Resumed uploading...")
        else:
            self._paused = True
            self._pause_btn.configure(text="Resume", fg_color="#2e8b57", hover_color="#3cb371")
            self._attach_btn.configure(state="normal")
            self._status.configure(text="Paused. Click Resume to continue.")

    def _retry_missed(self):
        """Reset failed/skipped rows to pending and restart upload."""
        if not self._driver or not self._uploading:
            # Need to restart the whole process
            self._start_upload()
            return

        # Reset missed rows to pending (NOT errors - those need manual fixing)
        for item in self._tree.get_children():
            status = self._tree.set(item, "status").lower()
            if status in ("missed",):
                filename = self._tree.set(item, "file")
                self._tree.set(item, "status", "Pending")
                self._tree.item(item, tags=("pending",))

        # Re-run upload for pending files (worker handles master-first automatically)
        self._paused = False
        self._pause_btn.configure(text="Pause", state="normal", fg_color="#c0392b", hover_color="#e74c3c")
        self._stop_requested = False
        if self._autopilot_bar and self._autopilot_bar.is_alive():
            self._autopilot_bar.set_phase("Retrying Missed Declarations...")
        threading.Thread(target=self._upload_worker, daemon=True).start()

    def _upload_finished(self):
        """Called when upload loop completes."""
        self._uploading = False
        if self._edge_launched:
            self._start_btn.configure(state="normal", text="Start Upload")
        else:
            self._start_btn.configure(state="normal", text="Click here to Begin")
        self._pause_btn.configure(state="disabled", text="Pause", fg_color="#c0392b", hover_color="#e74c3c")
        self._attach_btn.configure(state="normal")
        self._retry_btn.configure(state="normal")

        # Count results
        uploaded = errors = missed = pending = 0
        for item in self._tree.get_children():
            s = self._tree.set(item, "status").lower()
            if s == "uploaded":
                uploaded += 1
            elif s == "error":
                errors += 1
            elif s == "missed":
                missed += 1
            elif s == "pending":
                pending += 1

        self._status.configure(
            text=f"Done: {uploaded} uploaded, {errors} errors, {missed} missed, {pending} pending")

        # Auto-retry missed files if checkbox is checked
        # (errors are NOT auto-retried - they need manual fixing)
        if missed > 0 and self._auto_retry_var.get() and not self._stop_requested:
            self._status.configure(
                text=f"Auto-retrying {missed} missed file(s)...")
            if self._autopilot_bar and self._autopilot_bar.is_alive():
                self._autopilot_bar.set_phase("Retrying Missed Declarations...")
                self._autopilot_bar.set_progress(0, missed)
            self.win.after(2000, self._retry_missed)  # Wait 2s then retry
            return

        # Auto-attach supporting docs if checkbox is checked
        # (proceed even if some files marked "error" — they may have uploaded
        #  successfully but been misdetected; genuinely errored files won't
        #  have "Upload Supporting Documents" links on COLS so they're skipped)
        if self._auto_attach_var.get() and uploaded > 0 and not self._stop_requested:
            self._status.configure(
                text="Auto-attaching supporting documents...")
            if self._autopilot_bar and self._autopilot_bar.is_alive():
                self._autopilot_bar.set_phase("Attaching Declaration Supporting Docs...")
                self._autopilot_bar.set_progress(0, uploaded)
            self.win.after(2000, self._auto_attach_supporting_docs)
            return

        # If no auto-attach will run, complete the autopilot bar
        if not (self._auto_attach_var.get() and uploaded > 0 and not self._stop_requested):
            if self._autopilot_bar and self._autopilot_bar.is_alive():
                self._autopilot_bar.complete(uploaded, errors, missed, 0, 0)
        if errors > 0:
            messagebox.showinfo("Upload Complete",
                                f"Uploaded: {uploaded}\nErrors: {errors}\nMissed: {missed}\n\n"
                                f"Errors need to be fixed manually:\n"
                                f"  1. Inspect the file on COLS to see the issue\n"
                                f"  2. Fix the XML file\n"
                                f"  3. Re-upload it manually on COLS\n"
                                f"  4. Double-click the Declaration # column to enter the number\n\n"
                                f"Missed files can be retried via 'Retry Missed' under Advanced.\n"
                                f"Or check 'Auto-retry missed' at the top for automatic retry.")
        else:
            messagebox.showinfo("Upload Complete",
                                f"All {uploaded} file(s) uploaded successfully!\n\n"
                                f"You can now attach supporting docs and submit on COLS.")

    def _reset_buttons(self):
        if self._edge_launched:
            self._start_btn.configure(state="normal", text="Start Upload")
        else:
            self._start_btn.configure(state="normal", text="Click here to Begin")
        self._pause_btn.configure(state="disabled", text="Pause", fg_color="#c0392b", hover_color="#e74c3c")
        self._retry_btn.configure(state="disabled")
        self._attach_btn.configure(state="disabled")

    def _auto_attach_supporting_docs(self):
        """Auto-attach: find PDFs and attach without showing confirmation dialog."""
        self._attach_supporting_docs(skip_dialog=True)

    def _attach_supporting_docs(self, skip_dialog=False):
        """Find PDFs in the manifest folder and attach them to declarations on COLS."""
        if not self._driver:
            messagebox.showwarning("No Browser", "Launch Edge first, then use this button.")
            return
        if not self._xml_folder:
            messagebox.showwarning("No Folder", "Choose an XML folder first.")
            return

        # Find PDFs
        pdfs = self._find_supporting_pdfs()
        if not pdfs:
            messagebox.showwarning("No PDFs Found",
                                   "Could not find any PDF files.\n\n"
                                   "Searched in:\n"
                                   f"  {self._xml_folder}\n"
                                   f"  {self._xml_folder.parent}\n"
                                   "and their subfolders.\n\n"
                                   "Make sure your PDF files are in the manifest\n"
                                   "folder or a subfolder next to the XML files.")
            return

        # Clear the log at the start of each run (only if logging is enabled)
        if getattr(self, '_logging_var', None) and self._logging_var.get():
            _log_path_init = Path.home() / "Documents" / "attach_docs_log.txt"
            with open(_log_path_init, "w", encoding="utf-8") as _lf:
                import datetime
                _lf.write(f"attach_docs_log — run started {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")

        # Categorize PDFs: AWB/BOL vs invoices vs placeholder
        awb_bol_pdfs = []
        invoice_pdfs = []
        placeholder_pdf = None
        for pdf in pdfs:
            name = pdf.name.upper()
            # Remove leading # for detection
            clean_name = name.lstrip("#")
            if clean_name.startswith("378") or name.startswith("SMLU") or name.startswith("#SMLU"):
                awb_bol_pdfs.append(pdf)
            elif "MANIFEST" in name and "CUSTOM" in name:
                placeholder_pdf = pdf
            else:
                invoice_pdfs.append(pdf)

        # Match invoice PDFs to CBY numbers from the treeview
        cby_pdfs_map = {}  # cby_number -> list of PDF paths
        all_tree_cbys = []
        for item in self._tree.get_children():
            cby = self._tree.set(item, "cby")
            if not cby or cby == "MASTER":
                continue
            all_tree_cbys.append(cby)
            matched = []
            for pdf in invoice_pdfs:
                stem = pdf.stem  # filename without extension
                if stem.upper().startswith(cby.upper()):
                    # Check next char after CBY is not a digit
                    rest = stem[len(cby):]
                    if not rest or rest[0] in "_-. " or rest[0].isalpha():
                        matched.append(pdf)
            if matched:
                cby_pdfs_map[cby] = matched

        # For CBYs with no invoice, use the Manifest and Customs PDF as placeholder
        placeholder_cbys = []
        if placeholder_pdf:
            for cby in all_tree_cbys:
                if cby not in cby_pdfs_map:
                    cby_pdfs_map[cby] = [placeholder_pdf]
                    placeholder_cbys.append(cby)

        # Show confirmation dialog
        if not cby_pdfs_map and not awb_bol_pdfs:
            messagebox.showwarning("No Matches",
                                   "Found PDFs but could not match them to any CBY numbers.\n\n"
                                   f"PDFs found: {len(pdfs)}\n"
                                   f"CBYs in treeview: {len(all_tree_cbys)}\n\n"
                                   "Make sure PDF filenames start with the CBY number\n"
                                   "(e.g., '200_1.pdf' for CBY 200).")
            return

        if skip_dialog:
            # Auto mode: skip confirmation, go straight to attaching
            self._attach_btn.configure(state="disabled")
            threading.Thread(target=self._attach_docs_worker,
                             args=(cby_pdfs_map, awb_bol_pdfs, placeholder_pdf),
                             daemon=True).start()
        else:
            self._show_attach_confirm_dialog(cby_pdfs_map, awb_bol_pdfs, pdfs,
                                             placeholder_pdf, placeholder_cbys)

    def _show_attach_confirm_dialog(self, cby_pdfs_map, awb_bol_pdfs, all_pdfs,
                                     placeholder_pdf=None, placeholder_cbys=None):
        """Show confirmation dialog with PDFs grouped by CBY before attaching."""
        if placeholder_cbys is None:
            placeholder_cbys = []

        dialog = ctk.CTkToplevel(self.win)
        dialog.title("Attach Supporting Documents")
        dialog.configure(fg_color="#1a1a2e")
        dialog.geometry("550x500")
        dialog.resizable(False, False)
        dialog.transient(self.win)
        dialog.grab_set()

        dialog.update_idletasks()
        px = self.win.winfo_rootx() + (self.win.winfo_width() - 550) // 2
        py = self.win.winfo_rooty() + (self.win.winfo_height() - 500) // 2
        dialog.geometry(f"+{px}+{py}")

        ctk.CTkLabel(dialog, text="Supporting Documents Found",
                     font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(16, 4))
        ctk.CTkLabel(dialog,
                     text=f"Found {len(all_pdfs)} PDF file(s).\n"
                          "Review the matching below, then click Start.",
                     font=(MODERN_FONT, 11), text_color="#888888").pack(pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent", height=320)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # Show AWB/BOL section
        if awb_bol_pdfs:
            ctk.CTkLabel(scroll, text="AWB / Bill of Lading (attached to all CBYs):",
                         font=(MODERN_FONT, 11, "bold"), text_color="#117a65").pack(anchor="w", pady=(4, 2))
            for pdf in awb_bol_pdfs:
                ctk.CTkLabel(scroll, text=f"  {pdf.name}",
                             font=(MODERN_FONT, 10), text_color="#e8e8e8").pack(anchor="w")
            ctk.CTkLabel(scroll, text="",
                         font=(MODERN_FONT, 8)).pack()

        # Show placeholder info
        if placeholder_pdf and placeholder_cbys:
            ctk.CTkLabel(scroll, text=f"Placeholder for CBYs with no invoice:",
                         font=(MODERN_FONT, 11, "bold"), text_color="#e8820c").pack(anchor="w", pady=(4, 2))
            ctk.CTkLabel(scroll, text=f"  {placeholder_pdf.name}",
                         font=(MODERN_FONT, 10), text_color="#e8e8e8").pack(anchor="w")
            ctk.CTkLabel(scroll, text=f"  Used for: {', '.join(sorted(placeholder_cbys))}",
                         font=(MODERN_FONT, 10), text_color="#888888").pack(anchor="w")
            ctk.CTkLabel(scroll, text="",
                         font=(MODERN_FONT, 8)).pack()

        # Show CBY sections
        if cby_pdfs_map:
            ctk.CTkLabel(scroll, text="Documents by CBY:",
                         font=(MODERN_FONT, 11, "bold"), text_color="#e8820c").pack(anchor="w", pady=(4, 2))
            for cby in sorted(cby_pdfs_map.keys()):
                is_placeholder = cby in placeholder_cbys
                label = f"  CBY {cby}:" + ("  [placeholder]" if is_placeholder else "")
                ctk.CTkLabel(scroll, text=label,
                             font=(MODERN_FONT, 10, "bold"), text_color="#e8e8e8").pack(anchor="w")
                for pdf in cby_pdfs_map[cby]:
                    ctk.CTkLabel(scroll, text=f"    {pdf.name}",
                                 font=(MODERN_FONT, 10), text_color="#888888").pack(anchor="w")

        # Show unmatched CBYs (no invoice AND no placeholder)
        tree_cbys = set()
        for item in self._tree.get_children():
            cby = self._tree.set(item, "cby")
            if cby and cby != "MASTER":
                tree_cbys.add(cby)
        unmatched = tree_cbys - set(cby_pdfs_map.keys())
        if unmatched:
            ctk.CTkLabel(scroll, text="",
                         font=(MODERN_FONT, 8)).pack()
            ctk.CTkLabel(scroll, text=f"CBYs with no PDFs at all: {', '.join(sorted(unmatched))}",
                         font=(MODERN_FONT, 10), text_color="#c0392b").pack(anchor="w")

        result = [None]
        def on_start():
            result[0] = True
            dialog.destroy()
        def on_cancel():
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=(0, 16))
        ctk.CTkButton(btn_frame, text="Start Attaching", command=on_start,
                      fg_color="#117a65", hover_color="#138d75", width=140, height=32,
                      corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel,
                      fg_color="#555555", hover_color="#666666", width=80, height=32,
                      corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        self.win.wait_window(dialog)

        if result[0]:
            self._attach_btn.configure(state="disabled")
            threading.Thread(target=self._attach_docs_worker,
                             args=(cby_pdfs_map, awb_bol_pdfs, placeholder_pdf),
                             daemon=True).start()

    def _find_supporting_pdfs(self):
        """Find PDFs in the manifest folder and subfolders.
        The manifest folder is either the XML folder itself or its parent
        (if the XML folder is a subfolder like 'XML Files')."""
        search_dirs = set()

        # Always search the XML folder itself
        search_dirs.add(self._xml_folder)

        # Search immediate subfolders of the XML folder
        try:
            for item in self._xml_folder.iterdir():
                if item.is_dir():
                    search_dirs.add(item)
        except Exception:
            pass

        # Always search the parent folder and its subfolders too.
        # The manifest PDF (Manifest and Customs*.pdf) typically lives
        # in the parent, and the XML folder may be named 'invoices',
        # 'XML files', or anything else.
        parent = self._xml_folder.parent
        if parent != self._xml_folder and parent.exists():
            search_dirs.add(parent)
            try:
                for item in parent.iterdir():
                    if item.is_dir():
                        search_dirs.add(item)
            except Exception:
                pass

        # Find all PDFs in the search dirs
        pdfs = set()
        for d in search_dirs:
            if d.exists():
                for pdf in d.glob("*.pdf"):
                    pdfs.add(pdf)
                for pdf in d.glob("*.PDF"):
                    pdfs.add(pdf)

        return sorted(pdfs, key=lambda p: p.name.lower())

    def _attach_docs_worker(self, cby_pdfs_map, awb_bol_pdfs, placeholder_pdf=None):
        """Background thread: attach supporting docs to each declaration on COLS."""
        self._uploading = True
        self._stop_requested = False
        self.win.after(0, lambda: self._pause_btn.configure(state="disabled"))

        self.win.after(0, lambda: self._update_status(
            "Looking for 'Upload Supporting Documents' links on COLS..."))

        try:
            links = self._driver.find_elements(
                By.XPATH, "//a[contains(text(), 'Upload Supporting Documents')]")
        except Exception as e:
            self.win.after(0, lambda: self._status.configure(
                text=f"Error finding links: {e}"))
            self.win.after(0, lambda: self._attach_btn.configure(state="normal"))
            self._uploading = False
            return

        if not links:
            self.win.after(0, lambda: messagebox.showwarning(
                "No Links Found",
                "Could not find any 'Upload Supporting Documents' links on the COLS page.\n\n"
                "Make sure you have uploaded XML files and they appear\n"
                "in the table on the COLS page."))
            self.win.after(0, lambda: self._status.configure(
                text="No 'Upload Supporting Documents' links found"))
            self.win.after(0, lambda: self._attach_btn.configure(state="normal"))
            self._uploading = False
            return

        attached_count = 0
        failed_count = 0
        skipped_count = 0
        # Only count visible links — hidden/stale links can cause the script
        # to keep trying past the real declarations.
        try:
            visible_links = [l for l in links if l.is_displayed()]
        except Exception:
            # Links went stale already — re-find them
            links = self._driver.find_elements(
                By.XPATH, "//a[contains(text(), 'Upload Supporting Documents')]")
            try:
                visible_links = [l for l in links if l.is_displayed()]
            except Exception:
                visible_links = links
        total_links = len(visible_links)
        consecutive_failures = 0
        MAX_CONSECUTIVE_FAILURES = 3

        _logging_enabled = getattr(self, '_logging_var', None) and self._logging_var.get()
        if _logging_enabled:
            log_path = Path.home() / "Documents" / "attach_docs_log.txt"
            try:
                with open(log_path, "a", encoding="utf-8") as logf:
                    logf.write(f"\n--- Attach session started ---\n")
                    logf.write(f"Total links found: {len(links)}  Visible: {total_links}\n")
                    logf.flush()
            except Exception as e:
                print(f"[Attach Docs] Logging setup error: {e}")
        else:
            import os
            log_path = os.devnull

        # Track processed links by their stable DOM id instead of position.
        # Oracle ADF re-stamps/recycles table rows as you scroll, so indexing
        # by position causes random skips. By id, we always pick the next
        # link we haven't done yet, regardless of how the list reshuffles.
        processed_link_ids = set()
        i = -1  # iteration counter (kept for status/error messages)

        # CBYs we've confirmed docs were attached to (persists across the
        # double-check pass). Used to skip re-doing ones already handled.
        attached_cbys = set()
        # Track CBYs we attempted but failed/skipped during the main pass.
        # If this is empty at the end, the double-check pass is skipped entirely
        # (no point re-walking everything if nothing was missed).
        missed_cbys = set()
        # After the main pass finishes, we do ONE final "double-check" pass to
        # catch any declaration that got skipped. It runs only once — some
        # genuinely can't attach (e.g. oversized PDF) and we don't want an
        # endless loop chasing them.
        doing_double_check = False

        while not self._stop_requested:
            # If too many consecutive failures, the page is likely in a bad state.
            # Stop rather than spamming clicks that open dialogs we can't read.
            if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                self.win.after(0, lambda: self._update_status(
                    f"Stopped after {consecutive_failures} consecutive failures — page may need manual review"))
                self.win.after(0, lambda: messagebox.showwarning(
                    "Attachment Halted",
                    f"The script stopped after {consecutive_failures} consecutive dialogs could not be read.\n\n"
                    f"This usually means the COLS page is in an unexpected state.\n"
                    f"Attached: {attached_count}  Failed: {failed_count}  Skipped: {skipped_count}\n\n"
                    f"Please check the COLS page, close any open dialogs, and try again."))
                break

            # ---- Find the next link we haven't processed yet ----
            def _find_next_link():
                """Return (link, link_id) for the first unprocessed visible link,
                or (None, None) if there are none currently in the DOM."""
                found = self._driver.find_elements(
                    By.XPATH, "//a[contains(text(), 'Upload Supporting Documents')]")
                for l in found:
                    try:
                        if not l.is_displayed():
                            continue
                        lid = l.get_attribute("id")
                    except Exception:
                        continue
                    if lid and lid not in processed_link_ids:
                        return l, lid
                return None, None

            link, link_id = _find_next_link()

            # If nothing unprocessed is currently rendered, scroll down to force
            # ADF to render more rows, then look again. If still nothing after
            # reaching the bottom, we're genuinely done.
            if link is None:
                prev_scroll = self._driver.execute_script("return window.pageYOffset;")
                self._driver.execute_script("window.scrollBy(0, 800);")
                time.sleep(1.2)
                link, link_id = _find_next_link()
                if link is None:
                    new_scroll = self._driver.execute_script("return window.pageYOffset;")
                    if new_scroll == prev_scroll:
                        # Bottom reached and no unprocessed links remain.
                        if not doing_double_check:
                            # Only run the double-check pass if the main pass
                            # actually missed something. If everything attached
                            # cleanly, skip the re-walk entirely — saves several
                            # minutes on good runs.
                            if not missed_cbys:
                                if _logging_enabled:
                                    with open(log_path, "a", encoding="utf-8") as logf:
                                        logf.write(f"\nAll links processed cleanly "
                                                   f"({len(attached_cbys)} attached) — "
                                                   f"skipping double-check (no misses).\n")
                                break
                            # Start the ONE-TIME double-check pass: scroll back to
                            # the top, forget which link ids we've seen (so we
                            # re-scan everything), and re-walk the list. The CBY
                            # skip-check below means already-attached declarations
                            # are closed immediately, so this only re-attaches misses.
                            doing_double_check = True
                            processed_link_ids = set()
                            consecutive_failures = 0
                            # Remember how many declarations the double-check pass
                            # will walk, so its own progress bar can fill 0 -> total.
                            double_check_total = max(total_links, 1)
                            self._driver.execute_script("window.scrollTo(0, 0);")
                            time.sleep(1.2)
                            self.win.after(0, lambda: self._update_status(
                                "Double-checking for any missed documents..."))
                            if self._autopilot_bar and self._autopilot_bar.is_alive():
                                self._autopilot_bar.set_phase("Double-checking Documents...")
                                self._autopilot_bar.set_progress(0, double_check_total)
                            if _logging_enabled:
                                with open(log_path, "a", encoding="utf-8") as logf:
                                    logf.write(f"\n--- Double-check pass started "
                                               f"({len(attached_cbys)} attached so far, "
                                               f"{len(missed_cbys)} missed) ---\n")
                            continue
                        # Double-check already done — we're truly finished.
                        if _logging_enabled:
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"\nAll links processed "
                                           f"({len(attached_cbys)} attached total).\n")
                        break
                    # Scrolled but nothing new yet — loop again to keep scrolling.
                    continue

            # Mark as processed up front so a failure on this link doesn't cause
            # us to retry it forever.
            processed_link_ids.add(link_id)
            i += 1
            # The initial scan may have under-counted (ADF renders rows lazily),
            # so keep the displayed total at least as large as what we've done.
            if len(processed_link_ids) > total_links:
                total_links = len(processed_link_ids)

            if doing_double_check:
                # During the double-check, drive the bar 0 -> double_check_total
                # so the always-on-top window visibly moves instead of sitting
                # frozen at a full yellow fill.
                dc_total = max(double_check_total, len(processed_link_ids))
                self.win.after(0, lambda n=len(processed_link_ids), t=dc_total: self._update_status(
                    f"Double-checking {n} of {t}..."))
                if self._autopilot_bar and self._autopilot_bar.is_alive():
                    self._autopilot_bar.set_progress(len(processed_link_ids), dc_total)
            else:
                self.win.after(0, lambda n=len(processed_link_ids), t=total_links: self._update_status(
                    f"Attaching docs to declaration {n} of {t}..."))
                if self._autopilot_bar and self._autopilot_bar.is_alive():
                    self._autopilot_bar.set_progress(len(processed_link_ids), total_links)

            try:
                # Before clicking, make sure no dialog is already open from a
                # previous iteration. If one is lingering, close it first —
                # this prevents multiple dialogs stacking up.
                try:
                    existing = self._driver.find_elements(
                        By.XPATH, "//div[contains(text(), 'Supporting Documents - HBL')]")
                    if any(d.is_displayed() for d in existing):
                        self._close_supporting_docs_dialog()
                        time.sleep(1)
                except Exception:
                    pass

                # Scroll the link into view
                self._driver.execute_script("arguments[0].scrollIntoView(true);", link)
                time.sleep(0.5)

                # Click the link
                link.click()
                time.sleep(2)

                # Wait for dialog to appear and read title
                title_text = ""
                for attempt in range(10):  # Try for up to 5 seconds
                    try:
                        title_el = self._driver.find_element(
                            By.XPATH, "//div[contains(text(), 'Supporting Documents - HBL')]")
                        title_text = title_el.text
                        if title_text:
                            break
                    except Exception:
                        pass
                    time.sleep(0.5)

                if not title_text:
                    try:
                        title_el = self._driver.find_element(
                            By.XPATH, "//div[contains(@class, 'x1h9') and contains(text(), 'HBL')]")
                        title_text = title_el.text
                    except Exception:
                        pass

                with open(log_path, "a", encoding="utf-8") as logf:
                    logf.write(f"\nDialog title: '{title_text}'\n")
                    if not title_text:
                        logf.write(f"  WARNING: Could not read dialog title!\n")
                        # Dump page state for debugging
                        logf.write(f"  Current URL: {self._driver.current_url}\n")
                        divs = self._driver.find_elements(
                            By.XPATH, "//div[contains(@class, 'x1h9')]")
                        logf.write(f"  Dialog-like divs found: {len(divs)}\n")
                        for di, d in enumerate(divs[:5]):
                            try:
                                logf.write(f"    [{di}] text='{d.text[:100]}'\n")
                            except:
                                pass

                # Extract CBY from title (e.g., "Supporting Documents - HBL-CBY 2006 1091363994.xml")
                cby_match = re.search(r'CBY\s*(\d+)', title_text)
                if not cby_match:
                    self.win.after(0, lambda t=title_text: self._status.configure(
                        text=f"Could not find CBY in dialog title: {t}"))
                    # Aggressively close whatever might be open — press Escape
                    # multiple times and try close buttons, then wait longer.
                    self._close_supporting_docs_dialog()
                    time.sleep(0.5)
                    self._close_supporting_docs_dialog()
                    time.sleep(2)
                    consecutive_failures += 1
                    failed_count += 1
                    if not doing_double_check:
                        missed_cbys.add("?")  # unknown CBY — forces double-check
                    continue

                # Reset consecutive failure counter on success
                consecutive_failures = 0
                cby = cby_match.group(1)

                # If this CBY already got its docs (earlier in this pass, or in
                # the main pass before the double-check), just close and move on.
                if cby in attached_cbys:
                    if _logging_enabled:
                        with open(log_path, "a", encoding="utf-8") as logf:
                            logf.write(f"\nCBY {cby}: already attached, skipping.\n")
                    self._close_supporting_docs_dialog()
                    time.sleep(0.5)
                    continue

                # Update treeview: mark this CBY as Attaching...
                self.win.after(0, lambda c=cby: self._set_row_docs_status(c, "docs_attach", "Attaching..."))

                # Find PDFs for this CBY
                cby_pdfs = cby_pdfs_map.get(cby, [])
                all_pdfs_for_cby = cby_pdfs + awb_bol_pdfs

                # If no invoices matched but we have a placeholder, use it
                if not cby_pdfs and placeholder_pdf:
                    all_pdfs_for_cby = [placeholder_pdf] + awb_bol_pdfs
                    self.win.after(0, lambda c=cby: self._status.configure(
                        text=f"Using placeholder for CBY {c} (no invoice found)..."))

                if not all_pdfs_for_cby:
                    self.win.after(0, lambda c=cby: self._status.configure(
                        text=f"No PDFs found for CBY {c}, skipping..."))
                    self.win.after(0, lambda c=cby: self._set_row_docs_status(c, "docs_manual", "No PDFs - Check Manually"))
                    if self._autopilot_bar and self._autopilot_bar.is_alive():
                        self._autopilot_bar.add_issue()
                    self._close_supporting_docs_dialog()
                    skipped_count += 1
                    if not doing_double_check:
                        missed_cbys.add(cby)
                    time.sleep(1)
                    continue

                self.win.after(0, lambda c=cby, n=len(all_pdfs_for_cby): self._status.configure(
                    text=f"Attaching {n} document(s) to CBY {c}..."))

                if _logging_enabled:
                    with open(log_path, "a", encoding="utf-8") as logf:
                        logf.write(f"\n{'='*60}\n")
                        logf.write(f"CBY: {cby} | PDFs: {len(all_pdfs_for_cby)}\n")
                        for pdf in all_pdfs_for_cby:
                            logf.write(f"  - {pdf.name}\n")

                # Separate AWB/BOL from invoices
                cby_awb = [pdf for pdf in all_pdfs_for_cby
                           if pdf.name.upper().lstrip("#").startswith("378")
                           or pdf.name.upper().startswith("SMLU")
                           or pdf.name.upper().startswith("#SMLU")]
                cby_invoices = [pdf for pdf in all_pdfs_for_cby if pdf not in cby_awb]
                ordered_pdfs = cby_awb + cby_invoices

                with open(log_path, "a", encoding="utf-8") as logf:
                    logf.write(f"Ordered: {[p.name for p in ordered_pdfs]}\n")

                attach_failed = False
                for j, pdf in enumerate(ordered_pdfs):
                    if self._stop_requested:
                        break

                    with open(log_path, "a", encoding="utf-8") as logf:
                        logf.write(f"\n--- Field {j}: {pdf.name} ---\n")

                    # Step 1: If beyond the 2 default fields, click "Add Document"
                    if j >= 2:
                        with open(log_path, "a", encoding="utf-8") as logf:
                            logf.write(f"  Clicking 'Add Document' for extra field...\n")
                        try:
                            # Select 'Invoice' doc type BEFORE clicking Add Document
                            # (Diagnostic: user selects type first, then clicks Add)
                            from selenium.webdriver.support.ui import Select
                            dropdowns = self._driver.find_elements(
                                By.XPATH, "//select[contains(@id, 'soc1')]")
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  Found {len(dropdowns)} doc-type dropdowns\n")
                            if dropdowns:
                                # Select by visible text 'Invoice' first, fall back to value
                                try:
                                    Select(dropdowns[-1]).select_by_visible_text("Invoice")
                                except Exception:
                                    Select(dropdowns[-1]).select_by_value("1")
                                time.sleep(0.5)
                                with open(log_path, "a", encoding="utf-8") as logf:
                                    logf.write("  Selected Invoice doc type\n")
                            add_btn = self._driver.find_element(
                                By.XPATH, "//button[contains(text(), 'Add Document')]")
                            # Count existing file inputs before clicking
                            before = len(self._driver.find_elements(
                                By.XPATH, "//input[@type='file']"))
                            add_btn.click()
                            # Wait up to 5 s for a new file input to appear
                            for _aw in range(10):
                                time.sleep(0.5)
                                after = len(self._driver.find_elements(
                                    By.XPATH, "//input[@type='file']"))
                                if after > before:
                                    break
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  Add Document clicked: inputs {before}->{after}\n")
                        except Exception as e:
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  ERROR adding field: {e}\n")

                    # Step 2: RE-FIND dialog inputs fresh each time.
                    # ADF removes a filled input from the DOM after each attachment,
                    # so the next empty input is ALWAYS at index 0.
                    attached_ok = False
                    for retry in range(3):
                        try:
                            all_inputs = self._driver.find_elements(
                                By.XPATH, "//input[@type='file']")
                            dialog_inputs = [fi for fi in all_inputs
                                           if "if3" in (fi.get_attribute("id") or "")]
                            # Fallback: 'Add Document' new row may use different ID pattern
                            if not dialog_inputs:
                                dialog_inputs = all_inputs
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  Attempt {retry+1}: All={len(all_inputs)}, "
                                          f"Dialog={len(dialog_inputs)}\n")
                                for di, fi in enumerate(dialog_inputs):
                                    logf.write(f"    dialog_input[{di}]: "
                                              f"id={fi.get_attribute('id')}\n")

                            # Always use index 0 — ADF removes filled inputs from DOM
                            if len(dialog_inputs) > 0:
                                dialog_inputs[0].send_keys(str(pdf.resolve()))
                                time.sleep(3)  # Wait for ADF to re-render after attach
                                # Verify: check if filename appeared in dialog text.
                                # ADF shows the filename after successful attach.
                                # If the browser froze or attach failed silently,
                                # the filename won't be there — retry.
                                try:
                                    dialog_text = self._driver.find_element(
                                        By.XPATH, "//div[contains(@id,':d1')]").text
                                    pdf_stem = pdf.stem
                                    if pdf_stem in dialog_text or pdf.name in dialog_text:
                                        with open(log_path, "a", encoding="utf-8") as logf:
                                            logf.write(f"  Attached {pdf.name} to dialog_input[0] (verified)\n")
                                        attached_ok = True
                                        break
                                    else:
                                        with open(log_path, "a", encoding="utf-8") as logf:
                                            logf.write(f"  Retry {retry+1}: {pdf.name} not in dialog text yet, retrying...\n")
                                        time.sleep(1)
                                except Exception:
                                    # Can't read dialog text — assume success
                                    with open(log_path, "a", encoding="utf-8") as logf:
                                        logf.write(f"  Attached {pdf.name} to dialog_input[0] (no verify)\n")
                                    attached_ok = True
                                    break
                            else:
                                with open(log_path, "a", encoding="utf-8") as logf:
                                    logf.write(f"  Retry {retry+1}: no dialog inputs available\n")
                                time.sleep(1)
                        except Exception as e:
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  Retry {retry+1} error: {type(e).__name__}: {e}\n")
                            time.sleep(1)

                    if not attached_ok:
                        with open(log_path, "a", encoding="utf-8") as logf:
                            logf.write(f"  FAILED to attach {pdf.name} after 3 retries\n")
                        attach_failed = True
                        self.win.after(0, lambda n=pdf.name: self._status.configure(
                            text=f"Warning: could not attach {n}"))
                        self.win.after(0, lambda c=cby: self._set_row_docs_status(c, "docs_manual", "Check Manually"))
                        if self._autopilot_bar and self._autopilot_bar.is_alive():
                            self._autopilot_bar.add_issue()
                        if not doing_double_check:
                            missed_cbys.add(cby)

                # Step 3: Click the real 'Upload' <A> link in the dialog.
                # Diagnostic confirmed: Upload is an <A> inside a div whose id
                # ends in ':d1_ok'. Must use real Selenium .click() so ADF's
                # event handlers fire — JS .click() on upDecl bypasses them.
                upload_clicked = False
                if not attach_failed:
                    # Primary: find the A link inside the *:d1_ok footer div
                    for _w in range(20):
                        try:
                            upload_link = self._driver.find_element(
                                By.XPATH,
                                "//div[contains(@id,':d1_ok')]//a")
                            upload_link.click()
                            upload_clicked = True
                            time.sleep(3)
                            break
                        except Exception:
                            time.sleep(0.5)
                    # Fallback: find any <A> wrapping a SPAN with text 'Upload'
                    if not upload_clicked:
                        with open(log_path, "a", encoding="utf-8") as logf:
                            logf.write("  d1_ok link not found, trying span fallback...\n")
                        try:
                            upload_link = self._driver.find_element(
                                By.XPATH,
                                "//a[.//span[normalize-space(text())='Upload']]")
                            upload_link.click()
                            upload_clicked = True
                            time.sleep(3)
                        except Exception as e:
                            with open(log_path, "a", encoding="utf-8") as logf:
                                logf.write(f"  Span fallback failed: {e}\n")

                if not upload_clicked:
                    with open(log_path, "a", encoding="utf-8") as logf:
                        logf.write(f"\nWARNING: Could not find/click Upload link!\n")
                    self.win.after(0, lambda: self._status.configure(
                        text="Warning: could not find Upload link, skipping..."))
                    self.win.after(0, lambda c=cby: self._set_row_docs_status(c, "docs_manual", "Check Manually"))
                    self._close_supporting_docs_dialog()
                    if not doing_double_check:
                        missed_cbys.add(cby)
                    time.sleep(1)
                else:
                    with open(log_path, "a", encoding="utf-8") as logf:
                        logf.write(f"Upload clicked successfully for CBY {cby}\n")
                    # This CBY is now genuinely done — remember it so the
                    # double-check pass closes it immediately instead of redoing.
                    attached_cbys.add(cby)
                    # Wait for COLS to close the dialog automatically after upload.
                    # Poll for the dialog to disappear (max 10s), then proceed.
                    for _dw in range(20):
                        time.sleep(0.5)
                        try:
                            dialogs = self._driver.find_elements(
                                By.XPATH, "//div[contains(text(), 'Supporting Documents - HBL')]")
                            if not dialogs or not any(d.is_displayed() for d in dialogs):
                                break  # dialog is gone
                        except Exception:
                            break  # page changed, dialog likely gone

                attached_count += 1
                # Reclassify false "error" rows as "uploaded" — if docs were
                # attached, the file IS on COLS, so the error was a false positive.
                self.win.after(0, lambda c=cby: self._reclassify_error_if_docs_attached(c))
                self.win.after(0, lambda c=cby, n=len(all_pdfs_for_cby): self._update_status(
                    f"Attached {n} document(s) to CBY {c} ({attached_count} done)"))
                self.win.after(0, lambda c=cby: self._set_row_docs_status(c, "docs_ok", "Attached"))

            except Exception as e:
                failed_count += 1
                err_msg = f"{type(e).__name__}: {e}"
                self.win.after(0, lambda msg=err_msg: self._status.configure(
                    text=f"Failed (#{i+1}): {msg}"))
                print(f"[Attach Docs] Failed on link {i+1}: {err_msg}")
                if not doing_double_check:
                    missed_cbys.add("?")  # unknown — forces double-check
                try:
                    self._close_supporting_docs_dialog()
                except Exception:
                    pass
                time.sleep(2)

        # Show summary
        self._uploading = False
        self.win.after(0, lambda: self._attach_btn.configure(state="normal"))
        # Re-enable pause button if upload is paused
        if self._paused:
            self.win.after(0, lambda: self._pause_btn.configure(state="normal"))
        # Update autopilot bar to complete state
        if self._autopilot_bar and self._autopilot_bar.is_alive():
            docs_issues = failed_count + skipped_count
            # Count upload results from treeview
            up = err = miss = 0
            for item in self._tree.get_children():
                s = self._tree.set(item, "status").lower()
                if s == "uploaded": up += 1
                elif s == "error": err += 1
                elif s == "missed": miss += 1
            self._autopilot_bar.complete(up, err, miss, attached_count, docs_issues)
        self.win.after(0, lambda: self._update_status(
            f"Done: {attached_count} declarations had docs attached, "
                 f"{skipped_count} skipped, {failed_count} failed"))
        self.win.after(0, lambda: messagebox.showinfo(
            "Attach Supporting Docs Complete",
            f"Attached supporting documents to {attached_count} declaration(s).\n"
            f"Skipped: {skipped_count} (no matching PDFs)\n"
            f"Failed: {failed_count}\n\n"
            f"Review the COLS page to verify the documents are attached,\n"
            f"then check the declaration checkbox and click Submit when ready."))

    def _close_supporting_docs_dialog(self):
        """Close the Supporting Documents dialog on COLS."""
        # Strategy 1: Find close button by aria-label
        for xpath in [
            "//a[@aria-label='Close' and contains(@id, 'd1::close')]",
            "//a[@title='Close' and contains(@id, 'd1::close')]",
            "//a[@aria-label='Close']",
            "//a[contains(@class, 'x1h3')]",
            "//a[contains(@id, '::close')]",
        ]:
            try:
                close_btn = self._driver.find_element(By.XPATH, xpath)
                if close_btn.is_displayed():
                    close_btn.click()
                    time.sleep(1.5)
                    return
            except Exception:
                continue
        # Strategy 2: Press Escape (try multiple times)
        try:
            from selenium.webdriver.common.keys import Keys
            body = self._driver.find_element(By.TAG_NAME, "body")
            for _ in range(3):
                body.send_keys(Keys.ESCAPE)
                time.sleep(0.5)
        except Exception:
            pass

    def _reclassify_error_if_docs_attached(self, cby):
        """If a row is marked 'error' but docs were just attached to it on COLS,
        the file was actually uploaded successfully (false error). Reclassify
        as 'uploaded' so the copy feature can capture its declaration number."""
        for item in self._tree.get_children():
            row_cby = self._tree.set(item, "cby")
            row_status = self._tree.set(item, "status").lower()
            if row_cby == cby and row_status == "error":
                self._tree.set(item, "status", "Uploaded")
                self._tree.item(item, tags=("uploaded",))
                self._update_status(f"CBY {cby} reclassified: error → uploaded (docs attached)")
                break

    def _copy_decl_from_cols(self):
        """Read the COLS page for declaration numbers and filenames.
        COLS uses Oracle ADF with absolutely-positioned divs, not HTML tables.
        The _afrc attribute encodes column/row positions:
          _afrc="4 1 X 1..." = Declaration No column
          _afrc="6 1 X 1..." = Filename column
        We pair them by row number (X)."""
        if not self._driver:
            messagebox.showwarning("No Browser", "Launch Edge first, then use this button.")
            return

        self._status.configure(text="Reading declaration numbers from COLS page...")

        pairs = []  # list of (declaration_number, filename) tuples

        try:
            # Use JavaScript to extract data from the ADF grid
            # The _afrc attribute format is "col 1 row 1 start top"
            # Column 4 = Declaration No, Column 6 = Filename
            js_script = """
            var results = [];
            // Find all divs with _afrc attribute
            var allDivs = document.querySelectorAll('div[_afrc]');
            var declCells = {};
            var fileCells = {};

            for (var i = 0; i < allDivs.length; i++) {
                var div = allDivs[i];
                var afrc = div.getAttribute('_afrc');
                if (!afrc) continue;
                var parts = afrc.split(' ');
                if (parts.length < 2) continue;
                var col = parts[0];
                var row = parts[2];

                var text = div.textContent.trim();

                if (col === '4' && parseInt(row) > 2) {
                    // Declaration No column, skip header row (row 2)
                    declCells[row] = text;
                } else if (col === '6' && parseInt(row) > 2) {
                    // Filename column, skip header row
                    fileCells[row] = text;
                }
            }

            // Pair declaration numbers with filenames by row
            for (var row in fileCells) {
                var filename = fileCells[row];
                var declNum = declCells[row] || '-';
                if (filename && filename.indexOf('HBL-') >= 0) {
                    results.push([declNum, filename]);
                }
            }
            return results;
            """
            js_results = self._driver.execute_script(js_script)

            for item in js_results:
                decl_num = str(item[0]).strip()
                filename = str(item[1]).strip()
                # Skip rows where declaration number is "-" (not yet submitted)
                if decl_num and decl_num != "-" and filename:
                    pairs.append((decl_num, filename))
                elif filename:
                    # Include with empty decl number so user can see it and fill in
                    pairs.append(("", filename))

        except Exception as e:
            # Fallback: try reading by position (left:160px = decl, left:345px = filename)
            try:
                js_fallback = """
                var results = [];
                var spans = document.querySelectorAll('span');
                var byTop = {};
                for (var i = 0; i < spans.length; i++) {
                    var span = spans[i];
                    var text = span.textContent.trim();
                    if (!text) continue;
                    var parent = span.parentElement;
                    if (!parent) continue;
                    var left = parent.style.left;
                    var top = parent.style.top;
                    if (!left || !top) continue;
                    if (!byTop[top]) byTop[top] = {};
                    if (left === '160px') byTop[top]['decl'] = text;
                    if (left === '345px') byTop[top]['file'] = text;
                }
                for (var top in byTop) {
                    var entry = byTop[top];
                    if (entry['file'] && entry['file'].indexOf('HBL-') >= 0) {
                        results.push([entry['decl'] || '', entry['file']]);
                    }
                }
                return results;
                """
                js_results = self._driver.execute_script(js_fallback)
                for item in js_results:
                    decl_num = str(item[0]).strip()
                    filename = str(item[1]).strip()
                    if decl_num and decl_num != "-":
                        pairs.append((decl_num, filename))
                    elif filename:
                        pairs.append(("", filename))
            except Exception as e2:
                messagebox.showerror("Error", f"Could not read COLS page:\n{e2}")
                self._status.configure(text="Failed to read COLS page")
                return

        if not pairs:
            messagebox.showwarning("No Declaration Numbers Found",
                                   "Could not find any declaration data on the COLS page.\n\n"
                                   "Make sure you have uploaded files and they appear\n"
                                   "in the table on the COLS page.\n\n"
                                   "You can also enter them manually by double-clicking\n"
                                   "the Declaration # column in the uploader window.")
            self._status.configure(text="No declaration data found on COLS page")
            return

        # Show what we found and let user confirm
        result = self._show_decl_capture_dialog(pairs)
        if result is None:
            self._status.configure(text="Declaration number capture cancelled")
            return

        # Match declaration numbers to treeview rows by filename
        matched = 0
        for decl_num, col_filename in result:
            if not decl_num or not col_filename:
                continue
            for item in self._tree.get_children():
                tree_file = self._tree.set(item, "file")
                tree_status = self._tree.set(item, "status").lower()
                # Match by exact filename or by CBY number
                if tree_file == col_filename or col_filename in tree_file or tree_file in col_filename:
                    if tree_status in ("uploaded", "error"):
                        self._tree.set(item, "decl", decl_num)
                        self._decl_numbers[tree_file] = decl_num
                        matched += 1
                    break
                tree_cby = self._tree.set(item, "cby")
                if tree_cby and tree_cby != "MASTER":
                    col_cby_match = re.search(r'CBY\s*(\d+)', col_filename, re.IGNORECASE)
                    if col_cby_match and col_cby_match.group(1) == tree_cby:
                        if tree_status in ("uploaded", "error"):
                            self._tree.set(item, "decl", decl_num)
                            self._decl_numbers[tree_file] = decl_num
                            matched += 1
                        break

        self._status.configure(
            text=f"Copied {len(result)} declaration number(s) from COLS. Matched {matched} to files.")
        if matched > 0:
            messagebox.showinfo("Success",
                                f"Matched {matched} declaration number(s) to uploaded files.\n\n"
                                f"Review the Declaration # column to verify.\n"
                                f"Double-click any cell to edit manually.")
        else:
            messagebox.showwarning("No Matches",
                                f"Found {len(result)} declaration number(s) on COLS,\n"
                                f"but could not match them to files by filename.\n\n"
                                f"The filenames on COLS may differ from the XML filenames.\n"
                                f"Double-click the Declaration # column to enter them manually.")

    def _show_decl_capture_dialog(self, pairs):
        """Show a dialog with declaration numbers and filenames for user to confirm/edit.
        pairs is a list of (declaration_number, filename) tuples."""
        dialog = ctk.CTkToplevel(self.win)
        dialog.title("Declaration Numbers from COLS")
        dialog.configure(fg_color="#1a1a2e")
        dialog.geometry("550x450")
        dialog.resizable(False, False)
        dialog.transient(self.win)
        dialog.grab_set()

        dialog.update_idletasks()
        px = self.win.winfo_rootx() + (self.win.winfo_width() - 550) // 2
        py = self.win.winfo_rooty() + (self.win.winfo_height() - 450) // 2
        dialog.geometry(f"+{px}+{py}")

        ctk.CTkLabel(dialog, text="Declaration Numbers Found on COLS",
                     font=(MODERN_FONT, 14, "bold"), text_color="#e8e8e8").pack(pady=(16, 4))
        ctk.CTkLabel(dialog,
                     text="Review the declaration numbers and filenames below.\n"
                          "Edit if needed. They will be matched to uploaded files.",
                     font=(MODERN_FONT, 11), text_color="#888888").pack(pady=(0, 12))

        # Column headers
        header_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkLabel(header_frame, text="Declaration Number",
                     font=(MODERN_FONT, 11, "bold"), text_color="#e8820c",
                     width=140, anchor="w").pack(side="left", padx=(0, 8))
        ctk.CTkLabel(header_frame, text="Filename",
                     font=(MODERN_FONT, 11, "bold"), text_color="#888888",
                     width=340, anchor="w").pack(side="left")

        # Scrollable frame for entries
        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent", height=260)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        decl_entries = []
        file_entries = []
        for decl_num, filename in pairs:
            row_frame = ctk.CTkFrame(scroll, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)

            de = ctk.CTkEntry(row_frame, width=140, height=28,
                              fg_color="#0f0f1a", border_color="#e8820c", border_width=1,
                              corner_radius=4, text_color="#e8e8e8",
                              font=(MODERN_FONT, 12))
            de.insert(0, decl_num)
            de.pack(side="left", padx=(0, 8))
            decl_entries.append(de)

            fe = ctk.CTkEntry(row_frame, width=340, height=28,
                              fg_color="#0f0f1a", border_color="#333", border_width=1,
                              corner_radius=4, text_color="#e8e8e8",
                              font=(MODERN_FONT, 11))
            fe.insert(0, filename)
            fe.pack(side="left")
            file_entries.append(fe)

        result = [None]
        def on_ok():
            result[0] = [(de.get().strip(), fe.get().strip())
                         for de, fe in zip(decl_entries, file_entries)
                         if de.get().strip()]
            dialog.destroy()
        def on_cancel():
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=(0, 16))
        ctk.CTkButton(btn_frame, text="Apply to Uploader Window", command=on_ok,
                      fg_color="#2e8b57", hover_color="#3cb371", width=180, height=32,
                      corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel,
                      fg_color="#555555", hover_color="#666666", width=80, height=32,
                      corner_radius=6, font=(MODERN_FONT, 12, "bold")).pack(side="left", padx=6)

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        self.win.wait_window(dialog)
        return result[0]

    def _quick_paste(self):
        """Open a manifest file, find the COLS sheet, match CBY numbers,
        and paste declaration numbers into the declaration column.

        Reads with openpyxl (data_only) to find columns and CBY values,
        then edits the sheet XML directly inside the xlsx zip to write
        declaration values. This preserves all external links, drawings,
        and formulas — no Excel process needed, no corruption.
        """
        if not self._decl_numbers:
            # They haven't copied/uploaded declaration numbers yet.
            # Replicate the "Copy Decl #s from COLS" flow automatically
            # so they see the same review dialog before pasting.
            proceed = messagebox.askyesno(
                "Copy Declaration Numbers First",
                "You haven't copied any declaration numbers yet.\n\n"
                "Click Yes to copy them from the COLS page now\n"
                "(same as clicking 'Copy Decl #s from COLS').\n\n"
                "You'll be able to review and confirm the list before\n"
                "anything is pasted to the manifest.")
            if not proceed:
                return
            self._copy_decl_from_cols()
            if not self._decl_numbers:
                # Copy step already showed its own message (no browser,
                # nothing found, or the user cancelled the review dialog)
                return

        file_path = filedialog.askopenfilename(
            title="Choose manifest file",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        # Build CBY -> decl map
        cby_to_decl = {}
        master_decl = None
        for filename, decl_num in self._decl_numbers.items():
            m = re.search(r'CBY\s*(\d+)', filename, re.IGNORECASE)
            if m and decl_num:
                cby_to_decl[m.group(1)] = decl_num
            elif ("HBL-Master" in filename or "MBL" in filename) and decl_num:
                master_decl = decl_num

        if not cby_to_decl and not master_decl:
            # Build a diagnostic dump of exactly what WAS captured so the
            # user can hit 'Report Issue' and the developer sees the guts
            # of why nothing matched — no guessing over the phone.
            diag = "Captured declaration numbers (filename -> decl#):\n"
            if self._decl_numbers:
                for fn, dn in self._decl_numbers.items():
                    diag += f"  {fn!r} -> {dn!r}\n"
            else:
                diag += "  (none)\n"
            diag += f"\nCBY regex used: CBY\\s*(\\d+)  (case-insensitive)\n"
            diag += f"Master patterns: 'HBL-Master' or 'MBL' in filename\n"
            diag += f"cby_to_decl result: {cby_to_decl}\n"
            diag += f"master_decl result: {master_decl!r}"
            _show_error_with_report("No Matches",
                "No declaration numbers with matching CBY numbers found.\n\n"
                "The uploaded files don't have CBY numbers in their names,\n"
                "or no house declarations were captured.\n\n"
                "Make sure the house XML files (HBL-CBY*.xml) were uploaded\n"
                "and their declaration numbers appear in the Declaration # column.",
                traceback_text=diag,
                window_name=self._window_name)
            return
        if not cby_to_decl and master_decl:
            diag = "Captured declaration numbers (filename -> decl#):\n"
            for fn, dn in self._decl_numbers.items():
                diag += f"  {fn!r} -> {dn!r}\n"
            diag += f"\nCBY regex used: CBY\\s*(\\d+)  (case-insensitive)\n"
            diag += f"cby_to_decl result: {cby_to_decl}\n"
            diag += f"master_decl result: {master_decl!r}"
            _show_error_with_report("No House Declarations",
                "Only a Master declaration was found — no House/CBY declarations.\n\n"
                "The Paste to Manifest feature needs house declaration numbers\n"
                "(from HBL-CBY*.xml files) to match against CBY rows in the manifest.\n\n"
                "Make sure the house XML files were uploaded and their declaration\n"
                "numbers appear in the Declaration # column.\n\n"
                f"Master declaration found: {master_decl}",
                traceback_text=diag,
                window_name=self._window_name)
            return

        try:
            import openpyxl
            import zipfile
            import shutil
            import xml.etree.ElementTree as ET
            from io import BytesIO
        except ImportError as e:
            messagebox.showerror("Error", f"Missing required library: {e}")
            return

        file_path = Path(file_path)

        # Check if the file is open in Excel (Windows locks open files)
        try:
            # Try opening for read+write AND renaming a temp file over it
            # (Excel may allow read-write but block file replacement)
            with open(str(file_path), 'r+b'):
                pass
        except PermissionError:
            self._show_manifest_open_error(file_path)
            return

        try:
            # ── Step 1: Read with openpyxl to find columns and CBY values ──────
            wb = openpyxl.load_workbook(str(file_path), data_only=True)

            cols_sheet_name = None
            for name in wb.sheetnames:
                if name.strip().upper().replace(" ", "") == "COLS":
                    cols_sheet_name = name
                    break
            if not cols_sheet_name:
                for name in wb.sheetnames:
                    if "cols" in name.strip().lower():
                        cols_sheet_name = name
                        break
            if not cols_sheet_name:
                messagebox.showerror("Sheet Not Found",
                                     f"Could not find a 'COLS' sheet.\n"
                                     f"Available: {', '.join(wb.sheetnames)}")
                return

            ws = wb[cols_sheet_name]

            # Find CBY and Declaration columns.
            # CBY header is unique to the data header row (not in the top section).
            # Once we find the CBY row, search that SAME row for the Decl column.
            # This avoids picking up 'COLS Declaration #' from the header section.
            cby_col = None
            decl_col = None
            header_row = None
            for row in ws.iter_rows(min_row=1, max_row=15):
                for cell in row:
                    val = str(cell.value).strip().lower() if cell.value else ""
                    if "cby" in val and cby_col is None:
                        cby_col = cell.column
                        header_row = cell.row
                        break
                if cby_col:
                    break

            if cby_col and header_row:
                # Search the SAME row for the Decl column
                for cell in ws[header_row]:
                    val = str(cell.value).strip().lower() if cell.value else ""
                    if "decl" in val and decl_col is None:
                        decl_col = cell.column

            if not cby_col:
                messagebox.showerror("Column Not Found",
                                     "Could not find a CBY column in the COLS sheet.")
                return
            if not decl_col:
                messagebox.showerror("Column Not Found",
                                     "Could not find a Declaration column in the COLS sheet.")
                return

            # Collect all writes: (row, col, value)
            writes = []
            pasted = 0

            # Master declaration into header section (rows 1-8)
            if self._master_decl:
                for row in ws.iter_rows(min_row=1, max_row=8):
                    for cell in row:
                        val = str(cell.value).strip().lower() if cell.value else ""
                        if ("master decl" in val or "cols declaration" in val
                                or ("declaration" in val and "#" in val
                                    and cell.row <= 5)):
                            target = ws.cell(row=cell.row, column=cell.column + 1)
                            if not target.value:
                                try:
                                    mval = int(self._master_decl)
                                except (ValueError, TypeError):
                                    break  # skip master, don't write non-numeric
                                writes.append((cell.row, cell.column + 1, mval))
                                pasted += 1
                            break
                    else:
                        continue
                    break

            # House declarations
            for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
                cby_cell = row[cby_col - 1] if cby_col <= len(row) else None
                if cby_cell and cby_cell.value:
                    cby_raw = cby_cell.value
                    if isinstance(cby_raw, float) and cby_raw == int(cby_raw):
                        cby_val = str(int(cby_raw))
                    else:
                        cby_val = str(cby_raw).strip()
                    decl = cby_to_decl.get(cby_val)
                    if not decl:
                        decl = cby_to_decl.get(cby_val.lstrip("0"))
                    if not decl:
                        for k, v in cby_to_decl.items():
                            if k.lstrip("0") == cby_val.lstrip("0"):
                                decl = v
                                break
                    if decl:
                        decl_cell = row[decl_col - 1] if decl_col <= len(row) else None
                        if decl_cell and not decl_cell.value:
                            # Only write numeric declaration values — the
                            # zip-editing method writes <v>...</v> which is
                            # only valid for numbers.  Non-numeric values
                            # would produce invalid Excel XML and corrupt
                            # the file.  Skip them instead.
                            try:
                                write_val = int(decl)
                            except (ValueError, TypeError):
                                continue  # skip this cell, don't write
                            writes.append((decl_cell.row, decl_cell.column, write_val))
                            pasted += 1

            wb.close()

            if not writes:
                messagebox.showinfo("Quick Paste",
                                    "No empty declaration cells found to paste into.\n"
                                    "All matching cells may already have values.")
                return

            # ── Step 2: Edit the sheet XML directly inside the xlsx zip ────────
            # This preserves everything — external links, drawings, formulas.
            try:
                success = self._write_cells_to_xlsx(str(file_path), cols_sheet_name,
                                                    writes, cby_col, decl_col, header_row)
            except PermissionError:
                self._show_manifest_open_error(file_path)
                return

            if success:
                messagebox.showinfo("Quick Paste Complete",
                                    f"Pasted {pasted} declaration number(s) into:\n"
                                    f"{file_path.name}\n"
                                    f"Sheet: {cols_sheet_name}\n\n"
                                    f"CBYs matched: {', '.join(sorted(cby_to_decl.keys()))}")
                self._status.configure(
                    text=f"Quick Paste: {pasted} declarations pasted to {file_path.name}")
            else:
                # _write_cells_to_xlsx returned False without raising —
                # this means it couldn't find the sheet inside the xlsx zip
                # structure.  Don't attempt an openpyxl save fallback — that
                # can corrupt the manifest by stripping external links and
                # drawings.  Instead, report the issue so it can be fixed.
                raise RuntimeError(
                    f"Could not locate the sheet '{cols_sheet_name}' inside the\n"
                    f"xlsx zip structure. The manifest was not modified.")

        except Exception as e:
            if "permission" in str(e).lower() or isinstance(e, PermissionError):
                self._show_manifest_open_error(file_path)
            else:
                # Real error — show with Report button so the developer
                # gets the full traceback in Discord.
                messagebox.showerror("Error", f"Failed to paste declarations:\n{e}")

    def _show_manifest_open_error(self, file_path):
        """Show a prominent, unmissable error that the manifest file is open
        or otherwise locked / not writable."""
        file_name = file_path.name if hasattr(file_path, 'name') else str(file_path)
        # Use a custom Toplevel dialog so it's large, bold, and stays on top
        dlg = ctk.CTkToplevel()
        dlg.title("CANNOT WRITE TO MANIFEST FILE")
        dlg.configure(fg_color="#0f1117")
        dlg.geometry("520x340")
        dlg.attributes("-topmost", True)
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"+{int((sw-520)/2)}+{int((sh-340)/2)}")
        dlg.transient()
        dlg.grab_set()
        # Red warning header
        ctk.CTkLabel(dlg, text="CANNOT WRITE TO MANIFEST",
                     font=(MODERN_FONT, 20, "bold"),
                     text_color="#e74c3c").pack(pady=(24, 8))
        ctk.CTkLabel(dlg,
                     text=f"The manifest file could not be written to:\n\n"
                          f"  {file_name}\n\n"
                          f"Possible causes:\n"
                          f"  1. The file is OPEN in Excel — close it and try again\n"
                          f"  2. The file is on a read-only or network folder\n"
                          f"  3. Your user account lacks write permission\n"
                          f"  4. Antivirus is blocking the write\n\n"
                          f"If closing Excel doesn't help, try saving a copy\n"
                          f"to your Desktop and pasting to that instead.",
                     font=(MODERN_FONT, 12),
                     text_color="#e8e8e8", justify="left").pack(pady=(0, 16))
        ctk.CTkButton(dlg, text="OK", command=dlg.destroy,
                      fg_color="#e74c3c", hover_color="#c0392b",
                      width=120, height=36, corner_radius=6,
                      font=(MODERN_FONT, 13, "bold")).pack(pady=(0, 24))
        dlg.bind("<Return>", lambda e: dlg.destroy())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        self._status.configure(
            text=f"Quick Paste FAILED - close {file_name} in Excel and try again")

    def _write_cells_to_xlsx(self, file_path, sheet_name, writes, cby_col, decl_col, header_row):
        """Write cell values directly into the xlsx zip using string replacement.
        Preserves everything — external links, drawings, formulas, namespaces.
        No Excel process, no ElementTree (which corrupts namespace prefixes).
        Returns True on success, False on failure."""
        import zipfile
        import re
        import shutil
        import os

        def col_letter(n):
            result = ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                result = chr(65 + rem) + result
            return result

        # Build cell_ref -> value string map
        cell_writes = {}
        for row, col, val in writes:
            ref = f"{col_letter(col)}{row}"
            cell_writes[ref] = str(val)

        temp_path = file_path + ".tmp"

        try:
            # Step 1: Find which sheet file inside the zip is our COLS sheet
            with zipfile.ZipFile(file_path, 'r') as zin:
                wb_xml = zin.read('xl/workbook.xml').decode('utf-8')

                # Find sheet rId by name (string search, no XML parsing)
                # Pattern: <sheet name=" COLS" ... r:id="rId2"/>
                sheet_rid = None
                # Try exact name match first, then flexible
                for pattern in [
                    rf'<sheet name="{re.escape(sheet_name)}"[^>]*r:id="(rId\d+)"',
                    rf'<sheet name="\s*{re.escape(sheet_name.strip())}"[^>]*r:id="(rId\d+)"',
                ]:
                    m = re.search(pattern, wb_xml)
                    if m:
                        sheet_rid = m.group(1)
                        break

                if not sheet_rid:
                    # Fallback: find any sheet with "cols" in the name
                    m = re.search(r'<sheet name="[^"]*[Cc][Oo][Ll][Ss][^"]*"[^>]*r:id="(rId\d+)"', wb_xml)
                    if m:
                        sheet_rid = m.group(1)

                if not sheet_rid:
                    return False

                # Find sheet file path from rels
                rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                m = re.search(rf'<Relationship Id="{re.escape(sheet_rid)}"[^>]*Target="([^"]+)"', rels_xml)
                if not m:
                    return False
                sheet_file = 'xl/' + m.group(1)

                # Read the sheet XML as a string
                sheet_xml = zin.read(sheet_file).decode('utf-8')

            # Step 2: Modify the sheet XML using string replacement
            modified = sheet_xml
            written = 0

            for cell_ref, value in cell_writes.items():
                # Pattern 1: Cell exists with a value — replace the value
                # <c r="F10" s="153"><v>old</v></c>  or  <c r="F10" t="s"><v>old</v></c>
                # We need to handle various attribute orders and self-closing cells

                # Pattern A: Self-closing empty cell: <c r="F10" s="153"/>
                pattern_a = rf'<c r="{cell_ref}"([^>]*)/>'
                m_a = re.search(pattern_a, modified)
                if m_a:
                    attrs = m_a.group(1)
                    # Remove t="s" or t="str" if present (we're writing a number)
                    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
                    replacement = f'<c r="{cell_ref}"{attrs}><v>{value}</v></c>'
                    modified = modified[:m_a.start()] + replacement + modified[m_a.end():]
                    written += 1
                    continue

                # Pattern B: Cell with existing value: <c r="F10" ...>...</c>
                # Could contain <f>formula</f><v>value</v> or just <v>value</v>
                pattern_b = rf'(<c r="{cell_ref}"[^>]*>)(.*?)(</c>)'
                m_b = re.search(pattern_b, modified, re.DOTALL)
                if m_b:
                    attrs = m_b.group(1)
                    inner = m_b.group(2)
                    # Remove t="s" or t="str" if present
                    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
                    # Remove any existing <v>...</v> and <f>...</f>
                    inner = re.sub(r'<v>.*?</v>', '', inner, flags=re.DOTALL)
                    inner = re.sub(r'<f>.*?</f>', '', inner, flags=re.DOTALL)
                    replacement = f'{attrs}<v>{value}</v></c>'
                    modified = modified[:m_b.start()] + replacement + modified[m_b.end():]
                    written += 1
                    continue

                # Pattern C: Cell doesn't exist in the XML at all
                # We need to insert it into the right row
                row_num = int(''.join(c for c in cell_ref if c.isdigit()))
                col_part = ''.join(c for c in cell_ref if c.isalpha())

                # Find the row element for this row number
                row_pattern = rf'(<row r="{row_num}"[^>]*>)'
                m_row = re.search(row_pattern, modified)
                if m_row:
                    # Insert the cell at the end of the row (before </row>)
                    row_start = m_row.start()
                    row_end_tag = modified.find('</row>', row_start)
                    if row_end_tag >= 0:
                        new_cell = f'<c r="{cell_ref}"><v>{value}</v></c>'
                        modified = modified[:row_end_tag] + new_cell + modified[row_end_tag:]
                        written += 1
                        continue

                # If we get here, the row doesn't exist either.
                # Find the sheetData closing tag and insert a new row before it.
                # Find the last row before </sheetData> to insert after it.
                sheet_data_end = modified.find('</sheetData>')
                if sheet_data_end >= 0:
                    new_row = f'<row r="{row_num}"><c r="{cell_ref}"><v>{value}</v></c></row>'
                    modified = modified[:sheet_data_end] + new_row + modified[sheet_data_end:]
                    written += 1

            if written == 0:
                return False

            # Step 3: Write the new zip — copy everything except the sheet file,
            # then write the modified sheet XML
            with zipfile.ZipFile(file_path, 'r') as zin:
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == sheet_file:
                            zout.writestr(item, modified)
                        else:
                            zout.writestr(item, zin.read(item.filename))

            # Step 4: Replace original with temp
            shutil.move(temp_path, file_path)
            return True

        except PermissionError:
            # Clean up temp file on error
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            raise  # Let the caller handle the "file is open" message
        except Exception as e:
            # Clean up temp file on error
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            # Re-raise so the caller's outer except can show a proper
            # error dialog with Report button — don't swallow it as
            # "manifest is open" which is misleading.
            raise

    def _has_progress(self):
        """Check if there's any upload progress that would be lost."""
        if self._uploading:
            return True
        for item in self._tree.get_children():
            status = self._tree.set(item, "status").lower()
            if status in ("uploaded", "uploading", "error", "missed"):
                return True
        if self._decl_numbers:
            return True
        return False

    def _on_close(self):
        """Close window with confirmation if there's active progress."""
        if self._has_progress() or self._uploading:
            answer = messagebox.askyesno(
                "Close Upload Window?",
                "You have upload progress in this session.\n\n"
                "Closing this window will:\n"
                "- Stop any active uploads\n"
                "- Lose captured declaration numbers\n"
                "- The Edge browser will stay open\n\n"
                "Are you sure you want to close?",
                icon="warning")
            if not answer:
                return
        self._stop_requested = True
        self._paused = False
        # Don't kill the browser - let user close it manually
        self.launcher._upload_win = None
        self.win.destroy()
        self.launcher.show()


# ==============================================================================
# MAIN
# ==============================================================================
if __name__ == "__main__":
    app = Launcher()
    app.run()
