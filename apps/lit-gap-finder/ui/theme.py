# -*- coding: utf-8 -*-
"""字型與配色（海洋科學色票，與 FC 系列其他 App 一致）。"""
import sys

NAVY = "#0B1F3A"
DEEP = "#065A82"
TEAL = "#1C7293"
GREEN = "#2E7D5B"
MOSS = "#5FA98A"
GOLD = "#C99A3E"
BG = "#F5F9FA"
INK = "#1E293B"
GREY = "#64748B"
CREDIT_FG = "#204a87"
RED = "#B42318"

if sys.platform == "darwin":
    FAMILY = "PingFang TC"
elif sys.platform.startswith("win"):
    FAMILY = "Microsoft JhengHei UI"
else:
    FAMILY = "Noto Sans CJK TC"

F = (FAMILY, 10)
FB = (FAMILY, 10, "bold")
FS = (FAMILY, 9)
FT = (FAMILY, 13, "bold")
FMONO = ("Consolas", 10) if sys.platform.startswith("win") else ("Menlo", 11)


def apply_ttk_style(root):
    """統一 ttk 元件字型與 Treeview 列高。"""
    from tkinter import font as tkfont, ttk
    for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
        try:
            tkfont.nametofont(nm).configure(family=FAMILY, size=10)
        except Exception:  # noqa: BLE001
            pass
    st = ttk.Style(root)
    try:
        if "vista" in st.theme_names():
            st.theme_use("vista")
    except Exception:
        pass
    st.configure(".", font=F)
    st.configure("Treeview", font=F, rowheight=24)
    st.configure("Treeview.Heading", font=FB)
    st.configure("TNotebook.Tab", font=FB, padding=(14, 4))
    st.configure("Accent.TButton", font=FB)
    return st
