# -*- coding: utf-8 -*-
"""FcDialog：取代 tkinter 內建訊息框的自製對話框（右下角有 CREDIT）。

用法：
    info(parent, "標題", "內容")
    warn(parent, "標題", "內容")
    error(parent, "標題", "內容", detail="可展開的技術細節")
    ask(parent, "標題", "要繼續嗎？") -> bool

環境變數 LITGAP_NO_DIALOG=1：不顯示、直接回傳預設值（自動化／CI 使用），訊息寫入 HISTORY。
"""
import os
import tkinter as tk
from tkinter import ttk

from . import theme as T
from .credit import CreditToplevel

HISTORY = []          # (kind, title, message) — 自動化測試可檢查
_ICONS = {"info": ("i", T.TEAL), "warn": ("!", T.GOLD), "error": ("X", T.RED), "ask": ("?", T.DEEP)}


class FcDialog(CreditToplevel):
    def __init__(self, parent, title, message, kind="info", buttons=None, detail="", modal=True):
        super().__init__(parent, title, modal=modal, resizable=False)
        self.result = None
        buttons = buttons or (["是", "否"] if kind == "ask" else ["確定"])
        glyph, color = _ICONS.get(kind, _ICONS["info"])
        top = tk.Frame(self.body, bg=T.BG)
        top.pack(fill="both", expand=True)
        c = tk.Canvas(top, width=44, height=44, bg=T.BG, highlightthickness=0)
        c.create_oval(2, 2, 42, 42, fill=color, outline=color)
        c.create_text(22, 22, text=glyph, fill="white", font=(T.FAMILY, 16, "bold"))
        c.pack(side="left", anchor="n", padx=(0, 12), pady=4)
        tk.Label(top, text=message, font=T.F, fg=T.INK, bg=T.BG, justify="left",
                 wraplength=440, anchor="w").pack(side="left", fill="both", expand=True)
        if detail:
            box = tk.Frame(self.body, bg=T.BG)
            box.pack(fill="both", expand=True, pady=(8, 0))
            txt = tk.Text(box, height=7, width=64, font=T.FMONO, wrap="word", bg="white")
            txt.insert("1.0", detail)
            txt.config(state="disabled")
            txt.pack(fill="both", expand=True)
        bar = tk.Frame(self.body, bg=T.BG)
        bar.pack(fill="x", pady=(12, 2))
        self.buttons = []
        for label in reversed(buttons):
            b = ttk.Button(bar, text=label, width=10, command=lambda v=label: self._close(v))
            b.pack(side="right", padx=(6, 0))
            self.buttons.append(b)
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(None))
        self.bind("<Return>", lambda e: self._close(buttons[0]))
        self.bind("<Escape>", lambda e: self._close(None))
        self.show()
        if self.buttons:
            self.buttons[-1].focus_set()

    def _close(self, value):
        self.result = value
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()


def _headless():
    return os.environ.get("LITGAP_NO_DIALOG") == "1"


def _run(parent, title, message, kind, buttons=None, detail="", default=None):
    HISTORY.append((kind, title, message))
    if _headless():
        return default
    d = FcDialog(parent, title, message, kind=kind, buttons=buttons, detail=detail)
    parent.wait_window(d)
    return d.result


def info(parent, title, message, detail=""):
    _run(parent, title, message, "info", detail=detail)


def warn(parent, title, message, detail=""):
    _run(parent, title, message, "warn", detail=detail)


def error(parent, title, message, detail=""):
    _run(parent, title, message, "error", detail=detail)


def ask(parent, title, message, yes="是", no="否", default=False):
    r = _run(parent, title, message, "ask", buttons=[yes, no], default=yes if default else no)
    return r == yes
