# -*- coding: utf-8 -*-
"""右下角製作者字樣的統一掛載點。

規則（規格第 1 節）：每一個視窗／對話框右下角固定顯示 CREDIT。
- 主視窗：呼叫 attach_credit(root, status_var) 產生「左：狀態列｜右：CREDIT」底列。
- 所有子視窗：一律繼承 CreditToplevel（建構時自動掛 CREDIT 底列，內容放 self.body）。
- audit_*()：給 --selftest 用的程式化檢查（類別繼承、原始碼掃描、實體視窗位置）。
"""
import inspect
import os
import re
import tkinter as tk

from . import theme as T

CREDIT = "NDHU自資系游豐兆製作"
CREDIT_TAG = "fc_credit"


def attach_credit(win, status_var=None):
    """在 win 底部建立底列，右側放 CREDIT；回傳底列 Frame。必須在其他內容 pack 之前呼叫。"""
    foot = tk.Frame(win, bg=win.cget("bg"))
    foot.pack(side="bottom", fill="x", padx=10, pady=(0, 6))
    lbl = tk.Label(foot, text=CREDIT, font=T.FS, fg=T.CREDIT_FG, bg=win.cget("bg"), anchor="e")
    lbl.pack(side="right")
    lbl._fc_tag = CREDIT_TAG  # noqa: SLF001 - 稽核用標記
    if status_var is not None:
        tk.Label(foot, textvariable=status_var, font=T.FS, fg=T.GREY, bg=win.cget("bg"),
                 anchor="w").pack(side="left", fill="x", expand=True)
    win._fc_credit_label = lbl  # noqa: SLF001
    return foot


class CreditToplevel(tk.Toplevel):
    """所有子視窗的基底：自動掛 CREDIT、設定圖示、置中於父視窗。內容請放在 self.body。"""

    def __init__(self, master, title, size=None, modal=False, resizable=True):
        super().__init__(master)
        self.withdraw()
        self.title(title)
        try:
            self.configure(bg=T.BG)
        except tk.TclError:
            pass
        self.footer = attach_credit(self)
        self.body = tk.Frame(self, bg=T.BG)
        self.body.pack(side="top", fill="both", expand=True, padx=12, pady=(10, 4))
        _set_icon(self)
        if not resizable:
            self.resizable(False, False)
        self._size = size
        self._modal = modal
        try:
            self.transient(master.winfo_toplevel())
        except tk.TclError:
            pass

    def show(self):
        """排版完成後呼叫：置中、顯示、（modal 時）搶焦點。"""
        self.update_idletasks()
        w, h = self._size or (self.winfo_reqwidth(), self.winfo_reqheight())
        w = max(w, self.winfo_reqwidth())
        h = max(h, self.winfo_reqheight())
        try:
            m = self.master.winfo_toplevel()
            if m.winfo_viewable():
                x = m.winfo_rootx() + max(0, (m.winfo_width() - w) // 2)
                y = m.winfo_rooty() + max(0, (m.winfo_height() - h) // 3)
            else:
                x = (self.winfo_screenwidth() - w) // 2
                y = (self.winfo_screenheight() - h) // 3
        except tk.TclError:
            x, y = 100, 100
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
        self.lift()
        if self._modal:
            try:
                self.grab_set()
            except tk.TclError:
                pass
            self.focus_set()
        return self


_ICON_CACHE = {}


def _set_icon(win):
    try:
        from .resources import resource_path
        ico = resource_path("icon.ico")
        png = resource_path("icon.png")
        if os.name == "nt" and os.path.exists(ico):
            win.iconbitmap(ico)
        elif os.path.exists(png):
            img = _ICON_CACHE.get("png")
            if img is None:
                img = tk.PhotoImage(file=png)
                _ICON_CACHE["png"] = img
            win.iconphoto(False, img)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 稽核（--selftest 使用）
# ---------------------------------------------------------------------------
def find_credit_label(win):
    """回傳 win 內文字為 CREDIT 的 Label（遞迴），找不到回 None。"""
    stack = [win]
    while stack:
        w = stack.pop()
        try:
            if isinstance(w, tk.Label) and w.cget("text") == CREDIT:
                return w
        except tk.TclError:
            pass
        stack.extend(w.winfo_children())
    return None


def credit_in_bottom_right(win, tol=40):
    """實體視窗檢查：CREDIT 存在、可見，且右緣與下緣貼近視窗右下角。回 (ok, 說明)。"""
    lbl = find_credit_label(win)
    if lbl is None:
        return False, "找不到 CREDIT 標籤"
    win.update_idletasks()
    win.update()
    wx, wy = win.winfo_rootx(), win.winfo_rooty()
    ww, wh = win.winfo_width(), win.winfo_height()
    lx, ly = lbl.winfo_rootx(), lbl.winfo_rooty()
    lw, lh = lbl.winfo_width(), lbl.winfo_height()
    right_gap = (wx + ww) - (lx + lw)
    bottom_gap = (wy + wh) - (ly + lh)
    ok = lbl.winfo_ismapped() and 0 <= right_gap <= tol and 0 <= bottom_gap <= tol
    return ok, f"右距 {right_gap}px、下距 {bottom_gap}px、可見={bool(lbl.winfo_ismapped())}"


def audit_classes(modules):
    """所有 tk.Toplevel 子類別都必須繼承 CreditToplevel。回傳 (通過類別名單, 違規類別名單)。"""
    good, bad = [], []
    for mod in modules:
        for name, obj in inspect.getmembers(mod, inspect.isclass):
            if obj.__module__ != mod.__name__:
                continue
            if issubclass(obj, tk.Toplevel) and obj is not CreditToplevel:
                (good if issubclass(obj, CreditToplevel) else bad).append(f"{mod.__name__}.{name}")
    return good, bad


_RAW_TOPLEVEL = re.compile(r"(?<![A-Za-z_])(?:tk\.|tkinter\.)?Toplevel\(")
_MSGBOX = re.compile(r"\bmessagebox\b|tkinter\.simpledialog|\bsimpledialog\b")


def audit_sources(paths):
    """原始碼掃描：除 credit.py 外不得直接 Toplevel(…)，全專案不得用 tkinter.messagebox。回傳違規清單。"""
    problems = []
    for p in paths:
        try:
            src = open(p, encoding="utf-8").read()
        except OSError:
            continue
        base = os.path.basename(p)
        for i, line in enumerate(src.splitlines(), 1):
            s = line.split("#", 1)[0]
            if base != "credit.py" and _RAW_TOPLEVEL.search(s) and "class " not in s:
                problems.append(f"{base}:{i} 直接建立 Toplevel：{line.strip()}")
            if _MSGBOX.search(s) and "import" in s:
                problems.append(f"{base}:{i} 使用 messagebox／simpledialog：{line.strip()}")
    return problems
