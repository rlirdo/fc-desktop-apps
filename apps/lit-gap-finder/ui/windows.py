# -*- coding: utf-8 -*-
"""子視窗：設定、關於、手動加入中文文獻（皆繼承 CreditToplevel → 右下角 CREDIT）。"""
import re
import tkinter as tk
from tkinter import ttk

from . import backend as BK
from . import dialogs
from . import theme as T
from .credit import CREDIT, CreditToplevel

APP_TITLE = "FC文獻缺口搜尋工具 FC LitGap Finder"
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
KEY_HINT = ("不填也能用，但每日檢索額度較低（免金鑰同一網路每天約 100 次 OpenAlex 檢索）。"
            "金鑰只存在本機設定檔，不會寫進任何匯出檔或查詢紀錄。")
WEIGHT_LABELS = [("title", "標題命中（每次 ×）"), ("abstract", "摘要命中（每次 ×）"),
                 ("cited", "被引次數 log(1+n) ×"), ("recency", "年份近度 ×")]


def _row(parent, r, label, widget, hint=""):
    tk.Label(parent, text=label, font=T.F, bg=T.BG, fg=T.INK, anchor="e").grid(
        row=r, column=0, sticky="e", padx=(0, 8), pady=3)
    widget.grid(row=r, column=1, sticky="we", pady=3)
    if hint:
        tk.Label(parent, text=hint, font=T.FS, bg=T.BG, fg=T.GREY, anchor="w").grid(
            row=r, column=2, sticky="w", padx=(8, 0))


# ---------------------------------------------------------------------------
class SettingsWindow(CreditToplevel):
    """設定：Unpaywall email、章節清單、相關性權重、快取清除。存到 %APPDATA%\\FCLitGap\\config.json。"""

    def __init__(self, parent, cfg, on_save=None, modal=True):
        super().__init__(parent, "設定", size=(700, 640), modal=modal)
        self.cfg = dict(cfg)
        self.on_save = on_save
        tk.Label(self.body, text="設定", font=T.FT, fg=T.NAVY, bg=T.BG, anchor="w").pack(fill="x")
        tk.Label(self.body, text="設定檔位置：%APPDATA%\\FCLitGap\\config.json（不儲存任何密碼）",
                 font=T.FS, fg=T.GREY, bg=T.BG, anchor="w").pack(fill="x", pady=(0, 6))

        g = tk.Frame(self.body, bg=T.BG)
        g.pack(fill="x")
        g.columnconfigure(1, weight=1)
        self.var_email = tk.StringVar(value=self.cfg.get("email", ""))
        _row(g, 0, "Unpaywall email", ttk.Entry(g, textvariable=self.var_email, width=40),
             "也作為 Crossref／OpenAlex mailto")
        self.var_dir = tk.StringVar(value=self.cfg.get("last_dir", ""))
        _row(g, 1, "上次存檔資料夾", ttk.Entry(g, textvariable=self.var_dir, width=40))
        self.var_key = tk.StringVar(value=self.cfg.get("openalex_api_key", "") or "")
        kf = tk.Frame(g, bg=T.BG)
        self.ent_key = ttk.Entry(kf, textvariable=self.var_key, width=32, show="•")
        self.ent_key.pack(side="left", fill="x", expand=True)
        self.var_show = tk.BooleanVar(value=False)
        tk.Checkbutton(kf, text="顯示", variable=self.var_show, font=T.FS, bg=T.BG,
                       command=lambda: self.ent_key.config(show="" if self.var_show.get() else "•")).pack(side="left")
        _row(g, 2, "OpenAlex API 金鑰\n（選填，只存本機）", kf)
        tk.Label(g, text=KEY_HINT, font=T.FS, fg=T.GREY, bg=T.BG, justify="left", anchor="w",
                 wraplength=520).grid(row=3, column=1, columnspan=2, sticky="w")

        lf = tk.LabelFrame(self.body, text=" 章節標籤清單（一行一個） ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lf.pack(fill="both", expand=True, pady=(8, 4))
        self.txt_ch = tk.Text(lf, height=7, font=T.F, wrap="none")
        self.txt_ch.insert("1.0", "\n".join(self.cfg.get("chapters") or BK.DEFAULT_CHAPTERS))
        self.txt_ch.pack(fill="both", expand=True, padx=6, pady=6)

        lw = tk.LabelFrame(self.body, text=" 相關性權重（預設值；指標頁可臨時調整） ", font=T.FB,
                           fg=T.TEAL, bg=T.BG)
        lw.pack(fill="x", pady=4)
        w = self.cfg.get("weights") or {}
        self.var_w = {}
        for i, (k, lab) in enumerate(WEIGHT_LABELS):
            v = tk.DoubleVar(value=float(w.get(k, 1.0)))
            self.var_w[k] = v
            tk.Label(lw, text=lab, font=T.F, bg=T.BG).grid(row=i // 2, column=(i % 2) * 2, sticky="e", padx=6)
            ttk.Spinbox(lw, from_=0, to=10, increment=0.5, textvariable=v, width=6).grid(
                row=i // 2, column=(i % 2) * 2 + 1, sticky="w", pady=2)

        bar = tk.Frame(self.body, bg=T.BG)
        bar.pack(fill="x", pady=(8, 2))
        ttk.Button(bar, text="清除快取", command=self.clear_cache).pack(side="left")
        ttk.Button(bar, text="還原預設", command=self.reset).pack(side="left", padx=6)
        ttk.Button(bar, text="取消", width=8, command=self.destroy).pack(side="right")
        ttk.Button(bar, text="儲存", width=8, command=self.save).pack(side="right", padx=6)
        self.show()

    def reset(self):
        d = BK.DEFAULTS or {}
        self.var_email.set(d.get("email", ""))
        self.txt_ch.delete("1.0", "end")
        self.txt_ch.insert("1.0", "\n".join(d.get("chapters") or BK.DEFAULT_CHAPTERS))
        for k, v in (d.get("weights") or {}).items():
            if k in self.var_w:
                self.var_w[k].set(v)

    def clear_cache(self):
        try:
            n = BK.call(BK.cache_clear)
            dialogs.info(self, "快取", f"已清除快取{'（' + str(n) + ' 筆）' if isinstance(n, int) else ''}。")
        except Exception as e:  # noqa: BLE001
            dialogs.error(self, "快取", f"清除失敗：{e}")

    def collect(self):
        email = self.var_email.get().strip()
        if email and not EMAIL_RE.match(email):
            raise ValueError("email 格式不正確")
        chapters = [s.strip() for s in self.txt_ch.get("1.0", "end").splitlines() if s.strip()]
        if "未分類" not in chapters:
            chapters.append("未分類")
        cfg = dict(self.cfg)
        key = self.var_key.get().strip()
        if key and (len(key) < 8 or any(c.isspace() for c in key)):
            raise ValueError("OpenAlex API 金鑰格式不正確（不可含空白）")
        cfg.update({"email": email, "mailto": email, "last_dir": self.var_dir.get().strip(),
                    "openalex_api_key": key,
                    "chapters": chapters,
                    "weights": {k: float(v.get()) for k, v in self.var_w.items()}})
        return cfg

    def save(self):
        try:
            cfg = self.collect()
        except (ValueError, tk.TclError) as e:
            dialogs.warn(self, "設定", f"無法儲存：{e}")
            return
        try:
            BK.call(BK.config_save, cfg)
        except Exception as e:  # noqa: BLE001
            dialogs.error(self, "設定", f"寫入設定檔失敗：{e}")
            return
        if self.on_save:
            self.on_save(cfg)
        self.destroy()


# ---------------------------------------------------------------------------
class AboutWindow(CreditToplevel):
    def __init__(self, parent, version, backend_status="", modal=True):
        super().__init__(parent, "關於", size=(600, 470), modal=modal, resizable=False)
        head = tk.Frame(self.body, bg=T.NAVY)
        head.pack(fill="x")
        tk.Label(head, text="FC文獻缺口搜尋工具", font=(T.FAMILY, 18, "bold"), fg="white",
                 bg=T.NAVY, anchor="w").pack(fill="x", padx=14, pady=(12, 0))
        tk.Label(head, text=f"FC LitGap Finder　v{version}", font=(T.FAMILY, 11, "italic"),
                 fg="#cfe3ea", bg=T.NAVY, anchor="w").pack(fill="x", padx=14, pady=(0, 12))
        text = (
            "用途：以多組關鍵字檢索開放書目資料庫，依可勾選的合格指標篩選文獻、"
            "找出關鍵字組合間的研究缺口，並匯出 APA 7 參考文獻（Word／PDF）、Excel 清單、RIS／BibTeX 與趨勢圖。\n\n"
            "資料來源：Crossref、OpenAlex、Unpaywall、Semantic Scholar、DOAJ、doi.org。\n"
            "只下載合法開放取用（OA）全文；非 OA 文獻只提供 DOI 連結。\n"
            "中文文獻請用「手動加入」，匯出時會註明「使用者自填，未經 API 驗證」。\n\n"
            "限制：書目資料以資料庫回傳為準，送出前請人工核對；"
            "「待人工確認」代表該期刊無法自動判定是否收錄於 DOAJ 或主要出版社。"
        )
        tk.Label(self.body, text=text, font=T.F, fg=T.INK, bg=T.BG, justify="left", wraplength=560,
                 anchor="w").pack(fill="x", pady=(10, 6))
        if backend_status:
            tk.Label(self.body, text=backend_status, font=T.FS, fg=T.GREY, bg=T.BG, anchor="w").pack(fill="x")
        tk.Label(self.body, text=f"製作：{CREDIT}", font=T.FB, fg=T.GREEN, bg=T.BG,
                 anchor="w").pack(fill="x", pady=(6, 0))
        ttk.Button(self.body, text="關閉", width=10, command=self.destroy).pack(side="bottom", anchor="e")
        self.show()


# ---------------------------------------------------------------------------
class ManualAddWindow(CreditToplevel):
    """手動加入中文文獻 → Paper(source="manual")。"""

    FIELDS = [("authors", "作者 *", "多位以「、」或「;」分隔，例：王小明、陳大華"),
              ("year", "年 *", "西元年，例：2021"),
              ("title", "題名 *", ""),
              ("container", "刊名 *", "例：臺灣化學教育"),
              ("volume", "卷", ""), ("issue", "期", ""), ("pages", "頁", "例：13-19"),
              ("doi", "DOI（選填）", "例：10.6555/xxxx 或 https://doi.org/…")]

    def __init__(self, parent, chapters, on_add=None, modal=True):
        super().__init__(parent, "手動加入中文文獻", size=(720, 480), modal=modal)
        self.on_add = on_add
        self.result = None
        tk.Label(self.body, text="手動加入中文文獻", font=T.FT, fg=T.NAVY, bg=T.BG, anchor="w").pack(fill="x")
        tk.Label(self.body, text="中文文獻不自動檢索；此處填寫的資料會格式化為 APA 7 並一起匯出，"
                 "並標註「使用者自填，未經 API 驗證」。", font=T.FS, fg=T.GREY, bg=T.BG,
                 wraplength=600, justify="left", anchor="w").pack(fill="x", pady=(0, 6))
        g = tk.Frame(self.body, bg=T.BG)
        g.pack(fill="x")
        g.columnconfigure(1, weight=1)
        self.vars = {}
        for i, (k, lab, hint) in enumerate(self.FIELDS):
            v = tk.StringVar()
            self.vars[k] = v
            _row(g, i, lab, ttk.Entry(g, textvariable=v, width=44), hint)
        self.var_ch = tk.StringVar(value="未分類")
        _row(g, len(self.FIELDS), "章節標籤",
             ttk.Combobox(g, textvariable=self.var_ch, values=list(chapters), state="readonly", width=20))
        self.var_preview = tk.StringVar(value="")
        tk.Label(self.body, textvariable=self.var_preview, font=T.F, fg=T.DEEP, bg=T.BG, wraplength=600,
                 justify="left", anchor="w").pack(fill="x", pady=(8, 0))
        bar = tk.Frame(self.body, bg=T.BG)
        bar.pack(fill="x", side="bottom", pady=(6, 2))
        ttk.Button(bar, text="預覽 APA 7", command=self.preview).pack(side="left")
        ttk.Button(bar, text="取消", width=8, command=self.destroy).pack(side="right")
        ttk.Button(bar, text="加入", width=8, command=self.add).pack(side="right", padx=6)
        self.show()

    def build_paper(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        miss = [lab.rstrip(" *") for k, lab, _ in self.FIELDS if lab.endswith("*") and not v[k]]
        if miss:
            raise ValueError("請填寫：" + "、".join(miss))
        if not re.fullmatch(r"\d{4}", v["year"]):
            raise ValueError("年請填四位數西元年")
        names = [n.strip() for n in re.split(r"[、;；,，]", v["authors"]) if n.strip()]
        p = BK.Paper(
            doi=BK.models.normalize_doi(v["doi"]), title=v["title"],
            authors=[{"family": n, "given": ""} for n in names], year=int(v["year"]),
            container=v["container"], volume=v["volume"], issue=v["issue"],
            pages=v["pages"].replace("–", "-").replace("~", "-"), type="journal-article",
            language="zh", source="manual", chapter=self.var_ch.get() or "未分類", include=True,
            note="使用者自填，未經 API 驗證")
        return p

    def preview(self):
        try:
            p = self.build_paper()
            self.var_preview.set("APA 7：" + BK.format_apa7(p) + "\n文內引用：" + BK.in_text(p))
        except Exception as e:  # noqa: BLE001
            self.var_preview.set(f"⚠ {e}")

    def add(self):
        try:
            p = self.build_paper()
        except ValueError as e:
            dialogs.warn(self, "手動加入", str(e))
            return
        self.result = p
        if self.on_add:
            self.on_add(p)
        self.destroy()
