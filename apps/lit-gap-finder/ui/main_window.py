# -*- coding: utf-8 -*-
"""主視窗：三分頁（檢索／指標與缺口／結果）＋底部狀態列＋右下 CREDIT。"""
import datetime as _dt
import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, ttk

from . import backend as BK
from . import dialogs
from . import theme as T
from .credit import attach_credit, _set_icon
from .progress import run_task
from .results import ResultsTable
from .windows import APP_TITLE, AboutWindow, ManualAddWindow, SettingsWindow, WEIGHT_LABELS

MAX_GROUPS = 6
LANGS = [("不限", "any"), ("英文", "en"), ("中文", "zh")]
TYPE_CHOICES = [("journal-article", "期刊論文"), ("proceedings-article", "會議論文"),
                ("book-chapter", "專書章節")]
CHECK_KEYS = ["doi", "oa", "type", "year", "issn", "lang"]
ORDER_CHOICES = ["APA 規則（先中文筆畫、後英文字母）", "自訂（結果表拖曳後的順序）"]


def parse_groups(text):
    """每行一組、「;」或「；」分同義詞 → list[list[str]]（空行略過，最多 6 組）。"""
    groups = []
    for line in (text or "").splitlines():
        terms = [t.strip() for t in re.split(r"[;；]", line) if t.strip()]
        if terms:
            groups.append(terms)
    return groups[:MAX_GROUPS]


def safe_name(s, n=24):
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", s or "").strip("_")
    return s[:n] or "查詢"


def open_path(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class MainWindow:
    def __init__(self, root, version):
        self.root = root
        self.version = version
        self.cfg = BK.call(BK.config_load)
        self.papers = []          # 全部（含手動加入）
        self.gap = None
        self.task = None
        self.last_export_dir = ""
        self.query_info = {}

        root.title(APP_TITLE)
        root.geometry("1100x720")
        root.minsize(900, 600)
        root.configure(bg=T.BG)
        _set_icon(root)
        T.apply_ttk_style(root)

        self._build_menu()
        self.var_status = tk.StringVar(value=f"就緒。{BK.STATUS}")
        attach_credit(root, self.var_status)       # 底部：狀態列（左）＋ CREDIT（右）

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(8, 4))
        self.tab_search = tk.Frame(self.nb, bg=T.BG)
        self.tab_ind = tk.Frame(self.nb, bg=T.BG)
        self.tab_res = tk.Frame(self.nb, bg=T.BG)
        self.nb.add(self.tab_search, text="  ① 檢索  ")
        self.nb.add(self.tab_ind, text="  ② 指標與缺口  ")
        self.nb.add(self.tab_res, text="  ③ 結果  ")
        self._build_search_tab()
        self._build_ind_tab()
        self._build_res_tab()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # =====================================================================
    def _build_menu(self):
        mb = tk.Menu(self.root, font=T.F)
        m1 = tk.Menu(mb, tearoff=0, font=T.F)
        m1.add_command(label="設定…", command=self.open_settings)
        m1.add_command(label="開啟輸出資料夾", command=self.open_out_dir)
        m1.add_separator()
        m1.add_command(label="結束", command=self.on_close)
        mb.add_cascade(label="檔案", menu=m1)
        m2 = tk.Menu(mb, tearoff=0, font=T.F)
        m2.add_command(label="關於…", command=self.open_about)
        mb.add_cascade(label="說明", menu=m2)
        self.root.config(menu=mb)

    def _header(self, parent, title, sub):
        h = tk.Frame(parent, bg=T.BG)
        h.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(h, text=title, font=T.FT, fg=T.NAVY, bg=T.BG, anchor="w").pack(side="left")
        tk.Label(h, text="　" + sub, font=T.FS, fg=T.GREY, bg=T.BG, anchor="w").pack(side="left")
        return h

    # ---------------- ① 檢索 ----------------
    def _build_search_tab(self):
        tab = self.tab_search
        h = self._header(tab, "檢索", "Crossref＋OpenAlex 合併去重，再以 Unpaywall／Semantic Scholar 補 OA 與被引")
        ttk.Button(h, text="關於", command=self.open_about).pack(side="right")
        ttk.Button(h, text="設定", command=self.open_settings).pack(side="right", padx=6)

        lf = tk.LabelFrame(tab, text=" 關鍵字組（最多 6 組；每組一行，同義詞用「;」分隔，可中英混列） ",
                           font=T.FB, fg=T.TEAL, bg=T.BG)
        lf.pack(fill="x", padx=12, pady=4)
        self.txt_kw = tk.Text(lf, height=6, font=(T.FAMILY, 11), wrap="none", undo=True)
        self.txt_kw.pack(fill="x", padx=6, pady=(6, 2))
        tk.Label(lf, text="例：serpentine; 蛇紋石　／　deep sea water; 深層海水　／　far infrared; 遠紅外線",
                 font=T.FS, fg=T.GREY, bg=T.BG, anchor="w").pack(fill="x", padx=6, pady=(0, 6))

        g = tk.Frame(tab, bg=T.BG)
        g.pack(fill="x", padx=12, pady=4)
        this_year = _dt.date.today().year
        self.var_yf = tk.IntVar(value=int(self.cfg.get("year_from") or 2016))
        self.var_yt = tk.IntVar(value=int(self.cfg.get("year_to") or this_year))
        self.var_lang = tk.StringVar(value=dict((v, k) for k, v in LANGS).get(self.cfg.get("lang", "any"), "不限"))
        self.var_limit = tk.IntVar(value=int(self.cfg.get("limit") or 50))
        self.var_thr = tk.IntVar(value=int(self.cfg.get("gap_threshold") or 5))
        self.var_email = tk.StringVar(value=self.cfg.get("email", ""))
        self.var_dir = tk.StringVar(value=self.cfg.get("last_dir") or os.path.join(os.path.expanduser("~"), "Documents"))

        def lab(r, c, text):
            tk.Label(g, text=text, font=T.F, bg=T.BG, fg=T.INK).grid(row=r, column=c, sticky="e", padx=(10, 4), pady=4)
        lab(0, 0, "出版年 起")
        ttk.Spinbox(g, from_=1900, to=this_year, textvariable=self.var_yf, width=7).grid(row=0, column=1, sticky="w")
        lab(0, 2, "迄")
        ttk.Spinbox(g, from_=1900, to=this_year + 1, textvariable=self.var_yt, width=7).grid(row=0, column=3, sticky="w")
        lab(0, 4, "語言")
        ttk.Combobox(g, textvariable=self.var_lang, values=[k for k, _ in LANGS], state="readonly",
                     width=6).grid(row=0, column=5, sticky="w")
        lab(0, 6, "每組上限")
        ttk.Spinbox(g, from_=5, to=500, increment=5, textvariable=self.var_limit, width=6).grid(row=0, column=7, sticky="w")
        lab(0, 8, "缺口門檻（篇）")
        ttk.Spinbox(g, from_=1, to=100, textvariable=self.var_thr, width=5).grid(row=0, column=9, sticky="w")
        lab(1, 0, "Unpaywall email")
        ttk.Entry(g, textvariable=self.var_email, width=34).grid(row=1, column=1, columnspan=5, sticky="we")
        tk.Label(g, text="（僅作 API 禮貌識別，不會寄信）", font=T.FS, fg=T.GREY, bg=T.BG).grid(
            row=1, column=6, columnspan=4, sticky="w", padx=6)
        lab(2, 0, "存檔資料夾")
        ttk.Entry(g, textvariable=self.var_dir, width=60).grid(row=2, column=1, columnspan=7, sticky="we")
        ttk.Button(g, text="瀏覽…", command=self.browse_dir).grid(row=2, column=8, sticky="w", padx=6)

        bar = tk.Frame(tab, bg=T.BG)
        bar.pack(fill="x", padx=12, pady=(8, 4))
        self.btn_search = tk.Button(bar, text="開始檢索", font=T.FB, width=12, bg=T.DEEP, fg="white",
                                    activebackground=T.NAVY, activeforeground="white", command=self.start_search)
        self.btn_search.pack(side="left")
        self.btn_gap = tk.Button(bar, text="缺口矩陣", font=T.FB, width=12, bg=T.GREEN, fg="white",
                                 activebackground=T.NAVY, activeforeground="white", command=self.start_gap)
        self.btn_gap.pack(side="left", padx=8)
        self.btn_stop = ttk.Button(bar, text="停止", width=8, command=self.stop, state="disabled")
        self.btn_stop.pack(side="left")
        self.var_mode = tk.StringVar(value="union")
        tk.Label(bar, text="檢索方式：", font=T.F, bg=T.BG).pack(side="left", padx=(18, 2))
        tk.Radiobutton(bar, text="各組分別檢索（聯集）", value="union", variable=self.var_mode, font=T.F,
                       bg=T.BG).pack(side="left")
        tk.Radiobutton(bar, text="所有組同時符合（交集）", value="and", variable=self.var_mode, font=T.F,
                       bg=T.BG).pack(side="left")

        lf2 = tk.LabelFrame(tab, text=" 執行紀錄 ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lf2.pack(fill="both", expand=True, padx=12, pady=(4, 6))
        self.txt_log = tk.Text(lf2, height=8, font=T.FS, bg="white", fg=T.INK, state="disabled", wrap="word")
        sb = ttk.Scrollbar(lf2, command=self.txt_log.yview)
        self.txt_log.config(yscrollcommand=sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb.pack(side="right", fill="y", pady=6)

    # ---------------- ② 指標與缺口 ----------------
    def _build_ind_tab(self):
        tab = self.tab_ind
        self._header(tab, "指標與缺口", "勾選合格指標後按「套用指標」；缺口矩陣在「檢索」頁按「缺口矩陣」產生")
        body = tk.Frame(tab, bg=T.BG)
        body.pack(fill="both", expand=True, padx=12, pady=4)
        left = tk.Frame(body, bg=T.BG)
        left.pack(side="left", fill="y")
        right = tk.Frame(body, bg=T.BG)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))

        lf = tk.LabelFrame(left, text=" 合格指標 ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lf.pack(fill="x")
        checks = self.cfg.get("checks") or {}
        self.var_checks = {}
        for k in CHECK_KEYS:
            v = tk.BooleanVar(value=bool(checks.get(k, True)))
            self.var_checks[k] = v
            tk.Checkbutton(lf, text=BK.INDICATOR_NAMES.get(k, k), variable=v, font=T.F, bg=T.BG,
                           anchor="w").pack(fill="x", padx=6)
        self.var_classic = tk.BooleanVar(value=bool(self.cfg.get("classic")))
        tk.Checkbutton(lf, text="　⑥ 含經典（不限年）", variable=self.var_classic, font=T.F, bg=T.BG,
                       anchor="w").pack(fill="x", padx=6, pady=(0, 4))

        lt = tk.LabelFrame(left, text=" ③ 文獻類型（可複選） ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lt.pack(fill="x", pady=6)
        types = set(self.cfg.get("types") or ["journal-article"])
        self.var_types = {}
        for k, lab in TYPE_CHOICES:
            v = tk.BooleanVar(value=k in types)
            self.var_types[k] = v
            tk.Checkbutton(lt, text=lab, variable=v, font=T.F, bg=T.BG).pack(side="left", padx=6)

        lw = tk.LabelFrame(left, text=" ⑫ 相關性權重 ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lw.pack(fill="x")
        w = self.cfg.get("weights") or {}
        self.var_w = {}
        for k, lab in WEIGHT_LABELS:
            v = tk.DoubleVar(value=float(w.get(k, 1.0)))
            self.var_w[k] = v
            row = tk.Frame(lw, bg=T.BG)
            row.pack(fill="x", padx=6)
            tk.Label(row, text=lab, font=T.FS, bg=T.BG, width=18, anchor="w").pack(side="left")
            tk.Scale(row, from_=0, to=5, resolution=0.5, orient="horizontal", variable=v, length=150,
                     bg=T.BG, highlightthickness=0, font=T.FS).pack(side="left")

        lm = tk.LabelFrame(left, text=" 篩選方式 ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lm.pack(fill="x", pady=6)
        self.var_keep = tk.StringVar(value=self.cfg.get("keep_mode", "mark"))
        tk.Radiobutton(lm, text="全部保留但標示（✘ 紅字）", value="mark", variable=self.var_keep, font=T.F,
                       bg=T.BG).pack(anchor="w", padx=6)
        tk.Radiobutton(lm, text="只保留全部通過", value="strict", variable=self.var_keep, font=T.F,
                       bg=T.BG).pack(anchor="w", padx=6)
        tk.Button(left, text="套用指標", font=T.FB, bg=T.DEEP, fg="white", activebackground=T.NAVY,
                  activeforeground="white", command=self.apply_indicators).pack(fill="x", pady=(4, 0))

        lg = tk.LabelFrame(right, text=" 缺口矩陣（兩兩、三三組合在 OpenAlex 的命中篇數） ", font=T.FB,
                           fg=T.TEAL, bg=T.BG)
        lg.pack(fill="both", expand=True)
        cols = [("kind", "類型", 60), ("combo", "組合", 380), ("count", "篇數", 70), ("flag", "判定", 80)]
        self.gap_tree = ttk.Treeview(lg, columns=[c[0] for c in cols], show="headings", height=8)
        for cid, lab, wd in cols:
            self.gap_tree.heading(cid, text=lab)
            self.gap_tree.column(cid, width=wd, anchor="w" if cid == "combo" else "center",
                                 stretch=cid == "combo")
        self.gap_tree.tag_configure("gap", foreground=T.RED)
        self.gap_tree.pack(fill="both", expand=True, padx=6, pady=(6, 2))
        self.gap_tree.bind("<<TreeviewSelect>>", self._gap_selected)
        tk.Label(lg, text="選取一列可看前 10 篇：", font=T.FS, fg=T.GREY, bg=T.BG, anchor="w").pack(fill="x", padx=6)
        self.lst_top = tk.Listbox(lg, height=6, font=T.FS)
        self.lst_top.pack(fill="x", padx=6, pady=(0, 6))
        lc = tk.LabelFrame(right, text=" 近 10 年每年篇數（匯出時另存 趨勢圖.png） ", font=T.FB, fg=T.TEAL, bg=T.BG)
        lc.pack(fill="x", pady=(6, 0))
        self.cv_trend = tk.Canvas(lc, height=130, bg="white", highlightthickness=0)
        self.cv_trend.pack(fill="x", padx=6, pady=6)
        self.cv_trend.bind("<Configure>", lambda e: self._draw_trend())

    # ---------------- ③ 結果 ----------------
    def _build_res_tab(self):
        tab = self.tab_res
        h = self._header(tab, "結果", "點「納入」欄勾選、點「章節」欄選標籤、拖曳列調整匯出順序、雙擊開 DOI、右鍵更多")
        self.var_count = tk.StringVar(value="0 筆")
        tk.Label(h, textvariable=self.var_count, font=T.FB, fg=T.GREEN, bg=T.BG).pack(side="right")
        pw = ttk.PanedWindow(tab, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=12, pady=4)
        self.table = ResultsTable(pw, get_chapters=self.chapters, on_select=self.show_detail,
                                  on_change=self._update_count)
        pw.add(self.table, weight=4)
        side = tk.Frame(pw, bg=T.BG)
        pw.add(side, weight=2)
        pw.bind("<Map>", lambda e: pw.after(50, lambda: self._init_sash(pw)), add="+")
        tk.Label(side, text="APA 7 預覽與指標詳情", font=T.FB, fg=T.TEAL, bg=T.BG, anchor="w").pack(fill="x", padx=6)
        self.txt_detail = tk.Text(side, font=T.F, wrap="word", bg="white", width=36, state="disabled")
        self.txt_detail.pack(fill="both", expand=True, padx=6, pady=4)
        self.txt_detail.tag_configure("h", font=T.FB, foreground=T.NAVY)
        self.txt_detail.tag_configure("ok", foreground=T.GREEN)
        self.txt_detail.tag_configure("bad", foreground=T.RED)
        self.txt_detail.tag_configure("q", foreground=T.GOLD)

        bar = tk.Frame(tab, bg=T.BG)
        bar.pack(fill="x", padx=12, pady=(2, 6))
        ttk.Button(bar, text="手動加入中文文獻", command=self.open_manual).pack(side="left")
        ttk.Button(bar, text="全選", width=6, command=lambda: self.table.set_all(True)).pack(side="left", padx=(12, 2))
        ttk.Button(bar, text="全不選", width=7, command=lambda: self.table.set_all(False)).pack(side="left")
        self.var_group = tk.BooleanVar(value=False)
        tk.Checkbutton(bar, text="匯出依章節分組", variable=self.var_group, font=T.F, bg=T.BG).pack(side="left", padx=12)
        tk.Label(bar, text="參考文獻順序", font=T.F, bg=T.BG).pack(side="left")
        self.var_order = tk.StringVar(value=ORDER_CHOICES[0])
        ttk.Combobox(bar, textvariable=self.var_order, values=ORDER_CHOICES, state="readonly",
                     width=30).pack(side="left", padx=4)
        ttk.Button(bar, text="開啟輸出資料夾", command=self.open_out_dir).pack(side="right")
        ttk.Button(bar, text="下載 OA 全文", command=self.start_oa).pack(side="right", padx=6)
        tk.Button(bar, text="匯出全部", font=T.FB, bg=T.DEEP, fg="white", activebackground=T.NAVY,
                  activeforeground="white", width=10, command=self.start_export).pack(side="right")

    def _init_sash(self, pw):
        if getattr(self, "_sash_done", False):
            return
        w = pw.winfo_width()
        if w > 400:
            pw.sashpos(0, int(w * 0.66))
            self._sash_done = True

    # =====================================================================
    def chapters(self):
        return self.cfg.get("chapters") or BK.DEFAULT_CHAPTERS

    def log(self, msg):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"[{_dt.datetime.now():%H:%M:%S}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def status(self, msg):
        self.var_status.set(msg)

    def browse_dir(self):
        d = filedialog.askdirectory(parent=self.root, initialdir=self.var_dir.get() or None, title="選擇存檔資料夾")
        if d:
            self.var_dir.set(os.path.normpath(d))

    def settings_from_gui(self):
        lang = dict(LANGS).get(self.var_lang.get(), "any")
        types = [k for k, v in self.var_types.items() if v.get()] or ["journal-article"]
        try:
            yf, yt = int(self.var_yf.get()), int(self.var_yt.get())
        except (tk.TclError, ValueError):
            yf, yt = 2016, _dt.date.today().year
        groups = parse_groups(self.txt_kw.get("1.0", "end"))
        return {
            "checks": {k: v.get() for k, v in self.var_checks.items()},
            "types": types, "year_from": yf, "year_to": yt, "classic": self.var_classic.get(),
            "lang": lang, "weights": {k: float(v.get()) for k, v in self.var_w.items()},
            "keywords": groups, "groups": groups,
            "keep_mode": self.var_keep.get(),
        }

    def _remember(self):
        s = self.settings_from_gui()
        self.cfg.update({k: s[k] for k in ("checks", "types", "year_from", "year_to", "classic", "lang",
                                           "weights", "keep_mode")})
        try:
            self.cfg["limit"] = int(self.var_limit.get())
            self.cfg["gap_threshold"] = int(self.var_thr.get())
        except (tk.TclError, ValueError):
            pass
        email = self.var_email.get().strip()
        if email:
            self.cfg["email"] = self.cfg["mailto"] = email
        self.cfg["last_dir"] = self.var_dir.get().strip()
        try:
            BK.call(BK.config_save, self.cfg)
        except Exception:  # noqa: BLE001
            pass

    def _busy(self, on):
        self.btn_search.config(state="disabled" if on else "normal")
        self.btn_gap.config(state="disabled" if on else "normal")
        self.btn_stop.config(state="normal" if on else "disabled")

    def _start(self, title, job, on_done):
        if self.task is not None and not self.task.finished:
            dialogs.warn(self.root, "請稍候", "目前已有作業在執行，請等它完成或按「停止」。")
            return None
        self._busy(True)
        self.status(title + "…")

        def done(res):
            self._busy(False)
            on_done(res)

        def cancelled():
            self._busy(False)
            self.status(f"{title}：已取消（已取得的結果保留）")
            self.log(f"{title}：已取消")
            dialogs.info(self.root, "已取消", f"「{title}」已取消；已取得的結果保留，可照常匯出。")

        def failed(e, tb):
            self._busy(False)
            self.status(f"{title}：失敗")
            self.log(f"{title} 失敗：{e}")
            net = any(s in type(e).__name__ for s in ("Connection", "Timeout", "HTTP", "Request", "URL"))
            msg = (f"「{title}」失敗：{e}\n\n"
                   + ("看起來是網路或資料庫連線問題。請確認網路後再試；已取得的結果保留，可照常匯出。"
                      if net else "已取得的結果保留，可照常匯出。"))
            dialogs.error(self.root, "發生錯誤", msg, detail=tb)

        self.task = run_task(self.root, title, job, on_done=done, on_cancel=cancelled, on_error=failed)
        return self.task

    def stop(self):
        if self.task is not None and not self.task.finished:
            self.task.cancel()
            self.status("正在停止…")

    # ---------------- 檢索 ----------------
    def start_search(self):
        groups = parse_groups(self.txt_kw.get("1.0", "end"))
        if not groups:
            dialogs.warn(self.root, "缺少關鍵字", "請至少輸入一組關鍵字（每組一行，同義詞以「;」分隔）。")
            return
        self._remember()
        s = self.settings_from_gui()
        email = self.var_email.get().strip()
        limit = int(self.var_limit.get() or 50)
        yf = None if s["classic"] else s["year_from"]
        yt = s["year_to"]
        self.query_info = {"groups": groups, "year_from": yf, "year_to": yt, "lang": s["lang"],
                           "limit_per_group": limit, "time": _dt.datetime.now().isoformat(timespec="seconds")}
        self.log(f"開始檢索：{len(groups)} 組關鍵字；年 {yf or '不限'}–{yt}；語言 {s['lang']}；每組上限 {limit}")

        mode = self.var_mode.get()
        self.query_info["mode"] = "各組分別（聯集）" if mode == "union" else "所有組同時符合（交集）"
        queries = [["; ".join(g)] for g in groups] if mode == "union" else [["; ".join(g) for g in groups]]

        self._api_errors = []

        def job(progress, cancel):
            merged, seen = [], {}
            n = len(queries)
            for gi, q in enumerate(queries):
                def sub(msg, frac, gi=gi):
                    f = None if frac is None else (gi + float(frac)) / n * 0.6
                    progress((f"第 {gi + 1}/{n} 組：" if n > 1 else "") + str(msg), f)
                found = BK.call(BK.search, q, yf, yt, s["lang"], limit, progress=sub, cancel_event=cancel,
                                mailto=email or None)
                errs = getattr(sys.modules.get("core.search"), "last_errors", None) or []
                self._api_errors.extend(str(x) for x in errs)
                for p in found:
                    key = (p.doi or "").lower() or ("t:" + (p.title or "").lower())
                    if key in seen:
                        continue
                    seen[key] = p
                    merged.append(p)
            progress(f"合併去重後 {len(merged)} 筆；補充 OA／被引／DOI 驗證…", 0.6)

            def sub2(msg, frac):
                progress(msg, None if frac is None else 0.6 + float(frac) * 0.35)
            merged = BK.call(BK.enrich, merged, email, progress=sub2, cancel_event=cancel)
            progress("計算指標與相關性…", 0.97)
            merged = BK.call(BK.evaluate, merged, s)
            return merged

        def done(found):
            manual = [p for p in self.papers if p.source == "manual"]
            found = sorted(found, key=lambda p: -p.relevance)
            self.papers = found + manual
            for i, p in enumerate(self.papers):
                p.order = i
                if not p.chapter:
                    p.chapter = "未分類"
            self.refresh_results()
            self.log(f"檢索完成：{len(found)} 筆（另有手動加入 {len(manual)} 筆）")
            if self._api_errors:
                self.log("部分來源失敗：" + "；".join(dict.fromkeys(self._api_errors)))
                dialogs.warn(self.root, "部分來源失敗",
                             f"檢索完成 {len(found)} 筆，但有部分資料庫連線失敗（結果可能不完整）。\n"
                             "已取得的結果照常顯示與匯出；可稍後再試。",
                             detail="\n".join(dict.fromkeys(self._api_errors)))
            self.status(f"檢索完成：{len(found)} 筆")
            self.nb.select(self.tab_res)
        self._start("檢索文獻", job, done)

    # ---------------- 缺口 ----------------
    def start_gap(self):
        groups = parse_groups(self.txt_kw.get("1.0", "end"))
        if len(groups) < 2:
            dialogs.warn(self.root, "關鍵字組不足", "缺口矩陣至少需要 2 組關鍵字（每組一行）。")
            return
        self._remember()
        thr = int(self.var_thr.get() or 5)
        try:
            from core.gap import estimate_requests
            need = estimate_requests(len(groups))
        except Exception:  # noqa: BLE001
            need = None
        self.log(f"缺口矩陣：{len(groups)} 組、門檻 {thr} 篇"
                 + (f"；預計 {need} 次 OpenAlex 檢索（免金鑰每日約 100 次額度）" if need else ""))

        mailto = self.var_email.get().strip() or None     # Tk 變數只能在主執行緒讀

        def job(progress, cancel):
            return BK.call(BK.gap_matrix, groups, thr, progress=progress, cancel_event=cancel, mailto=mailto)

        def done(gap):
            self.gap = gap
            self.show_gap()
            n_gap = sum(1 for r in gap.get("pairs", []) + gap.get("triples", []) if self._is_gap(r, thr))
            self.log(f"缺口矩陣完成：{len(gap.get('pairs', []))} 個兩兩、{len(gap.get('triples', []))} 個三三組合，"
                     f"缺口 {n_gap} 個")
            self.status(f"缺口矩陣完成：缺口 {n_gap} 個")
            if gap.get("errors"):
                self.log("缺口矩陣部分請求失敗：" + "；".join(map(str, gap["errors"]))[:300])
            if gap.get("budget_exhausted"):
                dialogs.warn(self.root, "OpenAlex 今日額度已用完",
                             "OpenAlex 免金鑰查詢每天約 100 次（台灣時間早上 8 點重置），今天的額度已用完，"
                             "部分組合沒有取得篇數。\n\n已取得的結果照常顯示與匯出；請明天再按一次「缺口矩陣」。",
                             detail="\n".join(map(str, gap.get("errors") or [])))
            self.nb.select(self.tab_ind)
        self._start("缺口矩陣", job, done)

    def _is_gap(self, r, thr=None):
        if r.get("count") is None:
            return False
        if "gap" in r:
            return bool(r["gap"])
        thr = thr if thr is not None else (self.gap or {}).get("threshold", 5)
        return (r.get("count") or 0) < thr

    def _combo_text(self, idx):
        groups = (self.gap or {}).get("groups") or parse_groups(self.txt_kw.get("1.0", "end"))
        parts = []
        for i in idx:
            g = groups[i] if i < len(groups) else [f"組{i + 1}"]
            parts.append(f"[{i + 1}] {g[0]}")
        return " × ".join(parts)

    def show_gap(self):
        t = self.gap_tree
        t.delete(*t.get_children())
        self._gap_rows = {}
        if not self.gap:
            return
        for kind, key in (("兩兩", "pairs"), ("三三", "triples")):
            for r in self.gap.get(key, []):
                g = self._is_gap(r)
                failed = r.get("count") is None
                iid = t.insert("", "end", values=(kind, self._combo_text(r.get("groups", ())),
                                                  "未取得" if failed else r.get("count", 0),
                                                  "查詢失敗" if failed else ("缺口" if g else "—")),
                               tags=("gap",) if g else ())
                self._gap_rows[iid] = r
        self._draw_trend()

    def _gap_selected(self, _e=None):
        s = self.gap_tree.selection()
        self.lst_top.delete(0, "end")
        if not s:
            return
        r = self._gap_rows.get(s[0], {})
        for p in (r.get("top") or [])[:10]:
            if isinstance(p, dict):
                p = BK.Paper.from_dict(p)
            self.lst_top.insert("end", f"{p.year or ''}  {p.title}  ({p.container})")
        if not r.get("top"):
            self.lst_top.insert("end", "（此組合沒有命中文獻）")

    def _draw_trend(self):
        cv = self.cv_trend
        cv.delete("all")
        trend = (self.gap or {}).get("trend") or {}
        if not trend:
            cv.create_text(10, 60, anchor="w", text="尚無資料（按「缺口矩陣」後顯示）", fill=T.GREY, font=T.FS)
            return
        items = sorted((int(k), int(v)) for k, v in trend.items())[-10:]
        W = max(cv.winfo_width(), 300)
        H = 130
        mx = max(v for _, v in items) or 1
        n = len(items)
        x0, x1, y0, y1 = 40, W - 14, 12, H - 22
        pts = []
        for i, (y, v) in enumerate(items):
            x = x0 + (x1 - x0) * (i / max(1, n - 1))
            yy = y1 - (y1 - y0) * v / mx
            pts.append((x, yy))
            cv.create_text(x, H - 10, text=str(y), font=T.FS, fill=T.GREY)
            cv.create_text(x, yy - 8, text=str(v), font=T.FS, fill=T.DEEP)
        cv.create_line(x0, y1, x1, y1, fill="#CBD5E1")
        if len(pts) > 1:
            cv.create_line(*[c for p in pts for c in p], fill=T.TEAL, width=2)
        for x, yy in pts:
            cv.create_oval(x - 3, yy - 3, x + 3, yy + 3, fill=T.GOLD, outline="")

    # ---------------- 指標 ----------------
    def apply_indicators(self):
        self._remember()
        if not self.papers:
            dialogs.info(self.root, "套用指標", "目前沒有文獻；設定已記住，檢索後會自動套用。")
            return
        try:
            api = [p for p in self.papers if p.source != "manual"]
            BK.call(BK.evaluate, api, self.settings_from_gui())
        except Exception as e:  # noqa: BLE001
            dialogs.error(self.root, "套用指標", f"指標計算失敗：{e}")
            return
        self.refresh_results()
        self.status("已套用指標")
        self.nb.select(self.tab_res)

    def visible_papers(self):
        if self.var_keep.get() == "strict":
            return [p for p in self.papers if p.source == "manual" or p.passes_all()]
        return list(self.papers)

    def refresh_results(self):
        vis = sorted(self.visible_papers(), key=lambda p: p.order)
        self.table.set_papers(vis)
        self._update_count()

    def _update_count(self):
        vis = self.table.papers
        n_inc = sum(1 for p in vis if p.include)
        hidden = len(self.papers) - len(vis)
        self.var_count.set(f"顯示 {len(vis)} 筆、納入 {n_inc} 筆" + (f"、隱藏未通過 {hidden} 筆" if hidden else ""))
        # 保持全體 order 與畫面一致（隱藏者排在後面）
        shown = {id(p) for p in vis}
        rest = [p for p in self.papers if id(p) not in shown]
        for i, p in enumerate(rest):
            p.order = len(vis) + i

    def show_detail(self, p):
        t = self.txt_detail
        t.config(state="normal")
        t.delete("1.0", "end")
        if p is None:
            t.config(state="disabled")
            return
        try:
            apa = BK.format_apa7(p)
            cite = BK.in_text(p)
        except Exception as e:  # noqa: BLE001
            apa, cite = f"（格式化失敗：{e}）", ""
        t.insert("end", "APA 7\n", "h")
        t.insert("end", apa + "\n\n")
        t.insert("end", "文內引用　", "h")
        t.insert("end", cite + "\n\n")
        t.insert("end", "指標詳情\n", "h")
        for k, (v, note) in ((k, (x if isinstance(x, (list, tuple)) else (x, ""))) for k, x in p.indicators.items()):
            if k == "relevance":
                mark, tag = "•", "h"
            else:
                mark, tag = ("✔", "ok") if v is True else (("✘", "bad") if v is False else ("?", "q"))
            t.insert("end", f"{mark} ", tag)
            t.insert("end", f"{BK.INDICATOR_NAMES.get(k, k)}：{note}\n")
        t.insert("end", f"\n相關性 {p.relevance:.2f}　被引 {p.cited_by}　OA {p.oa_status or '—'}"
                        f"{'　授權 ' + p.oa_license if p.oa_license else ''}\n")
        t.insert("end", f"類型 {BK.TYPE_LABELS.get(p.type, p.type)}　來源 {p.source}\n")
        if p.source == "manual":
            t.insert("end", "使用者自填，未經 API 驗證\n", "q")
        if getattr(p, "tldr", ""):
            t.insert("end", "\nTLDR　", "h")
            t.insert("end", p.tldr + "\n")
        t.config(state="disabled")

    # ---------------- 手動加入 ----------------
    def open_manual(self):
        def add(p):
            p.order = len(self.papers)
            try:
                BK.call(BK.evaluate, [p], self.settings_from_gui())
            except Exception:  # noqa: BLE001
                pass
            self.papers.append(p)
            self.refresh_results()
            self.status(f"已手動加入：{p.title[:30]}")
        return ManualAddWindow(self.root, self.chapters(), on_add=add)

    # ---------------- 匯出 ----------------
    def export_name(self):
        groups = (self.query_info or {}).get("groups") or parse_groups(self.txt_kw.get("1.0", "end"))
        try:
            from core.export import default_name
            if groups:
                return default_name(groups)
        except Exception:  # noqa: BLE001
            pass
        kw = groups[0][0] if groups else "手動"
        return f"文獻缺口_{safe_name(kw)}_{_dt.date.today():%Y%m%d}"

    def export_papers(self):
        vis = [p for p in self.table.papers if p.include]
        return sorted(vis, key=lambda p: p.order)

    def _out_dir(self):
        d = self.var_dir.get().strip()
        if not d:
            self.browse_dir()
            d = self.var_dir.get().strip()
        if not d:
            return None
        os.makedirs(d, exist_ok=True)
        return d

    def start_export(self):
        papers = self.export_papers()
        if not papers:
            dialogs.warn(self.root, "匯出", "沒有被勾選「納入」的文獻。")
            return
        out = self._out_dir()
        if not out:
            return
        self._remember()
        name = self.export_name()
        groups = (self.query_info or {}).get("groups") or parse_groups(self.txt_kw.get("1.0", "end"))
        options = {"group_by_chapter": self.var_group.get(), "chapters": list(self.chapters()),
                   "order": "custom" if self.var_order.get() == ORDER_CHOICES[1] else "apa",
                   "keywords": groups, "only_included": True, "pdf_engine": "auto",
                   "query": dict(self.query_info or {}), "settings": self.settings_from_gui(),
                   "mailto": self.cfg.get("mailto") or self.var_email.get().strip(), "version": self.version}

        def job(progress, cancel):
            return BK.call(BK.export_all, papers, self.gap, out, name, options, progress=progress,
                           cancel_event=cancel)

        def done(res):
            res = res or {}
            ok = [p for p, v in res.items() if v is True or (isinstance(v, (list, tuple)) and v and v[0])]
            bad = {p: v for p, v in res.items() if p not in ok}
            self.last_export_dir = os.path.dirname(ok[0]) if ok else out
            lines = "\n".join("・" + os.path.basename(p) for p in ok)
            self.log(f"匯出完成：成功 {len(ok)} 個檔案，失敗 {len(bad)} 個 → {self.last_export_dir}")
            self.status(f"匯出完成：{len(ok)} 個檔案")
            if bad:
                dialogs.warn(self.root, "匯出完成（部分失敗）",
                             f"成功 {len(ok)} 個，失敗 {len(bad)} 個。\n\n{lines}",
                             detail="\n".join(f"{os.path.basename(p)}：{v}" for p, v in bad.items()))
            else:
                if dialogs.ask(self.root, "匯出完成", f"已匯出 {len(ok)} 個檔案到：\n{self.last_export_dir}\n\n{lines}\n\n要開啟資料夾嗎？",
                               yes="開啟資料夾", no="關閉"):
                    self.open_out_dir()
        self._start("匯出全部", job, done)

    def start_oa(self):
        papers = [p for p in self.export_papers() if p.oa_pdf_url]
        if not papers:
            dialogs.info(self.root, "下載 OA 全文", "納入的文獻中沒有可合法下載的 OA PDF。\n非 OA 文獻請用 DOI 連結到出版社頁面。")
            return
        out = self._out_dir()
        if not out:
            return

        def job(progress, cancel):
            return BK.call(BK.download_oa, papers, out, progress=progress, cancel_event=cancel)

        def done(rows):
            rows = rows or []
            okn = sum(1 for r in rows if str(r.get("status", "")).lower() in ("ok", "downloaded", "cached", "exists"))
            self.last_export_dir = out
            self.log(f"OA 全文：{len(papers)} 筆嘗試，成功 {okn} 筆；紀錄見 OA下載紀錄.csv")
            self.status(f"OA 全文下載完成：{okn}/{len(papers)}")
            dialogs.info(self.root, "下載 OA 全文", f"嘗試 {len(papers)} 筆，成功 {okn} 筆。\n"
                         "PDF 存於「OA全文」資料夾，來源與授權記錄在「OA下載紀錄.csv」。")
        self._start("下載 OA 全文", job, done)

    def open_out_dir(self):
        d = self.last_export_dir or self.var_dir.get().strip()
        if d and os.path.isdir(d):
            open_path(d)
        else:
            dialogs.warn(self.root, "開啟資料夾", "資料夾不存在，請先選擇存檔資料夾或匯出一次。")

    # ---------------- 子視窗 ----------------
    def open_settings(self):
        def saved(cfg):
            self.cfg = cfg
            self.var_email.set(cfg.get("email", ""))
            if cfg.get("last_dir"):
                self.var_dir.set(cfg["last_dir"])
            for k, v in (cfg.get("weights") or {}).items():
                if k in self.var_w:
                    self.var_w[k].set(v)
            self.table.refresh()
            self.status("設定已儲存")
        cfg = dict(self.cfg)
        cfg["email"] = self.var_email.get().strip() or cfg.get("email", "")
        cfg["last_dir"] = self.var_dir.get().strip()
        return SettingsWindow(self.root, cfg, on_save=saved)

    def open_about(self):
        return AboutWindow(self.root, self.version, BK.STATUS)

    def on_close(self):
        if self.task is not None and not self.task.finished:
            if not dialogs.ask(self.root, "結束", "還有作業在執行，確定要結束嗎？"):
                return
            self.task.cancel()
        try:
            self._remember()
        except Exception:  # noqa: BLE001
            pass
        self.root.destroy()
