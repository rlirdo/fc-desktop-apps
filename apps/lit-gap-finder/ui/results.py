# -*- coding: utf-8 -*-
"""結果表：Treeview＋勾選納入＋章節下拉＋拖曳排序＋欄位排序＋雙擊開 DOI＋右鍵選單。

資料來源是 list[Paper]（由主視窗持有）；本元件只改 Paper.include / chapter / order，
並在變動時呼叫 on_change()。排序／拖曳後依畫面由上到下重寫 order（＝匯出順序）。
"""
import tkinter as tk
import webbrowser
from tkinter import ttk

from . import theme as T

COLS = [("inc", "納入", 46, "center"), ("seq", "序", 40, "e"), ("rel", "相關性", 64, "e"),
        ("year", "年", 54, "center"), ("author", "作者", 110, "w"), ("title", "題名", 300, "w"),
        ("journal", "刊名", 140, "w"), ("doi", "DOI", 150, "w"), ("oa", "OA", 60, "center"),
        ("ind", "指標摘要", 86, "center"), ("chapter", "章節", 100, "w")]
CHECK_ON, CHECK_OFF = "☑", "☐"


def _authors_short(p):
    if not p.authors:
        return ""
    fam = (p.authors[0].get("family") or p.authors[0].get("given") or "").strip()
    if len(p.authors) == 1:
        return fam
    return f"{fam} 等" if p.is_chinese else f"{fam} et al."


class ResultsTable(tk.Frame):
    def __init__(self, master, get_chapters, on_select=None, on_change=None):
        super().__init__(master, bg=T.BG)
        self.get_chapters = get_chapters
        self.on_select = on_select
        self.on_change = on_change
        self.papers = []          # 目前顯示的 Paper（順序＝畫面順序）
        self._by_iid = {}
        self._sort_state = {}
        self._drag = None
        self._editor = None

        self.tree = ttk.Treeview(self, columns=[c[0] for c in COLS], show="headings", selectmode="browse")
        for cid, label, w, anchor in COLS:
            self.tree.heading(cid, text=label, command=lambda c=cid: self.sort_by(c))
            self.tree.column(cid, width=w, anchor=anchor, stretch=cid in ("title", "journal"))
        ys = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure("excluded", foreground="#9AA5B1")
        self.tree.tag_configure("manual", background="#FFF8E6")
        self.tree.tag_configure("fail", foreground="#B42318")
        # 拖曳時的插入線
        self._line = tk.Frame(self.tree, height=2, bg=T.GOLD)

        self.tree.bind("<ButtonPress-1>", self._press, add="+")
        self.tree.bind("<B1-Motion>", self._motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._release, add="+")
        self.tree.bind("<Double-1>", self._double)
        self.tree.bind("<Button-3>", self._context)
        self.tree.bind("<<TreeviewSelect>>", self._selected)
        self.tree.bind("<space>", lambda e: self.toggle_selected())

        self.menu = tk.Menu(self, tearoff=0, font=T.F)

    # ------------------------------------------------------------------
    def set_papers(self, papers):
        self._close_editor()
        self.papers = list(papers)
        self.refresh()

    def refresh(self):
        sel = self.selected_paper()
        self.tree.delete(*self.tree.get_children())
        self._by_iid.clear()
        for i, p in enumerate(self.papers):
            iid = f"p{id(p)}"
            self._by_iid[iid] = p
            tags = []
            if not p.include:
                tags.append("excluded")
            if p.source == "manual":
                tags.append("manual")
            elif hasattr(p, "passes_all") and p.indicators and not p.passes_all():
                tags.append("fail")
            self.tree.insert("", "end", iid=iid, values=self._values(i, p), tags=tags)
        if sel is not None:
            iid = f"p{id(sel)}"
            if self.tree.exists(iid):
                self.tree.selection_set(iid)
                self.tree.see(iid)

    def _values(self, i, p):
        oa = "OA" if p.oa_pdf_url else (p.oa_status or "—")
        if oa == "closed":
            oa = "—"
        return (CHECK_ON if p.include else CHECK_OFF, i + 1, f"{p.relevance:.2f}", p.year or "",
                _authors_short(p), p.title, p.container, p.doi or ("(自填)" if p.source == "manual" else ""),
                oa, p.indicator_summary() if hasattr(p, "indicator_summary") else "", p.chapter or "未分類")

    def _renumber(self):
        """畫面順序 → Paper.order（匯出順序）。"""
        order = [self._by_iid[iid] for iid in self.tree.get_children()]
        self.papers = order
        for i, p in enumerate(order):
            p.order = i
            iid = f"p{id(p)}"
            self.tree.set(iid, "seq", i + 1)
        if self.on_change:
            self.on_change()

    def selected_paper(self):
        s = self.tree.selection()
        return self._by_iid.get(s[0]) if s else None

    # ---------------- 排序 ----------------
    def sort_by(self, col):
        keyf = {
            "inc": lambda p: (not p.include,), "seq": lambda p: p.order,
            "rel": lambda p: p.relevance, "year": lambda p: p.year or 0,
            "author": lambda p: _authors_short(p).lower(), "title": lambda p: (p.title or "").lower(),
            "journal": lambda p: (p.container or "").lower(), "doi": lambda p: p.doi or "",
            "oa": lambda p: bool(p.oa_pdf_url), "ind": lambda p: p.indicator_summary(),
            "chapter": lambda p: self._chapter_index(p.chapter),
        }[col]
        rev = not self._sort_state.get(col, col in ("rel", "year", "oa"))
        self._sort_state = {col: rev}
        self.papers.sort(key=keyf, reverse=rev)
        self.refresh()
        self._renumber()
        arrow = " ▼" if rev else " ▲"
        for cid, label, *_ in COLS:
            self.tree.heading(cid, text=label + (arrow if cid == col else ""))

    def _chapter_index(self, ch):
        chs = list(self.get_chapters())
        return chs.index(ch) if ch in chs else len(chs)

    # ---------------- 勾選 / 章節 ----------------
    def toggle(self, iid):
        p = self._by_iid.get(iid)
        if p is None:
            return
        p.include = not p.include
        self.tree.set(iid, "inc", CHECK_ON if p.include else CHECK_OFF)
        tags = [t for t in self.tree.item(iid, "tags") if t != "excluded"]
        if not p.include:
            tags.append("excluded")
        self.tree.item(iid, tags=tags)
        if self.on_change:
            self.on_change()

    def toggle_selected(self):
        s = self.tree.selection()
        if s:
            self.toggle(s[0])

    def set_all(self, value):
        for p in self.papers:
            p.include = value
        self.refresh()
        if self.on_change:
            self.on_change()

    def set_chapter(self, iid, chapter):
        p = self._by_iid.get(iid)
        if p is None:
            return
        p.chapter = chapter
        self.tree.set(iid, "chapter", chapter)
        if self.on_change:
            self.on_change()

    def edit_chapter(self, iid):
        """在章節儲存格上方放一個唯讀下拉選單。"""
        self._close_editor()
        bbox = self.tree.bbox(iid, "chapter")
        if not bbox:
            return
        x, y, w, h = bbox
        p = self._by_iid[iid]
        var = tk.StringVar(value=p.chapter or "未分類")
        cb = ttk.Combobox(self.tree, textvariable=var, values=list(self.get_chapters()), state="readonly")
        cb.place(x=x, y=y, width=max(w, 120), height=h)
        self._editor = cb

        def done(_e=None):
            self.set_chapter(iid, var.get())
            self._close_editor()
        cb.bind("<<ComboboxSelected>>", done)
        cb.bind("<Escape>", lambda e: self._close_editor())
        cb.bind("<FocusOut>", lambda e: self.after(150, self._close_editor))
        cb.focus_set()
        try:
            cb.event_generate("<Button-1>")
        except tk.TclError:
            pass

    def _close_editor(self):
        if self._editor is not None:
            try:
                self._editor.destroy()
            except tk.TclError:
                pass
            self._editor = None

    # ---------------- 滑鼠事件 ----------------
    def _press(self, e):
        self._drag = None
        if self.tree.identify_region(e.x, e.y) != "cell":
            return
        iid = self.tree.identify_row(e.y)
        col = self.tree.identify_column(e.x)
        if not iid:
            return
        cid = self.tree["columns"][int(col[1:]) - 1] if col else ""
        if cid == "inc":
            self.toggle(iid)
            return "break"
        if cid == "chapter":
            self.tree.selection_set(iid)
            self.after(10, lambda: self.edit_chapter(iid))
            return "break"
        self._drag = {"iid": iid, "y0": e.y, "moving": False}

    def _target_index(self, y):
        kids = self.tree.get_children()
        if not kids:
            return 0, None
        row = self.tree.identify_row(y)
        if not row:
            last = kids[-1]
            bb = self.tree.bbox(last)
            if bb and y > bb[1] + bb[3]:
                return len(kids), last
            return 0, kids[0]
        bb = self.tree.bbox(row)
        idx = self.tree.index(row)
        if bb and y > bb[1] + bb[3] / 2:
            idx += 1
        return idx, row

    def _motion(self, e):
        d = self._drag
        if not d:
            return
        if not d["moving"] and abs(e.y - d["y0"]) < 6:
            return
        d["moving"] = True
        self.tree.configure(cursor="sb_v_double_arrow")
        idx, row = self._target_index(e.y)
        kids = self.tree.get_children()
        if idx < len(kids):
            bb = self.tree.bbox(kids[idx])
            ly = bb[1] if bb else 0
        else:
            bb = self.tree.bbox(kids[-1]) if kids else None
            ly = (bb[1] + bb[3]) if bb else 0
        self._line.place(x=0, y=max(0, ly - 1), relwidth=1.0)
        self._line.lift()

    def _release(self, e):
        d = self._drag
        self._drag = None
        self._line.place_forget()
        self.tree.configure(cursor="")
        if not d or not d["moving"]:
            return
        idx, _ = self._target_index(e.y)
        self.move_item(d["iid"], idx)

    def move_item(self, iid, idx):
        """把 iid 移到畫面索引 idx（以移動前的索引計），並重寫 order。"""
        cur = self.tree.index(iid)
        if idx > cur:
            idx -= 1
        if idx == cur:
            return
        self.tree.move(iid, "", idx)
        self._sort_state = {}
        for cid, label, *_ in COLS:
            self.tree.heading(cid, text=label)
        self._renumber()
        self.tree.selection_set(iid)

    def _double(self, e):
        iid = self.tree.identify_row(e.y)
        col = self.tree.identify_column(e.x)
        cid = self.tree["columns"][int(col[1:]) - 1] if col else ""
        if iid and cid not in ("inc", "chapter"):
            p = self._by_iid.get(iid)
            if p and p.doi:
                webbrowser.open(p.doi_url)

    def _selected(self, _e=None):
        if self.on_select:
            self.on_select(self.selected_paper())

    def _context(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        p = self._by_iid[iid]
        m = self.menu
        m.delete(0, "end")
        m.add_command(label="開 OA 全文", state="normal" if p.oa_pdf_url else "disabled",
                      command=lambda: webbrowser.open(p.oa_pdf_url))
        m.add_command(label="開 DOI 連結", state="normal" if p.doi else "disabled",
                      command=lambda: webbrowser.open(p.doi_url))
        m.add_separator()
        m.add_command(label="排除此筆" if p.include else "納入此筆", command=lambda: self.toggle(iid))
        sub = tk.Menu(m, tearoff=0, font=T.F)
        for ch in self.get_chapters():
            sub.add_command(label=ch, command=lambda c=ch: self.set_chapter(iid, c))
        m.add_cascade(label="章節標籤", menu=sub)
        m.add_separator()
        m.add_command(label="移到最上", command=lambda: self.move_item(iid, 0))
        m.add_command(label="移到最下", command=lambda: self.move_item(iid, len(self.tree.get_children())))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()
