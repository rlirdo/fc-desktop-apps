# -*- coding: utf-8 -*-
"""FC文獻缺口搜尋工具 FC LitGap Finder

雙擊執行 → 開啟視窗（檢索／指標與缺口／結果三分頁）。

命令列：
  FCLitGapFinder.exe --selftest              → 離線自我測試（樣本→指標→缺口→匯出全部格式→隱私檢查→
                                               介面稽核），寫 %TEMP%\\FCLitGapFinder_selftest.log，exit 0/1
  FCLitGapFinder.exe --check-privacy [資料夾] → 檢查輸出檔不含 email 以外個資（不給資料夾＝用樣本匯出後檢查）
  FCLitGapFinder.exe --version               → 顯示版本

環境變數：
  LITGAP_BACKEND=auto|real|stub   核心模組選擇（預設 auto：真模組可用就用）
  LITGAP_NO_DIALOG=1              對話框不顯示、直接回傳預設值（自動化用）
  FCLITGAP_HOME=資料夾            設定檔與快取位置（預設 %APPDATA%\\FCLitGap）
"""
import datetime as _dt
import getpass
import json
import os
import re
import shutil
import socket
import sys
import tempfile
import time
import traceback
import zipfile

APP_NAME = "FC文獻缺口搜尋工具"
APP_EN = "FC LitGap Finder"
EXE_NAME = "FCLitGapFinder"
VERSION = "1.0.0"
CREDIT = "NDHU自資系游豐兆製作"

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

_LOG_LINES = []


def out(msg=""):
    """console=False 封裝時 stdout 可能不存在；一律記錄到 log，有 console 就順便印。"""
    _LOG_LINES.append(str(msg))
    try:
        if sys.stdout is not None:
            print(msg)
            sys.stdout.flush()
    except Exception:  # noqa: BLE001
        pass


def write_log(name):
    path = os.path.join(tempfile.gettempdir(), name)
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(_LOG_LINES) + "\n")
    except OSError:
        return None
    return path


# ===========================================================================
# 隱私檢查（--check-privacy）
# ===========================================================================
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
STUDENT_ID_RE = re.compile(r"(?<![\d.])(?:41|61|81)\d{7}(?![\d@])")
TW_ID_RE = re.compile(r"(?<![A-Za-z0-9])[A-Z][12]\d{8}(?!\d)")
PHONE_RE = re.compile(r"(?<!\d)09\d{2}-?\d{3}-?\d{3}(?!\d)")
TEXT_EXT = (".json", ".ris", ".bib", ".csv", ".txt", ".md", ".log")
ZIP_EXT = (".docx", ".xlsx", ".pptx")
RAW_ONLY_WHO = "#rawonly"


def _file_texts(path):
    """回傳 [(部位, 文字)]。docx/xlsx 讀所有 XML（含 docProps 作者欄），PDF 讀原始位元組（含中繼資料）。"""
    low = path.lower()
    try:
        if low.endswith(TEXT_EXT):
            return [("text", open(path, encoding="utf-8", errors="ignore").read())]
        if low.endswith(ZIP_EXT):
            z = zipfile.ZipFile(path)
            return [(n, z.read(n).decode("utf-8", "ignore")) for n in z.namelist()
                    if n.endswith((".xml", ".rels"))]
        if low.endswith(".pdf"):
            raw = open(path, "rb").read()
            lat = raw.decode("latin-1", "ignore")
            # 中繼資料（Info 字典＋XMP）做完整檢查；整檔原始位元組只比對帳號／路徑（避免二進位誤判）
            meta = re.findall(r"/(?:Author|Creator|Producer|Title|Subject|Keywords)\s*(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f]*>)",
                              lat)
            meta += re.findall(r"<x:xmpmeta.*?</x:xmpmeta>", lat, flags=re.S)
            dec = []
            for m in meta:
                if m.startswith("<") and not m.startswith("<x:"):
                    try:
                        b = bytes.fromhex(m[1:-1])
                        dec.append(b.decode("utf-16") if b[:2] in (b"\xfe\xff", b"\xff\xfe") else b.decode("latin-1"))
                    except ValueError:
                        pass
                elif m.startswith("(") and m[1:3] == "\xfe\xff":
                    dec.append(m[1:-1].encode("latin-1", "ignore").decode("utf-16", "ignore"))
                else:
                    dec.append(m)
            parts = [("pdf-meta", "\n".join(dec)), ("pdf-raw" + RAW_ONLY_WHO, lat)]
            try:
                parts.append(("pdf-utf16" + RAW_ONLY_WHO, raw.decode("utf-16-be", "ignore")))
            except Exception:  # noqa: BLE001
                pass
            try:
                from pypdf import PdfReader  # 選用：有裝才抽文字
                r = PdfReader(path)
                parts.append(("pdf-text", "\n".join((pg.extract_text() or "") for pg in r.pages)))
                parts.append(("pdf-meta", json.dumps({k: str(v) for k, v in (r.metadata or {}).items()},
                                                     ensure_ascii=False)))
            except Exception:  # noqa: BLE001
                pass
            return parts
    except Exception as e:  # noqa: BLE001
        return [("error", f"<讀取失敗 {e}>")]
    return []


def _json_email_violations(path):
    """查詢紀錄 json：email 只允許出現在 key 名含 mailto 的欄位。"""
    bad = []
    try:
        data = json.load(open(path, encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return bad

    def walk(o, key=""):
        if isinstance(o, dict):
            for k, v in o.items():
                walk(v, str(k))
        elif isinstance(o, list):
            for v in o:
                walk(v, key)
        elif isinstance(o, str):
            if EMAIL_RE.search(o) and "mailto" not in key.lower():
                bad.append(f"{os.path.basename(path)}：欄位「{key}」含 email")
    walk(data)
    return bad


def privacy_scan(folder):
    """回傳 (檢查檔數, 問題清單)。"""
    problems = []
    who = set()
    try:
        u = getpass.getuser()
        if u and len(u) >= 2:
            who.add(u)
    except Exception:  # noqa: BLE001
        pass
    for k in ("USERNAME", "COMPUTERNAME"):
        v = os.environ.get(k, "")
        if v and len(v) >= 2 and v.lower() not in ("runneradmin", "runner", "user", "admin"):
            who.add(v)
    try:
        h = socket.gethostname()
        if h and len(h) >= 3:
            who.add(h)
    except Exception:  # noqa: BLE001
        pass
    home = os.path.expanduser("~")
    n = 0
    for root, _dirs, files in os.walk(folder):
        for f in files:
            p = os.path.join(root, f)
            texts = _file_texts(p)
            if not texts:
                continue
            n += 1
            is_query_json = f.lower().endswith(".json")
            for part, t in texts:
                where = f"{f}［{part.replace(RAW_ONLY_WHO, '')}］"
                if part.endswith(RAW_ONLY_WHO):
                    for w in who:
                        if len(w) >= 3 and w in t:
                            problems.append(f"{where} 含本機帳號／電腦名稱「{w}」")
                    continue
                if not is_query_json:
                    for m in set(EMAIL_RE.findall(t)):
                        problems.append(f"{where} 含 email：{m}")
                for m in set(STUDENT_ID_RE.findall(t)):
                    problems.append(f"{where} 含疑似學號：{m}")
                for m in set(TW_ID_RE.findall(t)):
                    problems.append(f"{where} 含疑似身分證字號：{m}")
                for m in set(PHONE_RE.findall(t)):
                    problems.append(f"{where} 含疑似手機號碼：{m}")
                for w in who:
                    if w in t:
                        problems.append(f"{where} 含本機帳號／電腦名稱「{w}」")
                if home and len(home) > 4 and (home in t or home.replace("\\", "/") in t
                                               or home.replace("\\", "\\\\") in t):
                    problems.append(f"{where} 含使用者家目錄路徑")
            if is_query_json:
                problems.extend(_json_email_violations(p))
    return n, sorted(set(problems))


# ===========================================================================
# 自我測試
# ===========================================================================
EXPECT_SUFFIX = ["參考文獻_APA7.docx", "參考文獻_APA7.pdf", "文獻清單與指標.xlsx", "缺口矩陣.xlsx",
                 "趨勢圖.png", "references.ris", "references.bib", "查詢紀錄.json"]


def _sample_pipeline(BK, tmp, log, cancel=None):
    """樣本→指標→缺口→匯出；回傳 (papers, gap, export_result, out_dir)。"""
    papers = BK.load_samples()
    log(f"  樣本 {len(papers)} 筆（中文 {sum(1 for p in papers if p.is_chinese)} 筆、"
        f"≥21 作者 {sum(1 for p in papers if len(p.authors) >= 21)} 筆）")
    if len(papers) < 15:
        raise AssertionError(f"樣本少於 15 筆：{len(papers)}")
    # 離線「檢索」：stub.search 走同一份樣本，確認 progress／cancel 介面可用
    from core import stub
    msgs = []
    got = stub.search(["serpentine"], 2010, None, "any", 50, progress=lambda m, f: msgs.append(m),
                      cancel_event=cancel)
    log(f"  離線檢索介面：{len(got)} 筆、progress 回呼 {len(msgs)} 次")
    settings = {"checks": {k: True for k in ("doi", "oa", "type", "year", "issn", "lang")},
                "types": ["journal-article", "proceedings-article", "book-chapter"],
                "year_from": 2016, "year_to": None, "classic": True, "lang": "any",
                "weights": {"title": 3.0, "abstract": 1.0, "cited": 1.0, "recency": 1.0},
                "keywords": ["serpentine", "deep sea water", "far infrared"], "keep_mode": "mark"}
    papers = BK.call(BK.evaluate, papers, settings)
    n_ind = sum(1 for p in papers if p.indicators)
    log(f"  指標：{n_ind}/{len(papers)} 筆有指標；相關性最高 {max(p.relevance for p in papers):.2f}")
    if n_ind != len(papers):
        raise AssertionError("有文獻沒有指標結果")
    bad_apa = []
    for p in papers:
        s = BK.format_apa7(p)
        if not s or len(s) < 20:
            bad_apa.append(p.title[:30])
        BK.in_text(p)
    if bad_apa:
        raise AssertionError(f"APA 7 格式化失敗：{bad_apa}")
    log(f"  APA 7：{len(papers)} 筆皆可格式化；例：{BK.format_apa7(papers[0])[:90]}…")
    groups = [["serpentine", "蛇紋石"], ["deep sea water", "深層海水"], ["far infrared", "遠紅外線"]]
    gap = stub.gap_matrix(groups, 5, cancel_event=cancel)   # 離線假資料（真模組需連網）
    log(f"  缺口矩陣（離線）：兩兩 {len(gap['pairs'])}、三三 {len(gap['triples'])}、趨勢 {len(gap['trend'])} 年")
    for i, p in enumerate(papers):
        p.order = i
        p.chapter = ["1 緒論", "2.1 文獻回顧", "3 方法", "未分類"][i % 4]
    out_dir = os.path.join(tmp, "匯出 [測試] 資料夾")   # 刻意含空白與中括號
    os.makedirs(out_dir, exist_ok=True)
    name = f"文獻缺口_serpentine_{_dt.date.today():%Y%m%d}"
    options = {"group_by_chapter": True, "chapters": ["1 緒論", "2.1 文獻回顧", "3 方法", "未分類"],
               "order": "apa", "keywords": groups, "only_included": True,
               "pdf_engine": os.environ.get("LITGAP_PDF_ENGINE", "reportlab"),
               "query": {"groups": groups, "year_from": 2016, "year_to": None, "lang": "any",
                         "limit_per_group": 50, "time": _dt.datetime.now().isoformat(timespec="seconds")},
               "settings": settings, "mailto": "selftest@example.org", "version": VERSION}
    t0 = time.time()
    res = BK.call(BK.export_all, papers, gap, out_dir, name, options, progress=None, cancel_event=cancel)
    log(f"  匯出耗時 {time.time() - t0:.1f} 秒")
    return papers, gap, res or {}, out_dir


def _check_exports(res, out_dir, log):
    problems = []
    files = []
    for r, _d, fs in os.walk(out_dir):
        files += [os.path.join(r, f) for f in fs]
    for suf in EXPECT_SUFFIX:
        hit = [f for f in files if f.endswith(suf) and os.path.getsize(f) > 0]
        if hit:
            log(f"  ✔ {os.path.basename(hit[0])}（{os.path.getsize(hit[0]):,} bytes）")
        else:
            problems.append(f"缺少匯出檔：*{suf}")
            log(f"  ✘ 缺少 *{suf}")
    for p, v in res.items():
        okv = v is True or (isinstance(v, (list, tuple)) and v and v[0] is True)
        if not okv:
            problems.append(f"export_all 回報失敗：{os.path.basename(p)} → {v}")
    # docx 內容抽查：APA 懸掛縮排、Times New Roman、中文先於英文
    docx = [f for f in files if f.endswith("參考文獻_APA7.docx")]
    if docx:
        try:
            import docx as _docx
            d = _docx.Document(docx[0])
            txt = "\n".join(p.text for p in d.paragraphs)
            log(f"  docx 段落 {len(d.paragraphs)}；含「指標證據表」={'指標證據表' in txt}、"
                f"含「檢索策略」={'檢索策略' in txt}、含「使用者自填」={'使用者自填' in txt}")
            for key in ("指標證據表", "檢索策略"):
                if key not in txt and not any(key in c.text for t in d.tables for r in t.rows for c in r.cells):
                    problems.append(f"docx 缺少「{key}」")
        except Exception as e:  # noqa: BLE001
            problems.append(f"docx 讀取失敗：{e}")
    return problems


def _gui_selftest(BK, papers, gap, log):
    """實際建立視窗（不進 mainloop）：主視窗三分頁、所有子視窗類別、CREDIT 位置、拖曳排序、背景工作與取消。"""
    import tkinter as tk
    from ui import credit as C
    from ui import dialogs, main_window, progress, results, windows

    problems = []
    mods = [credit_mod for credit_mod in (C, dialogs, main_window, progress, results, windows)]
    good, bad = C.audit_classes(mods)
    log(f"  視窗類別稽核：繼承 CreditToplevel {len(good)} 個（{', '.join(g.split('.')[-1] for g in good)}）；違規 {len(bad)} 個")
    problems += [f"類別未繼承 CreditToplevel：{b}" for b in bad]
    if not getattr(sys, "frozen", False):
        srcs = [os.path.join(HERE, "app.py")] + [os.path.join(HERE, "ui", f) for f in os.listdir(os.path.join(HERE, "ui"))
                                                 if f.endswith(".py")]
        sp = C.audit_sources(srcs)
        log(f"  原始碼稽核（禁 messagebox、禁直接 Toplevel）：{len(srcs)} 檔、問題 {len(sp)} 個")
        problems += sp
    else:
        log("  原始碼稽核：封裝版無原始碼，略過（CI 原始碼階段已檢查）")

    try:
        root = tk.Tk()
    except tk.TclError as e:
        log(f"  （無圖形環境，略過實體視窗測試：{e}）")
        return problems
    try:
        mw = main_window.MainWindow(root, VERSION)
        root.update()
        tabs = [mw.nb.tab(t, "text").strip() for t in mw.nb.tabs()]
        log(f"  主視窗：{root.title()}；分頁 {tabs}")
        if len(tabs) != 3:
            problems.append("主視窗分頁數不是 3")
        ok, why = C.credit_in_bottom_right(root)
        log(f"  主視窗 CREDIT：{'✔' if ok else '✘'} {why}")
        if not ok:
            problems.append(f"主視窗 CREDIT 位置不正確：{why}")
        # 結果表：載入、勾選、章節、拖曳（以 move_item 模擬放開位置）、排序
        mw.papers = list(papers)
        mw.gap = gap
        mw.show_gap()
        mw.refresh_results()
        mw.nb.select(mw.tab_res)
        root.update()
        t = mw.table
        kids = t.tree.get_children()
        first = t._by_iid[kids[0]]
        t.move_item(kids[0], 3)
        root.update()
        if t.papers.index(first) != 2 or first.order != 2:
            problems.append("拖曳排序後 order 未更新")
        t.toggle(kids[1])
        t.set_chapter(kids[1], "3 方法")
        t.sort_by("year")
        orders = [p.order for p in t.papers]
        if orders != list(range(len(orders))):
            problems.append("欄位排序後 order 不連續")
        log(f"  結果表：{len(kids)} 列；拖曳／勾選／章節／排序 OK；{mw.var_count.get()}")
        t.tree.selection_set(t.tree.get_children()[0])
        mw.show_detail(t.selected_paper())
        if "APA 7" not in mw.txt_detail.get("1.0", "end"):
            problems.append("右側 APA 預覽未顯示")
        # 所有子視窗
        wins = [("設定", lambda: windows.SettingsWindow(root, mw.cfg, modal=False)),
                ("關於", lambda: windows.AboutWindow(root, VERSION, BK.STATUS, modal=False)),
                ("手動加入", lambda: windows.ManualAddWindow(root, mw.chapters(), modal=False)),
                ("進度", lambda: progress.ProgressWindow(root, "測試進度")),
                ("對話框", lambda: dialogs.FcDialog(root, "測試", "這是測試訊息。", kind="info", modal=False)),
                ("錯誤對話框", lambda: dialogs.FcDialog(root, "錯誤", "測試錯誤。", kind="error",
                                                    detail="Traceback…", modal=False)),
                ("詢問對話框", lambda: dialogs.FcDialog(root, "詢問", "要繼續嗎？", kind="ask", modal=False))]
        for label, mk in wins:
            w = mk()
            root.update()
            ok, why = C.credit_in_bottom_right(w)
            log(f"  {label}視窗 CREDIT：{'✔' if ok else '✘'} {why}")
            if not ok:
                problems.append(f"{label}視窗 CREDIT 位置不正確：{why}")
            w.destroy()
        # 手動加入：建立 Paper
        ma = windows.ManualAddWindow(root, mw.chapters(), modal=False)
        for k, v in {"authors": "王小明、陳大華", "year": "2021", "title": "測試題名", "container": "臺灣化學教育",
                     "volume": "40", "issue": "2", "pages": "13-19"}.items():
            ma.vars[k].set(v)
        p = ma.build_paper()
        ma.destroy()
        if p.source != "manual" or "使用者自填" not in p.note:
            problems.append("手動加入未標 source=manual")
        log(f"  手動加入：{BK.format_apa7(p)}")
        # 背景工作＋取消
        import threading

        def slow(progress_cb, cancel):
            for i in range(200):
                if cancel.is_set():
                    raise BK.Cancelled("取消")
                progress_cb(f"第 {i} 步", i / 200)
                time.sleep(0.01)
            return "done"
        res = {}
        task = progress.run_task(root, "取消測試", slow, on_done=lambda r: res.setdefault("r", r),
                                 on_cancel=lambda: res.setdefault("r", "cancel"))
        t0 = time.time()
        while time.time() - t0 < 0.3:
            root.update()
            time.sleep(0.02)
        task.cancel()
        while not task.finished and time.time() - t0 < 10:
            root.update()
            time.sleep(0.02)
        log(f"  背景工作取消：{res.get('r')}（GUI 期間持續回應，主執行緒={threading.current_thread().name}）")
        if res.get("r") != "cancel":
            problems.append("背景工作取消失敗")
        task2 = progress.run_task(root, "完成測試", lambda pc, c: 42, on_done=lambda r: res.__setitem__("r2", r))
        t0 = time.time()
        while not task2.finished and time.time() - t0 < 10:
            root.update()
            time.sleep(0.02)
        if res.get("r2") != 42:
            problems.append("背景工作完成回呼失敗")
        log("  背景工作完成回呼：OK" if res.get("r2") == 42 else "  背景工作完成回呼：失敗")
    finally:
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass
    return problems


def run_selftest(check_privacy_only=False, folder=None):
    os.environ["LITGAP_NO_DIALOG"] = "1"
    tmp = tempfile.mkdtemp(prefix="FCLitGap_selftest_")
    os.environ["FCLITGAP_HOME"] = os.path.join(tmp, "home")       # 不碰使用者真實設定與快取
    out(f"{EXE_NAME} {VERSION} — {'隱私檢查' if check_privacy_only else '自我測試'}")
    out(f"Python {sys.version.split()[0]}；封裝版={bool(getattr(sys, 'frozen', False))}；平台 {sys.platform}")
    problems = []
    try:
        if check_privacy_only and folder:
            n, pv = privacy_scan(folder)
            out(f"[隱私] 檢查 {n} 個檔案：{'OK' if not pv else '發現 %d 個問題' % len(pv)}")
            for x in pv:
                out("  ✘ " + x)
            return 0 if not pv else 1
        from ui import backend as BK
        mode = BK.init()
        out(f"[核心] {BK.STATUS}（mode={mode}）")
        if mode != "real" and os.environ.get("LITGAP_BACKEND", "auto") != "stub":
            problems.append("真核心模組無法匯入：" + BK.STATUS)
        out("[1] 樣本→指標→缺口→匯出")
        papers, gap, res, out_dir = _sample_pipeline(BK, tmp, out)
        out(f"[2] 匯出檔檢查（{out_dir}）")
        problems += _check_exports(res, out_dir, out)
        out("[3] 隱私檢查")
        n, pv = privacy_scan(out_dir)
        out(f"  檢查 {n} 個檔案；問題 {len(pv)} 個")
        for x in pv:
            out("  ✘ " + x)
        problems += pv
        if not check_privacy_only:
            out("[4] 介面稽核與實體視窗測試")
            problems += _gui_selftest(BK, papers, gap, out)
    except Exception as e:  # noqa: BLE001
        problems.append(f"例外：{type(e).__name__}: {e}")
        out(traceback.format_exc())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if problems:
        out(f"SELFTEST FAIL（{len(problems)} 個問題）")
        for x in problems:
            out("  - " + x)
        return 1
    out("SELFTEST OK" if not check_privacy_only else "PRIVACY OK")
    return 0


# ===========================================================================
def run_gui():
    import tkinter as tk
    from ui import backend as BK
    from ui import dialogs
    from ui.main_window import MainWindow
    BK.init()
    root = tk.Tk()
    try:
        MainWindow(root, VERSION)
    except Exception as e:  # noqa: BLE001
        dialogs.error(root, "啟動失敗", f"程式啟動失敗：{e}", detail=traceback.format_exc())
        return 1

    def report_callback_exception(exc, val, tb):
        dialogs.error(root, "發生錯誤", f"{exc.__name__}: {val}",
                      detail="".join(traceback.format_exception(exc, val, tb)))
    root.report_callback_exception = report_callback_exception
    root.mainloop()
    return 0


def main():
    args = sys.argv[1:]
    if "--version" in args or "-v" in args:
        out(f"{EXE_NAME} {VERSION}（{APP_NAME} {APP_EN}）")
        write_log(f"{EXE_NAME}_version.log")
        return 0
    if "--help" in args or "-h" in args:
        out(__doc__)
        return 0
    if "--selftest" in args:
        code = run_selftest()
        out(f"（測試紀錄：{os.path.join(tempfile.gettempdir(), EXE_NAME + '_selftest.log')}）")
        write_log(f"{EXE_NAME}_selftest.log")
        return code
    if "--check-privacy" in args:
        i = args.index("--check-privacy")
        folder = args[i + 1] if i + 1 < len(args) and not args[i + 1].startswith("--") else None
        code = run_selftest(check_privacy_only=True, folder=folder)
        write_log(f"{EXE_NAME}_privacy.log")
        return code
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())
