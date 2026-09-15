# -*- coding: utf-8 -*-
"""
pipeline.py — 「學生作業文字雲」流程核心（可被 GUI 或 CLI 呼叫）。

原始版本是命令列工具 ta_wordcloud.py；這裡把流程改寫成純函式：
  * 所有訊息都走 log callback（GUI 可以即時顯示在文字框），不直接 print。
  * 出錯丟 PipelineError，不呼叫 sys.exit（GUI 不能被關掉）。
  * 所有資料檔（jieba 字典、sample_input、logo、config.json）一律用 resource_path()
    定位，PyInstaller onefile 打包後會落在 sys._MEIPASS。

兩步流程（2.0）：
  ① 先用「姓名遮罩與學生編號」（NameMasker 2.0）處理 Zuvio 匯出的 xlsx
  ② 把它的輸出放進 input 資料夾再跑本程式
  （沒有先跑第 ① 步也可以，本程式會用 core/mask.py 自己補上同一套規則。）

隱私鐵律（不可關閉）：
  * 學生一律用學生編號表示（例 115-1_EC_3）。
  * 姓名遮罩（第 2 字改 O），學號與電子郵件每個字元都改成 O，長度不變。
  * 自由文字裡的名冊姓名換成該生的學生編號。
  * 「學生編號連結姓名」工作表（再識別鑰匙）解析時一律忽略，絕不讀入。
  * 原始 xlsx 只留在使用者自己的輸入資料夾，工具不會上傳任何資料。
"""
import os
import sys
import json
import shutil
import tempfile
import datetime as dt

APP_NAME = "HomeworkWordCloud"
APP_TITLE = "學生作業文字雲"
VERSION = "2.0.0"

# --check-privacy 掃描時要看的純文字副檔名
TEXT_EXT = {".csv", ".json", ".md", ".txt"}


class PipelineError(Exception):
    """使用者看得懂的錯誤（GUI 會直接顯示訊息，不印 traceback）。"""


# ------------------------------------------------------------------ 路徑
def is_frozen():
    return bool(getattr(sys, "frozen", False))


def resource_path(*parts):
    """打包後的唯讀資料檔位置（onefile 解壓在 sys._MEIPASS）。"""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def app_home():
    """使用者設定檔存放處（不放在 exe 旁邊，避免 Program Files 沒有寫入權限）。"""
    p = os.path.join(os.path.expanduser("~"), ".homework_wordcloud")
    os.makedirs(p, exist_ok=True)
    return p


def default_output_root():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.path.expanduser("~")
    return os.path.join(desktop, f"{APP_NAME}_輸出")


def open_folder(path, log=print):
    """用作業系統的檔案總管打開資料夾。"""
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        raise PipelineError(f"資料夾還不存在：{path}")
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                                 # noqa: S606
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception as e:                                     # pragma: no cover
        raise PipelineError(f"打不開資料夾：{e}")
    log(f"已開啟資料夾：{path}")


# ------------------------------------------------------------------ 設定
DEFAULTS = {
    "course_name": "我的課程", "teacher": "", "ta_name": "",
    "course_code": "EC",              # 課程縮寫（2 個英文大寫字母），學生編號用
    "semester": "115-1",              # 學期（例 115-1），學生編號用
    "semester_start": "", "logo_path": "", "stopwords_extra": [],
    "keep_short": [], "user_words": [], "deck_series": "",
    "synonyms": {},                   # {正式概念: [同義詞, ...]}
    "question_types": {},             # {類別名: [關鍵詞, ...]}，覆寫預設兩類
    "appendix_matrix": True,          # 要不要做附錄「概念矩陣」頁
    "top_concepts": 3,                # 每題抓幾個重點概念
    "min_words_for_cloud": 5, "input_dir": "", "output_dir": "",
}


def load_config(path=None):
    """讀設定：優先讀使用者設定檔，其次讀打包進來的 config.json 範本。"""
    cfg = dict(DEFAULTS)
    for p in [resource_path("config.json"), path or user_config_path()]:
        if p and os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
                    cfg.update({k: v for k, v in json.load(f).items() if k in DEFAULTS})
            except Exception:
                pass
    # 範本裡的相對路徑（input/output）在桌面版沒有意義，交給 GUI 自己決定
    for k in ("input_dir", "output_dir"):
        if cfg.get(k) and not os.path.isabs(cfg[k]):
            cfg[k] = ""
    return cfg


def user_config_path():
    return os.path.join(app_home(), "settings.json")


def save_config(cfg):
    data = {k: cfg.get(k, v) for k, v in DEFAULTS.items()}
    with open(user_config_path(), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return user_config_path()


def norm_course_code(s):
    from core.mask import normalize_course_code
    return normalize_course_code(s)


def norm_semester(s):
    from core.mask import normalize_semester
    return normalize_semester(s)


def logo_file(cfg):
    p = (cfg.get("logo_path") or "").strip()
    if p and os.path.isabs(p) and os.path.exists(p):
        return p
    bundled = resource_path("assets", "logo_placeholder.png")
    return bundled if os.path.exists(bundled) else ""


def sample_input_dir():
    return resource_path("sample_input")


# ------------------------------------------------------------------ jieba
_JIEBA_READY = False


def init_jieba(log=print):
    """打包後 jieba 找不到自己的 dict.txt，必須明確指定路徑。"""
    global _JIEBA_READY
    if _JIEBA_READY:
        return
    import logging
    import jieba
    jieba.setLogLevel(logging.WARN)
    d = resource_path("jieba", "dict.txt")
    if os.path.exists(d):
        try:
            jieba.set_dictionary(d)
        except Exception as e:
            log(f"  [提醒] 指定 jieba 字典失敗，改用預設字典：{e}")
    jieba.dt.tmp_dir = tempfile.gettempdir()
    _JIEBA_READY = True


# ------------------------------------------------------------------ 週次
def week_start(today=None):
    """回傳 today 所在那一週的星期日（週期＝週日～週六）。"""
    today = today or dt.date.today()
    return today - dt.timedelta(days=(today.weekday() + 1) % 7)


def parse_semester_start(s):
    s = (s or "").strip().replace("/", "-")
    if not s:
        raise PipelineError("「學期起日」是空的，無法自動判週。\n"
                            "請填第 1 週的星期日（例如 2026-09-06），或改用「指定週次」。")
    try:
        d = dt.date.fromisoformat(s)
    except ValueError:
        raise PipelineError(f"學期起日格式看不懂：{s}\n請用 YYYY-MM-DD，例如 2026-09-06。")
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)      # 校正成星期日


def week_number(semester_start, d=None):
    s = parse_semester_start(semester_start)
    return (week_start(d) - s).days // 7 + 1


def week_folder(n, semester_start, d=None):
    """Wn-MMDD-MMDD（週日～週六）。"""
    if (semester_start or "").strip():
        sun = parse_semester_start(semester_start) + dt.timedelta(weeks=(n - 1))
    else:
        sun = week_start(d)
    sat = sun + dt.timedelta(days=6)
    return f"W{n}-{sun:%m%d}-{sat:%m%d}", sun, sat


# ------------------------------------------------------------------ 主流程
def _core():
    from core import parse_zuvio as PZ
    from core import analyze as AN
    from core import wordcloud_gen as WG
    from core import deck as DK
    return PZ, AN, WG, DK


def can_make_thumbs():
    """縮圖靠 PowerPoint COM：只有 Windows 而且裝了 PowerPoint 才做得到。"""
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client                                # noqa: F401
    except Exception:
        return False
    return True


def run_pipeline(cfg, input_dir, out_dir, week_label, deck_name,
                 thumbs=True, log=print):
    """完整跑一次：解析 → 統計 → 文字雲 → 簡報 → QA。回傳結果摘要 dict。"""
    PZ, AN, WG, DK = _core()
    input_dir = os.path.abspath(input_dir)
    out_dir = os.path.abspath(out_dir)
    if not os.path.isdir(input_dir):
        raise PipelineError(f"找不到輸入資料夾：{input_dir}")
    if not PZ.list_inputs(input_dir):
        raise PipelineError(
            f"輸入資料夾裡沒有任何 .xlsx 或 .csv：\n  {input_dir}\n\n"
            "請先到 Zuvio 老師端「下載數據」，把 xlsx 放進這個資料夾再試一次。")

    os.makedirs(out_dir, exist_ok=True)
    init_jieba(log)
    AN.load_dict(cfg.get("user_words"), cfg.get("stopwords_extra"), cfg.get("keep_short"),
                 synonyms=cfg.get("synonyms"), question_types=cfg.get("question_types"))

    from core.mask import normalize_course_code, normalize_semester
    course_code = normalize_course_code(cfg.get("course_code"))
    semester = normalize_semester(cfg.get("semester"))

    log("")
    log("[1/5] 解析匯出檔（學生編號優先）並完成去識別化")
    log(f"  學生編號格式：{semester}_{course_code}_n"
        f"（輸入檔若已有 A 欄「學生編號」就直接沿用）")
    index = PZ.parse_dir(input_dir, out_dir, semester=semester, course=course_code, log=log)
    with open(os.path.join(out_dir, "questions_index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

    log("")
    log("[2/5] 斷詞、三個重點／概念矩陣、提問分類")
    result = AN.analyse_dir(out_dir, index=index,
                            top_concepts=int(cfg.get("top_concepts", 3) or 3), log=log)
    AN.write_outputs(result, out_dir, cfg.get("course_name", ""), week_label)

    log("")
    log("[3/5] 產生文字雲")
    wc_map = WG.generate_all(out_dir, min_words=int(cfg.get("min_words_for_cloud", 5) or 5),
                             log=log)

    log("")
    log("[4/5] 組裝簡報")
    meta = {
        "course_name": cfg.get("course_name", ""),
        "teacher": cfg.get("teacher", ""),
        "ta_name": cfg.get("ta_name", ""),
        "week_label": week_label,
        "date": dt.date.today().strftime("%Y/%m/%d"),
        "series": cfg.get("deck_series") or f"{cfg.get('course_name', '')}　同學作答分析",
        "logo": logo_file(cfg),
        "appendix_matrix": bool(cfg.get("appendix_matrix", True)),
    }
    pptx_path = os.path.join(out_dir, deck_name)
    prs, total = DK.build_deck(result, pptx_path, meta, wc_map, log=log)

    log("")
    log("[5/5] 版面 QA")
    qa_dir = os.path.join(out_dir, "qa")
    rep = DK.layout_scan(prs, os.path.join(qa_dir, "layout_report.json"))
    log(f"  幾何掃描：出界 {len(rep['out_of_bounds'])} 處、"
        f"重疊>8% {len(rep['overlaps'])} 處、刻意出血裝飾 {len(rep['刻意出血裝飾'])} 個")
    for o in rep["out_of_bounds"][:6]:
        log(f"    出界：{o}")
    for o in rep["overlaps"][:6]:
        log(f"    重疊：{o}")

    if thumbs:
        if can_make_thumbs():
            DK.export_thumbs(pptx_path, qa_dir, log=log)
        else:
            log("  [略過] 這台電腦沒有 PowerPoint（或非 Windows），不產生縮圖，不影響簡報。")

    with open(os.path.join(out_dir, "run_meta.json"), "w", encoding="utf-8") as f:
        json.dump({
            "產生時間": dt.datetime.now().isoformat(timespec="seconds"),
            "工具": f"{APP_TITLE} {VERSION}",
            "課程": cfg.get("course_name", ""), "週次標籤": week_label,
            # 只記資料夾名稱，不記完整本機路徑（輸出可能會分享出去，
            # 而且路徑裡的數字會被隱私掃描誤判）
            "輸入資料夾": os.path.basename(os.path.normpath(input_dir)),
            "輸出資料夾": os.path.basename(os.path.normpath(out_dir)),
            "簡報": deck_name, "頁數": total, "題數": len(result["題目"]),
            "學生編號格式": f"{semester}_{course_code}_n",
            "去識別化": "姓名第 2 字改 O；學號與 email 全部改 O；主鍵用學生編號",
            "幾何QA": {"出界": len(rep["out_of_bounds"]), "重疊": len(rep["overlaps"])},
        }, f, ensure_ascii=False, indent=2)

    log("")
    log("=" * 60)
    log(f"完成！輸出資料夾：{out_dir}")
    log(f"  簡報：{deck_name}（{total} 頁，每頁備忘稿都有逐字稿）")
    log("  摘要：summary.md　資料：analysis.json、Q*.csv")
    log("  每題兩份分析表：Q0N_concept_matrix.csv（概念矩陣）、Q0N_questions.csv（提問分類）")
    log("=" * 60)
    return {"out_dir": out_dir, "pptx": pptx_path, "pages": total,
            "questions": len(result["題目"]), "wordclouds": len(wc_map),
            "analysis": result}


def run_demo(cfg, out_root, thumbs=True, log=print):
    """用內建合成資料（虛構姓名）跑一次示範。"""
    out_dir = os.path.join(os.path.abspath(out_root), "示範")
    if os.path.isdir(out_dir):
        for fn in os.listdir(out_dir):
            p = os.path.join(out_dir, fn)
            try:
                shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
            except OSError:
                pass
    log("=" * 60)
    log("示範模式：使用內建合成資料（虛構姓名，不是真實學生）")
    log("=" * 60)
    demo_cfg = dict(cfg)
    demo_cfg.setdefault("course_name", "我的課程")
    return run_pipeline(demo_cfg, sample_input_dir(), out_dir,
                        "示範週", "同學作答分析-示範.pptx", thumbs=thumbs, log=log)


def pick_input_dir(input_dir, wfolder):
    """輸入資料夾底下若有 Wn-MMDD-MMDD 子資料夾就優先用它。"""
    base = os.path.abspath(input_dir)
    sub = os.path.join(base, wfolder)
    try:
        from core import parse_zuvio as PZ
        if os.path.isdir(sub) and PZ.list_inputs(sub):
            return sub
    except Exception:
        pass
    return base


def run_week(cfg, input_dir, out_root, week_no=None, thumbs=True, log=print):
    """跑某一週（week_no=None 代表依學期起日自動判斷本週）。"""
    from core import deck as DK
    n = int(week_no) if week_no else week_number(cfg.get("semester_start"))
    if n < 1:
        raise PipelineError(f"算出來的週次是第 {n} 週（今天早於學期起日）。\n"
                            "請確認「學期起日」，或改用「指定週次」。")
    wfolder, sun, sat = week_folder(n, cfg.get("semester_start"))
    week_label = f"第{DK.cn_week(n)}週（{sun:%Y/%m/%d}–{sat:%m/%d}）"
    deck_name = f"同學作答分析-第{DK.cn_week(n)}周.pptx"
    in_dir = pick_input_dir(input_dir, wfolder)
    out_dir = os.path.join(os.path.abspath(out_root), wfolder)

    log("=" * 60)
    log(f"課程：{cfg.get('course_name', '')}　週次：{week_label}")
    log(f"輸入：{in_dir}")
    log(f"輸出：{out_dir}")
    log("=" * 60)
    r = run_pipeline(cfg, in_dir, out_dir, week_label, deck_name, thumbs=thumbs, log=log)
    r["week"] = n
    return r


# ------------------------------------------------------------------ 隱私掃描
def pptx_texts(path):
    from pptx import Presentation
    out = []
    prs = Presentation(path)
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                out.append((f"slide{i}", sh.text_frame.text))
        if s.has_notes_slide:
            t = s.notes_slide.notes_text_frame.text
            if t.strip():
                out.append((f"slide{i}/notes", t))
    return out


def check_privacy(target, input_dir=None, log=print, require_roster=False):
    """掃描輸出資料夾，看有沒有漏掉的個資。回傳命中筆數（0 才算通過）。

    2.0 版比對三件事：
      1. 名冊上的真實姓名（從原始輸入檔重建，只留在記憶體）
      2. **9 碼學號**（`(?<!\\d)\\d{9}(?!\\d)`）
      3. **電子郵件樣式**
    學生編號（例 115-1_EC_3）是刻意要出現的，不算命中。
    """
    from core import parse_zuvio as PZ
    from core.mask import SID9_IN_TEXT, EMAIL, STUDENT_CODE_IN_TEXT
    target = os.path.abspath(target)
    if not os.path.isdir(target):
        raise PipelineError(f"找不到要掃描的資料夾：{target}")

    if not input_dir:
        meta_p = os.path.join(target, "run_meta.json")
        if os.path.exists(meta_p):
            try:
                with open(meta_p, encoding="utf-8") as f:
                    cand = json.load(f).get("input_dir") or ""
                if cand and os.path.isdir(cand):
                    input_dir = cand
            except Exception:
                pass
    reals = []
    if input_dir and os.path.isdir(input_dir):
        reals = [r for r, _ in PZ.roster_from_dir(input_dir)]
        log(f"名冊比對來源：{input_dir}（{len(reals)} 位真實姓名）")
        if not reals:
            log("  （輸入檔的姓名已經是遮罩版，名冊自然是空的 —— 這是正確的。）")
    elif require_roster:
        raise PipelineError(
            "找不到原始輸入資料夾，無法比對名冊。\n"
            "請先在上方把「輸入資料夾」指到當初那批 xlsx，再按一次「隱私檢查」。\n"
            "（提示：也可以直接選當週的輸出資料夾，例如 W3-0920-0926）")
    else:
        log("名冊比對來源：無（只做 9 碼學號與電子郵件樣式掃描）")
    log(f"掃描目標：{target}")
    log("允許出現：學生編號（例 115-1_EC_3）")

    hits = []

    def scan(where, txt):
        for name in reals:
            if name and name in txt:
                hits.append({"檔案": where, "類型": "名冊姓名", "命中": name})
        # 學生編號本身不含 9 碼連續數字，先挖掉以策安全
        clean_txt = STUDENT_CODE_IN_TEXT.sub(" ", txt)
        for m in SID9_IN_TEXT.finditer(clean_txt):
            hits.append({"檔案": where, "類型": "9 碼學號", "命中": m.group(0)})
        for m in EMAIL.finditer(clean_txt):
            hits.append({"檔案": where, "類型": "電子郵件", "命中": m.group(0)})

    for root, _dirs, files in os.walk(target):
        for fn in files:
            p = os.path.join(root, fn)
            ext = os.path.splitext(fn)[1].lower()
            rel = os.path.relpath(p, target)
            chunks = []
            if ext in TEXT_EXT:
                try:
                    with open(p, encoding="utf-8-sig", errors="replace") as f:
                        chunks = [(rel, f.read())]
                except Exception as e:
                    log(f"  [略過] 讀不到 {rel}：{e}")
            elif ext == ".pptx":
                try:
                    chunks = [(f"{rel}#{w}", t) for w, t in pptx_texts(p)]
                except Exception as e:
                    log(f"  [略過] 讀不到 {rel}：{e}")
            for where, txt in chunks:
                scan(where, txt)

    log("")
    log("=" * 60)
    if hits:
        kinds = {}
        for h in hits:
            kinds[h["類型"]] = kinds.get(h["類型"], 0) + 1
        log(f"[不通過] 發現 {len(hits)} 處可能的個資（"
            + "、".join(f"{k} {v} 處" for k, v in kinds.items()) + "）：")
        for h in hits[:40]:
            log(f"  - {h['檔案']}　{h['類型']}：「{h['命中']}」")
        log("請先不要把輸出交出去，並回報這個問題。")
    else:
        log(f"[通過] 0 命中。比對 {len(reals)} 個名冊姓名、9 碼學號與電子郵件樣式，"
            "輸出資料夾內都沒有出現。")
    log("=" * 60)
    return len(hits)


# ------------------------------------------------------------------ 自我測試
def _link_sheet_leak(input_dir, out_dir, log=print):
    """驗證「學生編號連結姓名」工作表的姓名／學號／email 完全沒有進到輸出。"""
    import openpyxl
    from core.parse_zuvio import LINK_SHEET_NAMES
    secrets = set()
    for fn in os.listdir(input_dir):
        if not fn.lower().endswith(".xlsx"):
            continue
        try:
            wb = openpyxl.load_workbook(os.path.join(input_dir, fn), read_only=True,
                                        data_only=True)
        except Exception:
            continue
        for nm in wb.sheetnames:
            if nm.strip() not in LINK_SHEET_NAMES:
                continue
            for row in wb[nm].iter_rows(min_row=2, values_only=True):
                for c in (row[1:4] if row else []):
                    c = str(c or "").strip()
                    if len(c) >= 2 and c not in ("姓名", "學號（原）", "電子郵件（原）"):
                        secrets.add(c)
        wb.close()
    if not secrets:
        log("  連結表檢查：輸入檔沒有「學生編號連結姓名」工作表（略過）")
        return 0
    n = 0
    for root, _d, files in os.walk(out_dir):
        for fn in files:
            p = os.path.join(root, fn)
            ext = os.path.splitext(fn)[1].lower()
            texts = []
            if ext in TEXT_EXT:
                try:
                    with open(p, encoding="utf-8-sig", errors="replace") as f:
                        texts = [f.read()]
                except Exception:
                    pass
            elif ext == ".pptx":
                try:
                    texts = [t for _w, t in pptx_texts(p)]
                except Exception:
                    pass
            for t in texts:
                for sec in secrets:
                    if sec in t:
                        n += 1
                        log(f"    [外洩] {os.path.relpath(p, out_dir)} 出現「{sec}」")
    log(f"  連結表檢查：{len(secrets)} 個機敏字串，輸出命中 {n} 處")
    return n


def _selftest_questions(log):
    """提問偵測與兩類分類的規則驗證（含陳述句干擾與各型問句）。"""
    from core import analyze as AN
    cases = [
        # (句子, 是不是提問, 應該分到哪一類)
        ("了解物質是由原子組成，以及原子如何組成不同物質", False, None),
        ("理解物質的微觀結構如何決定性質", False, None),
        ("我知道要怎麼做實驗了", False, None),
        ("原子核帶正電，電子帶負電。", False, None),
        ("原子是什麼", False, None),                  # 沒問號、疑問詞不在句首
        ("好想知道更多", False, None),
        ("有什麼不同？", True, "概念理解類"),
        ("Why?", True, "概念理解類"),
        ("Why important?", True, "概念理解類"),
        ("為什麼原子會結合成分子？", True, "概念理解類"),
        ("什麼是原子", True, "概念理解類"),           # 句首疑問詞
        ("這樣對嗎", True, "概念理解類"),             # 句尾「嗎」
        ("下列何者正確？", True, "概念理解類"),
        ("要怎麼看到？", True, "操作應用類"),
        ("這個公式要怎麼算？", True, "操作應用類"),
        ("那要用什麼儀器？", True, "操作應用類"),
        ("How do we measure it?", True, "操作應用類"),
        ("我想知道為什麼會這樣", True, "概念理解類"),
    ]
    problems = []
    n_q = n_other = 0
    for s, want_q, want_c in cases:
        got_q = AN.is_question_sentence(s)
        if got_q != want_q:
            problems.append(f"提問偵測錯誤（應為 {want_q}）：{s}")
            continue
        if not got_q:
            continue
        n_q += 1
        got_c = AN.classify_question(s)[0]
        if got_c == AN.OTHER_TYPE:
            n_other += 1
        if got_c != want_c:
            problems.append(f"分類錯誤（應為 {want_c}，得到 {got_c}）：{s}")
    if n_q and n_other * 100.0 / n_q > 20:
        problems.append(f"「其他」佔 {n_other * 100.0 / n_q:.0f}%，超過兩成")
    log(f"  提問規則：{len(cases)} 個案例，問句 {n_q} 個，其他 {n_other} 個，"
        f"錯誤 {len(problems)} 個")
    return problems


def _selftest_compat(tmp, cfg, log):
    """向下相容檢查：直接餵一份「原始 Zuvio 匯出」格式（學號是 int、姓名沒遮罩），
    確認程式會自己補學生編號、把學號改成 O、把自由文字裡的同學姓名換成編號。"""
    from core import parse_zuvio as PZ
    from core.mask import SID9_IN_TEXT
    import openpyxl

    names = ["甲小明", "乙小華", "丙大文"]
    sids = [411354041, 411354042, 411354043]        # 刻意用 int，複刻真實匯出
    ans = ["原子經濟性讓我很有感，反應物要盡量都變成產物。為什麼一定要這樣算？",
           "我和甲小明討論過，實驗室要先算好用量再倒試劑。",
           "廢棄物減量最重要，源頭減量比事後回收有效。這個公式怎麼算？"]

    path = os.path.join(tmp, "compat_raw.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["資料夾名稱", "相容測試"])
    ws.append(["問題題型", "問答"])
    ws.append(["題幹", "向下相容測試題"])
    ws.append([])
    ws.append(["問題類型", "", "", "第1題:問答題"])
    ws.append(["問題敘述", "", "", "向下相容測試題"])
    ws.append([])
    ws.append(["學號", "姓名", "作答時間", "回答"])
    for i, n in enumerate(names):
        ws.append([sids[i], n, "2026-09-08 19:00:00", ans[i]])
    wb.save(path)

    rows, rec = PZ.parse_file(path, 1, semester="115-1", course="GC")
    problems = []
    if rec.get("學生編號來源", "").startswith("A 欄"):
        problems.append("相容模式誤判成已有學生編號欄")
    codes = {r[3] for r in rows}
    if len(codes) != 3 or not all(c.startswith("115-1_GC_") for c in codes):
        problems.append(f"相容模式的學生編號不對：{sorted(codes)}")
    for r in rows:
        if r[4] != "OOOOOOOOO":
            problems.append(f"相容模式沒有把學號改成全 O：{r[4]}")
            break
    blob = "\n".join(r[7] for r in rows) + "\n" + "\n".join(r[5] for r in rows)
    for n in names:
        if n in blob:
            problems.append(f"相容模式沒有換掉自由文字／姓名欄的「{n}」")
    if SID9_IN_TEXT.search(blob):
        problems.append("相容模式的輸出仍有 9 碼學號")
    if "115-1_GC_1" not in blob:
        problems.append("相容模式沒有把自由文字裡的同學姓名換成學生編號")
    log(f"  向下相容（原始 Zuvio 匯出、學號 int）：學生編號 {sorted(codes)}、"
        f"學號欄 {rows[0][4]}、姓名欄 {rows[0][5]}")
    return problems


def selftest(log=print, keep=False):
    """用內建合成資料（NameMasker 2.0 格式）跑完整流程並驗收。

    驗收項目：
      1. 簡報頁數 = 3（封面／總覽／結尾）＋ 題數（＋附錄頁）
      2. 每題都有 Q0N_concept_matrix.csv 與 Q0N_questions.csv
      3. 每個覆蓋率都介於 0–1
      4. 幾何 QA：出界 0、重疊 0
      5. 隱私檢查 0 命中（名冊姓名＋9 碼學號＋email 樣式）
      6. 文字雲 ≥ 3 張
      7. 每一頁都有備忘稿逐字稿
    """
    from pptx import Presentation
    tmp = tempfile.mkdtemp(prefix="hwwc_selftest_")
    try:
        log(f"暫存資料夾：{tmp}")
        cfg = load_config()
        cfg["course_name"] = cfg.get("course_name") or "我的課程"
        r = run_demo(cfg, tmp, thumbs=False, log=log)

        out_dir = r["out_dir"]
        pptx = r["pptx"]
        result = r.get("analysis") or {}
        qs = result.get("題目", [])
        n_q = len(qs)
        problems = []

        # ---- 1. 頁數
        want_appendix = bool(cfg.get("appendix_matrix", True)) and \
            any(q.get("重點概念") for q in qs)
        expect = 3 + n_q + (1 if want_appendix else 0)
        if not os.path.exists(pptx):
            problems.append("沒有產生 pptx")
            pages = 0
            prs = None
        else:
            prs = Presentation(pptx)
            pages = len(prs.slides)
        if pages != expect:
            problems.append(f"簡報 {pages} 頁，應該是 {expect} 頁"
                            f"（3 + 題數 {n_q}{' + 附錄 1' if want_appendix else ''}）")

        # ---- 2. 每題兩個 CSV
        for q in qs:
            for fn in (f"{q['題號']}_concept_matrix.csv", f"{q['題號']}_questions.csv"):
                if not os.path.exists(os.path.join(out_dir, fn)):
                    problems.append(f"少了 {fn}")

        # ---- 3. 覆蓋率介於 0–1
        for q in qs:
            vals = [q.get("整體覆蓋率", 0), q.get("三個都提到比例", 0)] + \
                   [f["覆蓋率"] for f in q.get("重點概念", [])]
            for v in vals:
                if not (0.0 <= float(v) <= 1.0):
                    problems.append(f"{q['題號']} 覆蓋率 {v} 不在 0–1 之間")
        n_focus = sum(1 for q in qs if len(q.get("重點概念", [])) == 3)
        if n_focus < 1:
            problems.append("沒有任何一題抓到 3 個重點概念")

        # ---- 4. 幾何 QA
        rep_p = os.path.join(out_dir, "qa", "layout_report.json")
        oob = ovl = -1
        if os.path.exists(rep_p):
            with open(rep_p, encoding="utf-8") as f:
                rep = json.load(f)
            oob, ovl = len(rep["out_of_bounds"]), len(rep["overlaps"])
            if oob:
                problems.append(f"版面出界 {oob} 處")
                for o in rep["out_of_bounds"][:5]:
                    log(f"    出界：{o}")
            if ovl:
                problems.append(f"版面重疊 {ovl} 處")
                for o in rep["overlaps"][:5]:
                    log(f"    重疊：{o}")
        else:
            problems.append("沒有產生 layout_report.json")

        # ---- 5. 文字雲
        wc_dir = os.path.join(out_dir, "wc")
        pngs = [f for f in os.listdir(wc_dir)] if os.path.isdir(wc_dir) else []
        pngs = [f for f in pngs if f.lower().endswith(".png")]
        if len(pngs) < 3:
            problems.append(f"文字雲只有 {len(pngs)} 張，少於 3 張")

        # ---- 6. 每頁備忘稿（一定要有，而且不能超過 350 字）
        if prs is not None:
            empty, toolong = [], []
            for i, s in enumerate(prs.slides, 1):
                t = (s.notes_slide.notes_text_frame.text
                     if s.has_notes_slide else "")
                if not t.strip():
                    empty.append(i)
                elif len(t) > 350:
                    toolong.append((i, len(t)))
            if empty:
                problems.append(f"第 {empty} 頁沒有備忘稿逐字稿")
            if toolong:
                problems.append(f"逐字稿超過 350 字：{toolong}")

        # ---- 6b. 提問偵測與分類規則
        log("")
        log("提問規則檢查")
        problems += _selftest_questions(log)

        # ---- 7. 隱私
        hits = check_privacy(out_dir, log=log)
        if hits:
            problems.append(f"隱私檢查命中 {hits} 處")

        # ---- 7b.「學生編號連結姓名」工作表絕不可被讀進來
        leaked = _link_sheet_leak(sample_input_dir(), out_dir, log=log)
        if leaked:
            problems.append(f"「學生編號連結姓名」的內容外洩 {leaked} 處")

        # ---- 8. 向下相容（原始 Zuvio 匯出格式）
        log("")
        log("向下相容檢查")
        try:
            problems += _selftest_compat(tmp, cfg, log)
        except Exception as e:
            problems.append(f"向下相容檢查失敗：{e}")

        log("")
        log(f"檢查結果：簡報 {pages}/{expect} 頁、題數 {n_q}、每題重點 3 個的有 {n_focus} 題、"
            f"文字雲 {len(pngs)} 張、出界 {oob} 處、重疊 {ovl} 處、隱私命中 {hits} 處")
        if problems:
            for p in problems:
                log(f"  [失敗] {p}")
            log("SELFTEST FAILED")
            return 1
        log("SELFTEST OK")
        return 0
    finally:
        if not keep:
            shutil.rmtree(tmp, ignore_errors=True)
