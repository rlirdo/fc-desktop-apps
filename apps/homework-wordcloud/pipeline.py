# -*- coding: utf-8 -*-
"""
pipeline.py — 「學生作業文字雲」流程核心（可被 GUI 或 CLI 呼叫）。

原始版本是命令列工具 ta_wordcloud.py；這裡把流程改寫成純函式：
  * 所有訊息都走 log callback（GUI 可以即時顯示在文字框），不直接 print。
  * 出錯丟 PipelineError，不呼叫 sys.exit（GUI 不能被關掉）。
  * 所有資料檔（jieba 字典、sample_input、logo、config.json）一律用 resource_path()
    定位，PyInstaller onefile 打包後會落在 sys._MEIPASS。

隱私鐵律（不可關閉）：
  * 輸出的姓名一律遮罩（第 2 字改 O），自由文字裡的名冊姓名也一併替換。
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
VERSION = "1.0.0"

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
    "semester_start": "", "logo_path": "", "stopwords_extra": [],
    "keep_short": [], "user_words": [], "deck_series": "",
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
    AN.load_dict(cfg.get("user_words"), cfg.get("stopwords_extra"), cfg.get("keep_short"))

    log("")
    log("[1/5] 解析匯出檔並遮罩姓名")
    index = PZ.parse_dir(input_dir, out_dir, log=log)
    with open(os.path.join(out_dir, "questions_index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

    log("")
    log("[2/5] 斷詞與統計")
    result = AN.analyse_dir(out_dir, index=index, log=log)
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
            "input_dir": input_dir, "output_dir": out_dir,
            "簡報": deck_name, "頁數": total, "題數": len(result["題目"]),
            "姓名遮罩": "已套用（第 2 字改 O）",
            "幾何QA": {"出界": len(rep["out_of_bounds"]), "重疊": len(rep["overlaps"])},
        }, f, ensure_ascii=False, indent=2)

    log("")
    log("=" * 60)
    log(f"完成！輸出資料夾：{out_dir}")
    log(f"  簡報：{deck_name}（{total} 頁，每頁備忘稿都有逐字稿）")
    log("  摘要：summary.md　資料：analysis.json、Q*.csv（姓名已遮罩）")
    log("=" * 60)
    return {"out_dir": out_dir, "pptx": pptx_path, "pages": total,
            "questions": len(result["題目"]), "wordclouds": len(wc_map)}


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


def check_privacy(target, input_dir=None, log=print):
    """掃描輸出資料夾，看有沒有漏遮罩的真實姓名。回傳命中筆數（0 才算通過）。"""
    from core import parse_zuvio as PZ
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
    if not input_dir or not os.path.isdir(input_dir):
        raise PipelineError(
            "找不到原始輸入資料夾，無法比對名冊。\n"
            "請先在上方把「輸入資料夾」指到當初那批原始 xlsx，再按一次「隱私檢查」。\n"
            "（提示：也可以直接選當週的輸出資料夾，例如 W3-0920-0926）")

    reals = [r for r, _ in PZ.roster_from_dir(input_dir)]
    log(f"名冊比對來源：{input_dir}（{len(reals)} 位）")
    log(f"掃描目標：{target}")

    hits = []
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
                for name in reals:
                    if name and name in txt:
                        hits.append({"檔案": where, "命中姓名": name})

    log("")
    log("=" * 60)
    if hits:
        log(f"[不通過] 發現 {len(hits)} 處未遮罩的真實姓名：")
        for h in hits[:40]:
            log(f"  - {h['檔案']}　命中「{h['命中姓名']}」")
        log("請先不要把輸出交出去，並回報這個問題。")
    else:
        log(f"[通過] 0 命中。比對 {len(reals)} 個名冊姓名，輸出資料夾內都沒有出現。")
    log("=" * 60)
    return len(hits)


# ------------------------------------------------------------------ 自我測試
def selftest(log=print, keep=False):
    """用內建合成資料跑完整流程並驗收：pptx ≥4 頁、文字雲 ≥3 張、隱私 0 命中。"""
    from pptx import Presentation
    tmp = tempfile.mkdtemp(prefix="hwwc_selftest_")
    try:
        log(f"暫存資料夾：{tmp}")
        cfg = load_config()
        cfg["course_name"] = cfg.get("course_name") or "我的課程"
        r = run_demo(cfg, tmp, thumbs=False, log=log)

        out_dir = r["out_dir"]
        pptx = r["pptx"]
        problems = []

        if not os.path.exists(pptx):
            problems.append("沒有產生 pptx")
            pages = 0
        else:
            pages = len(Presentation(pptx).slides)
        if pages < 4:
            problems.append(f"簡報只有 {pages} 頁，少於 4 頁")

        wc_dir = os.path.join(out_dir, "wc")
        pngs = [f for f in os.listdir(wc_dir)] if os.path.isdir(wc_dir) else []
        pngs = [f for f in pngs if f.lower().endswith(".png")]
        if len(pngs) < 3:
            problems.append(f"文字雲只有 {len(pngs)} 張，少於 3 張")

        hits = check_privacy(out_dir, log=log)
        if hits:
            problems.append(f"隱私檢查命中 {hits} 處")

        log("")
        log(f"檢查結果：簡報 {pages} 頁、文字雲 {len(pngs)} 張、隱私命中 {hits} 處")
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
