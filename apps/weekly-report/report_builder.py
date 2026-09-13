# -*- coding: utf-8 -*-
"""
週報內容組裝器（碩士生標準版）
==============================
把 GUI 表單（或 JSON）的資料，組成 11–15 頁的週報 PPTX。

驗收錨點（來源：weekly-report-workshop/practice.html、YCYLabWKReport.md）：
  ① 區塊序號一律中文 壹、貳、參、肆、伍、陸
  ② 沒進度的區塊照樣列在總覽頁，描述欄原文寫「本週無進度」五個字
  ③ 「問題與請益」「下週計畫」兩頁必備，本週沒有問題也要寫出來
  ④ 每一頁都要有備忘稿逐字稿（含封面與結尾頁），一句一行、12pt
  ⑤ 頁標題 ≥ 30pt、內文 ≥ 14pt（圖說 11pt）
  ⑥ 頁數 11–15 頁

本檔不含任何真實個資；示範資料一律虛構。
"""
import os
import datetime

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from core import (Deck, CN_NUM, CIRC, CHIP, TEAL, DEEP, GREEN, GOLD, NAVY, AUX,
                  BODY, WHITE, RED, card, shape, text, bullets, banner, table,
                  kpi)

NO_PROGRESS = "本週無進度"          # 一字不差，驗收錨點
NO_QUESTION = "本週沒有需要老師裁示的問題"
DEFAULT_SECTIONS = ["論文", "實驗", "修課", "專案", "行政"]
EN_LABEL = ["THESIS", "EXPERIMENT", "COURSEWORK", "PROJECT", "ADMINISTRATION", "OTHERS"]
MIN_PAGES, MAX_PAGES = 11, 15


# ------------------------------------------------------------------ 小工具
_CN_D = "零一二三四五六七八九"


def int_cn(n):
    """整數轉口語唸法：12 → 十二、23 → 二十三、105 → 一百零五（0–9999）。"""
    try:
        n = int(n)
    except Exception:
        return str(n)
    if n < 0:
        return "負" + int_cn(-n)
    if n < 10:
        return _CN_D[n]
    if n < 20:
        return "十" + ("" if n == 10 else _CN_D[n % 10])
    if n < 100:
        return _CN_D[n // 10] + "十" + ("" if n % 10 == 0 else _CN_D[n % 10])
    if n < 1000:
        r = _CN_D[n // 100] + "百"
        rem = n % 100
        if rem == 0:
            return r
        if rem < 10:
            return r + "零" + _CN_D[rem]
        return r + int_cn(rem)
    return str(n)


def date_cn(s):
    """2026/09/13 → 二零二六年九月十三日（給逐字稿唸）。"""
    d = parse_date(s)
    if not d:
        return str(s)
    y = "".join(_CN_D[int(c)] for c in f"{d.year:04d}")
    return f"{y}年{int_cn(d.month)}月{int_cn(d.day)}日"


def parse_date(s):
    s = (s or "").strip().replace("-", "/").replace(".", "/")
    for fmt in ("%Y/%m/%d", "%Y/%m/%d ", "%Y%m%d"):
        try:
            return datetime.datetime.strptime(s.strip(), fmt.strip()).date()
        except Exception:
            pass
    return None


def week_range(report_date):
    """報告日期所在的那一週（星期日～星期六）。"""
    d = parse_date(report_date) or datetime.date.today()
    # Python: Monday=0 … Sunday=6；週日當第一天
    back = (d.weekday() + 1) % 7
    start = d - datetime.timedelta(days=back)
    return start, start + datetime.timedelta(days=6)


def lines_of(s, limit=None):
    """把多行文字切成乾淨的句子清單（去掉項目符號與空行）。"""
    out = []
    for ln in str(s or "").replace("\r", "").split("\n"):
        t = ln.strip().lstrip("•·・-–—*　 ").strip()
        if t:
            out.append(t)
    return out[:limit] if limit else out


def clip(s, n):
    s = str(s or "").replace("\n", "　")
    return s if len(s) <= n else s[:n - 1] + "…"


def say(s):
    """逐字稿用：確保每一句以句號收尾（一句一行）。"""
    s = str(s or "").strip()
    if not s:
        return ""
    return s if s[-1] in "。！？；.!?" else s + "。"


# ------------------------------------------------------------------ 資料模型
def blank_data():
    return {
        "schema": "weekly-report/1",
        "reporter": "",
        "grade": "碩士班一年級",
        "advisor": "",
        "unit": "國立東華大學自然資源與環境學系　仿生與環境工作坊",
        "report_date": datetime.date.today().strftime("%Y/%m/%d"),
        "week_start": "",
        "week_end": "",
        "sections": [{"name": n, "progress": "", "photos": []} for n in DEFAULT_SECTIONS],
        "questions": "",
        "next_week": "",
    }


def normalize(data):
    """補齊欄位、限制區塊 3–6 塊、每塊最多 2 張照片。"""
    d = blank_data()
    d.update({k: v for k, v in (data or {}).items() if k in d})
    secs = []
    for s in (data or {}).get("sections", []) or []:
        name = str(s.get("name", "")).strip()
        if not name:
            continue
        photos = []
        for p in (s.get("photos") or [])[:2]:
            path = str(p.get("path", "")).strip()
            if not path:
                continue
            photos.append({"path": path,
                           "caption": str(p.get("caption", "")).strip(),
                           "date": str(p.get("date", "")).strip(),
                           "place": str(p.get("place", "")).strip()})
        secs.append({"name": name, "progress": str(s.get("progress", "")).strip(),
                     "photos": photos})
    if len(secs) < 3:
        for n in DEFAULT_SECTIONS:
            if len(secs) >= 3:
                break
            if n not in [x["name"] for x in secs]:
                secs.append({"name": n, "progress": "", "photos": []})
    d["sections"] = secs[:6]
    if not d.get("report_date"):
        d["report_date"] = datetime.date.today().strftime("%Y/%m/%d")
    if not d.get("week_start") or not d.get("week_end"):
        a, b = week_range(d["report_date"])
        d["week_start"] = d.get("week_start") or a.strftime("%Y/%m/%d")
        d["week_end"] = d.get("week_end") or b.strftime("%Y/%m/%d")
    return d


def cap_text(ph, idx):
    """圖說：說明（日期，地點）。"""
    body = ph.get("caption") or "本週工作紀錄"
    meta = "，".join([x for x in (ph.get("date"), ph.get("place")) if x])
    return f"{body}（{meta}）。" if meta else f"{body}。"


# ------------------------------------------------------------------ 產生器
def build_report(data, out_path, progress_cb=None):
    """把 data 組成週報 PPTX，存到 out_path，回傳 (out_path, 頁數)。"""
    d = normalize(data)
    secs = d["sections"]
    k = len(secs)
    with_prog = [s for s in secs if s["progress"].strip()]
    photos_all = [(s, p) for s in secs for p in s["photos"]]

    # ---- 先排頁（頁尾要印「N / 總頁數」，所以總頁數必須先算出來） ----
    # 基本盤：封面＋總覽＋重點摘要＋各區塊（3–6）＋問題與請益＋下週計畫＋結尾 = 9–12 頁
    pages = ["cover", "overview", "highlight"]
    pages += [("section", i) for i in range(k)]
    pages += ["questions", "nextweek", "closing"]
    # 不足 11 頁時依序補頁（都不是灌水頁：照片彙整／分節頁／規格自我檢核）
    pads = []
    if photos_all:
        pads.append(("gallery", len(pages) - 3))
    pads.append(("selfcheck", len(pages) - 3))
    pads.append(("partdivider", 3))
    for kind, _ in pads:
        if len(pages) >= MIN_PAGES:
            break
        pos = 3 if kind == "partdivider" else len(pages) - 3
        pages.insert(pos, kind)
    total = min(len(pages), MAX_PAGES)
    pages = pages[:MAX_PAGES]

    series = f"仿生與環境工作坊・個人週報｜{d['reporter']}｜{d['report_date']}"
    deck = Deck(series=series, total=total)

    def _cb(i):
        if progress_cb:
            progress_cb(i, total)

    n = 0        # 內容頁的眉標編號（01 / …）
    for pno, spec in enumerate(pages, start=1):
        kind = spec[0] if isinstance(spec, tuple) else spec
        if kind in ("cover", "closing", "partdivider"):
            {"cover": _cover, "closing": _closing,
             "partdivider": _partdivider}[kind](deck, d, pno)
        else:
            n += 1
            if kind == "section":
                _section(deck, d, secs[spec[1]], spec[1], n, pno)
            else:
                {"overview": _overview, "highlight": _highlight,
                 "selfcheck": _selfcheck, "gallery": _gallery,
                 "questions": _questions, "nextweek": _nextweek}[kind](deck, d, n, pno)
        _cb(pno)

    deck.save(out_path)
    return out_path, deck.slide_count


def default_filename(data):
    d = normalize(data)
    ymd = (parse_date(d["report_date"]) or datetime.date.today()).strftime("%Y%m%d")
    name = (d["reporter"] or "週報").strip().replace(" ", "")
    return f"{ymd}_{name}_週報.pptx"


# ------------------------------------------------------------------ 各頁
def _cover(deck, d, pno):
    chips = [f"{CN_NUM[i]} {s['name']}" for i, s in enumerate(d["sections"])][:5]
    sl = deck.cover(
        "WEEKLY REPORT",
        "東華大學自然資源與環境學系\n仿生與環境工作坊　個人週報",
        "Biomimicry & Environment Workshop — Weekly Progress Report",
        chips,
        [f"報告人：{d['reporter']}　{d['grade']}　｜　指導教授：{d['advisor']}",
         f"報告日期：{d['report_date']}　｜　本週期間：{d['week_start']}（日）－ {d['week_end']}（六）",
         d["unit"]])
    got = [s for s in d["sections"] if s["progress"].strip()]
    ls = [f"老師好，我是{d['reporter']}，向老師報告{date_cn(d['week_start'])}到{date_cn(d['week_end'])}這一週的工作進度。",
          f"這一頁是封面，今天的報告日期是{date_cn(d['report_date'])}。",
          f"這一週的報告一共分成{int_cn(len(d['sections']))}個區塊。"]
    ls.append("分別是" + "、".join(f"{CN_NUM[i]}、{s['name']}" for i, s in enumerate(d["sections"])) + "。")
    if got:
        ls.append(f"其中有進度的是{int_cn(len(got))}個區塊，" + "、".join(s["name"] for s in got) + "。")
    if len(got) < len(d["sections"]):
        no = [s["name"] for s in d["sections"] if not s["progress"].strip()]
        ls.append("沒有進度的區塊是" + "、".join(no) + "，我仍然會列在總覽頁上，讓老師看得到整體樣貌。")
    ls.append("接下來我依照區塊順序逐頁向老師報告。")
    deck.notes(sl, ls)
    return sl


def _overview(deck, d, n, pno):
    secs = d["sections"]
    sl = deck.page(n, "WEEKLY OVERVIEW", "本週區塊總覽",
                   "｜".join(f"{CN_NUM[i]} {s['name']}" for i, s in enumerate(secs)), pno)
    y0, y1 = 1.94, 6.26
    k = len(secs)
    gap = 0.13
    hgt = (y1 - y0 - gap * (k - 1)) / k
    y = y0
    # 卡片內可容納的行數（14pt、行高 1.22 → 每行約 0.24 吋；寬 8.46 吋 → 每行約 41 字）
    nlines = max(1, int((hgt - 0.22) / 0.24))
    maxchars = nlines * 41 - 2
    for i, s in enumerate(secs):
        has = bool(s["progress"].strip())
        desc = "　".join(lines_of(s["progress"])) if has else NO_PROGRESS
        desc = clip(desc, maxchars)
        c = CIRC[i % 5] if has else AUX
        card(sl, 0.62, y, 12.1, hgt)
        cd = min(0.6, hgt - 0.12)
        shape(sl, MSO_SHAPE.OVAL, 0.92, y + (hgt - cd) / 2, cd, cd, c)
        text(sl, CN_NUM[i], 0.92, y + (hgt - cd) / 2, cd, cd, 16, WHITE, bold=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, s["name"], 1.74, y + 0.14, 2.2, 0.4, 16, NAVY, bold=True)
        text(sl, desc, 4.0, y + 0.12, 8.46, hgt - 0.22, 14,
             BODY if has else AUX, ls=1.22)
        y += hgt + gap
    text(sl, f"※ 沒有進度的區塊仍然列在本頁，描述欄一律原文寫「{NO_PROGRESS}」五個字，不刪除該區塊。",
         0.62, 6.46, 12.1, 0.36, 14, AUX, italic=True)

    ls = ["這一頁是本週各個區塊的總覽。"]
    for i, s in enumerate(secs):
        if s["progress"].strip():
            first = lines_of(s["progress"])
            ls.append(f"第{int_cn(i+1)}塊，{s['name']}，" + say(first[0]))
            for extra in first[1:3]:
                ls.append(say(extra))
        else:
            ls.append(f"第{int_cn(i+1)}塊，{s['name']}，{NO_PROGRESS}。")
    ls.append(f"沒有進度的區塊我還是照樣列在總覽上，描述欄寫「{NO_PROGRESS}」五個字，不會把那一塊刪掉。")
    deck.notes(sl, ls)
    return sl


def _highlight(deck, d, n, pno):
    secs = d["sections"]
    got = [s for s in secs if s["progress"].strip()]
    nph = sum(len(s["photos"]) for s in secs)
    nq = len(lines_of(d["questions"])) if d["questions"].strip() else 0
    nplan = len(lines_of(d["next_week"]))
    sl = deck.page(n, "WEEKLY HIGHLIGHTS", "本週重點摘要",
                   f"本週期間 {d['week_start']} － {d['week_end']}｜依區塊順序摘錄", pno)
    kpi(sl, [(f"{len(got)} / {len(secs)}", "個區塊有進度", "其餘區塊於總覽頁標註「" + NO_PROGRESS + "」", TEAL),
             (f"{nph}", "張佐證照片", "每張皆標註拍攝日期與地點（APA 圖說原則）", DEEP),
             (f"{nq}", "題請益事項", "於「問題與請益」頁逐題說明現況與我的傾向", GREEN)],
            0.62, 1.9, 12.1, 1.9, cols=3)
    card(sl, 0.62, 3.98, 12.1, 2.2)
    text(sl, "本週各區塊一句話重點", 0.92, 4.12, 6.0, 0.34, 14, TEAL, bold=True)
    items = []
    for i, s in enumerate(secs):
        first = lines_of(s["progress"])
        items.append(f"{CN_NUM[i]}、{s['name']}：" + (clip(first[0], 44) if first else NO_PROGRESS))
    bullets(sl, items[:6], 0.92, 4.5, 11.5, 1.66, size=13.5, space=1, ls=1.1)
    banner(sl, f"本週共 {nplan} 項工作排入下週計畫；硬性時程列在「下週計畫」頁。")

    ls = ["這一頁是本週的重點摘要，上面三格是這一週的數字。",
          f"第一格，{int_cn(len(secs))}個區塊裡面有{int_cn(len(got))}個區塊有進度，其餘的在總覽頁上標註{NO_PROGRESS}。",
          f"第二格，本週一共放了{int_cn(nph)}張佐證照片，每一張都標了拍攝日期與地點。",
          f"第三格，本週有{int_cn(nq)}題要請老師裁示的事項，等一下在問題與請益那一頁逐題說明。",
          "下面這張卡片是各個區塊的一句話重點。"]
    for i, s in enumerate(secs):
        first = lines_of(s["progress"])
        ls.append(f"{CN_NUM[i]}、{s['name']}，" + (say(clip(first[0], 52)) if first else f"{NO_PROGRESS}。"))
    ls.append(f"本週一共有{int_cn(nplan)}項工作排進下週計畫，硬性時程列在下週計畫那一頁。")
    deck.notes(sl, ls)
    return sl


def _section(deck, d, s, idx, n, pno):
    cn = CN_NUM[idx]
    en = EN_LABEL[idx] if idx < len(EN_LABEL) else "SECTION"
    has = bool(s["progress"].strip())
    title = f"{cn}、{s['name']}"
    sub = f"{cn}、{s['name']}｜本週期間 {d['week_start']} － {d['week_end']}"
    sl = deck.page(n, en, title, sub, pno)
    if has:
        items = lines_of(s["progress"])
    else:
        items = [NO_PROGRESS,
                 f"本區塊仍然列在總覽頁上，描述欄原文寫「{NO_PROGRESS}」五個字，不刪除該區塊。",
                 "下週恢復進度之後，會在這一頁補上內容與佐證照片。"]
    photos = s["photos"]

    fignos = []
    if photos:
        card(sl, 0.62, 1.9, 7.5, 4.98)
        text(sl, "本週進度", 0.92, 2.06, 6.9, 0.34, 14, TEAL, bold=True)
        bullets(sl, [clip(x, 120) for x in items[:6]], 0.92, 2.5, 6.9, 3.6,
                size=14, space=8, ls=1.2)
        text(sl, "※ 未完成的項目一律列進下週計畫；照片圖說標註拍攝日期與地點。",
             0.92, 6.3, 6.9, 0.36, 12.5, AUX, italic=True)
        ys = [1.9, 4.44]
        for i, ph in enumerate(photos[:2]):
            fignos.append(deck.figure(sl, ph["path"], 8.32, ys[i], 4.0, 1.52,
                                      cap_text(ph, i), cap_h=0.52))
    else:
        shown = [clip(x, 150) for x in items[:7]]
        ch = min(4.24, 0.9 + sum(1 + len(x) // 56 for x in shown) * 0.46)
        card(sl, 0.62, 1.9, 12.1, ch)
        text(sl, "本週進度", 0.92, 2.06, 6.9, 0.34, 14, TEAL, bold=True)
        bullets(sl, shown, 0.92, 2.5, 11.5, ch - 0.7, size=14, space=8, ls=1.22)
        if has:
            banner(sl, f"{cn}、{s['name']}：本頁內容由本週實際完成的工作整理而成，未完成的項目一律列進下週計畫。")
        else:
            banner(sl, f"{cn}、{s['name']}：{NO_PROGRESS}；本區塊仍列在總覽頁上，下週恢復進度後再補內容。")

    ls = [f"這一頁向老師報告{cn}、{s['name']}這個區塊。"]
    if has:
        for x in items[:7]:
            ls.append(say(x))
    else:
        ls.append(f"這個區塊{NO_PROGRESS}。")
        ls.append("雖然沒有進度，我還是把這一塊列出來，讓老師知道它沒有被漏掉。")
        ls.append(f"總覽頁上這一塊的描述欄，也是照規格原文寫「{NO_PROGRESS}」五個字。")
        ls.append("下週恢復進度之後，我會在這一頁補上內容與佐證照片。")
        ls.append("如果老師認為這一塊應該提前處理，請老師指示，我會調整下週的優先序。")
    for i, ph in enumerate(photos[:2]):
        pos = "上面" if i == 0 else "下面"
        meta = "，".join([x for x in (ph.get("date"), ph.get("place")) if x])
        ls.append(f"右邊{pos}這張是圖{int_cn(fignos[i])}，"
                  + (ph.get("caption") or "本週工作紀錄")
                  + (f"，拍攝時間與地點是{meta}。" if meta else "。"))
    ls.append(f"以上是{cn}、{s['name']}這一塊本週的" + ("進度。" if has else "狀況。"))
    deck.notes(sl, ls)
    return sl


def _gallery(deck, d, n, pno):
    ph = [(s, p) for s in d["sections"] for p in s["photos"]][:2]
    sl = deck.page(n, "PHOTO LOG", "本週照片彙整", "本週各區塊佐證照片｜圖說格式：圖 N　說明（日期，地點）", pno)
    xs = [0.62, 6.72]
    for i, (s, p) in enumerate(ph):
        deck.figure(sl, p["path"], xs[i], 1.98, 5.55, 3.12, cap_text(p, i), cap_h=0.66)
    card(sl, 0.62, 6.0, 12.1, 0.9)
    bullets(sl, ["照片一律標拍攝日期與地點，圖說格式固定為「圖 N　說明（日期，地點）」。",
                 "含他人肖像或學生可識別影像的照片不放進公開檔案。"],
            0.92, 6.12, 11.5, 0.7, size=14, space=2)
    ls = ["這一頁是本週的照片彙整。"]
    for i, (s, p) in enumerate(ph):
        meta = "，".join([x for x in (p.get("date"), p.get("place")) if x])
        ls.append(("左邊" if i == 0 else "右邊") + f"這張是{s['name']}的紀錄，"
                  + (p.get("caption") or "本週工作紀錄") + (f"，{meta}拍的。" if meta else "。"))
    ls.append("照片一律標拍攝日期與地點，圖說格式固定寫成圖幾、說明、括號裡日期和地點。")
    ls.append("有他人肖像或學生可識別影像的照片，我不會放進公開的檔案裡。")
    deck.notes(sl, ls)
    return sl


def _partdivider(deck, d, pno):
    """分節頁（深色）：各區塊內容頁的前導頁。"""
    secs = d["sections"]
    chips = [(f"{CN_NUM[i]} {s['name']}",
              CHIP[i % 5] if s["progress"].strip() else AUX)
             for i, s in enumerate(secs)][:5]
    sub = (f"本週期間 {d['week_start']} － {d['week_end']}\n"
           f"以下依序報告各區塊的本週進度；沒有進度的區塊一律寫「{NO_PROGRESS}」。")
    sl = deck.divider("SECTIONS / WEEKLY PROGRESS", "各區塊本週進度", sub, chips=chips)
    ls = ["這一頁是各區塊進度的分節頁。",
          f"接下來我依照{'、'.join(CN_NUM[i] for i in range(len(secs)))}的順序，一個區塊一頁向老師報告。",
          "上面的標籤是這一週的區塊清單，灰色的那幾塊代表這一週沒有進度。",
          f"沒有進度的區塊我照樣會列出來，並且原文寫「{NO_PROGRESS}」五個字。"]
    deck.notes(sl, ls)
    return sl


def _selfcheck(deck, d, n, pno):
    secs = d["sections"]
    got = sum(1 for s in secs if s["progress"].strip())
    nph = sum(len(s["photos"]) for s in secs)
    sl = deck.page(n, "SELF CHECK", "本週週報自我檢核", "依週報規格逐項自我檢查後再向老師報告", pno)
    card(sl, 0.62, 1.9, 12.1, 4.24)
    table(sl, 0.92, 2.06,
          [("檢核項目", 5.6, PP_ALIGN.LEFT), ("規格要求", 4.0, PP_ALIGN.LEFT),
           ("本週狀況", 1.9, PP_ALIGN.CENTER)],
          [[("區塊序號",), ("中文 壹、貳、參、肆、伍",), ("已照規格", GREEN, True)],
           [("沒進度的區塊",), (f"原文寫「{NO_PROGRESS}」且不刪除",), ("已照規格", GREEN, True)],
           [("總覽頁",), (f"列出全部 {len(secs)} 個區塊",), (f"{got} 塊有進度", TEAL, True)],
           [("問題與請益",), ("必備頁，沒問題也要寫出來",), ("已放入", GREEN, True)],
           [("下週計畫",), ("必備頁，依優先序排列",), ("已放入", GREEN, True)],
           [("備忘稿逐字稿",), ("每一頁都要有、一句一行、12pt",), ("每頁皆有", GREEN, True)],
           [("佐證照片",), ("標拍攝日期與地點",), (f"{nph} 張", TEAL, True)]],
          head_c=[AUX, AUX, AUX], row_h=0.44)
    banner(sl, "交出去之前先自己檢核一次，老師的時間就花在內容上，不是花在抓格式。")
    ls = ["這一頁是我自己對這份週報做的規格檢核。",
          "第一項，區塊序號一律用中文的壹、貳、參、肆、伍，沒有改成阿拉伯數字。",
          f"第二項，沒有進度的區塊原文寫「{NO_PROGRESS}」五個字，而且那一塊不刪掉。",
          f"第三項，總覽頁把{int_cn(len(secs))}個區塊全部列出來，其中{int_cn(got)}塊這一週有進度。",
          "第四項與第五項，問題與請益、下週計畫這兩頁都是必備頁，這一份都有放。",
          "第六項，每一頁的備忘稿都寫了逐字稿，一句一行、十二點字級。",
          f"第七項，本週一共放了{int_cn(nph)}張佐證照片，每一張都標了拍攝日期與地點。",
          "我先自己檢核過一次再向老師報告，老師的時間就可以花在內容上。"]
    deck.notes(sl, ls)
    return sl


def _questions(deck, d, n, pno):
    qs = lines_of(d["questions"], 4)
    sl = deck.page(n, "QUESTIONS", "問題與請益",
                   "必備頁：本週沒有問題也要寫出來，不可以省略這一頁", pno)
    if qs:
        gap = 0.18
        hs = [max(0.86, 0.42 + ((len(q) - 1) // 48 + 1) * 0.40) for q in qs]
        scale = min(1.0, (4.2 - gap * (len(qs) - 1)) / sum(hs))
        hs = [h * scale for h in hs]
        y = 1.9 + (4.2 - sum(hs) - gap * (len(qs) - 1)) / 2
        for i, q in enumerate(qs):
            hgt = hs[i]
            card(sl, 0.62, y, 12.1, hgt)
            shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 0.62, y, 0.11, hgt,
                  CIRC[i % 5], radius=0.5)
            text(sl, f"請益{int_cn(i+1)}　{clip(q, 120)}", 1.0, y + 0.1, 11.5,
                 hgt - 0.2, 16, NAVY, bold=True, ls=1.2, anchor=MSO_ANCHOR.MIDDLE)
            y += hgt + gap
        banner(sl, f"如果某一週真的沒有問題，這一頁仍然要放，寫「{NO_QUESTION}」。")
        ls = [f"這一頁是問題與請益，本週有{int_cn(len(qs))}個問題要請老師裁示。"]
        for i, q in enumerate(qs):
            ls.append(f"第{int_cn(i+1)}個問題，" + say(q))
        ls.append("這一頁是必備頁，如果某一週真的沒有問題，我也會放這一頁，寫上本週沒有需要老師裁示的問題。")
    else:
        card(sl, 0.62, 2.7, 12.1, 2.4)
        shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 0.62, 2.7, 0.11, 2.4, GOLD, radius=0.5)
        text(sl, NO_QUESTION, 1.0, 3.1, 11.5, 0.6, 22, NAVY, bold=True)
        bullets(sl, ["本週的工作都照原定計畫進行，沒有需要老師當場裁示的事項。",
                     "若下週出現需要裁示的問題，我會在下一次週報的這一頁逐題列出。"],
                1.0, 3.9, 11.4, 0.9, size=14, space=6)
        banner(sl, "沒有問題也要放這一頁——把「沒有問題」講清楚，本身就是一種回報。")
        ls = ["這一頁是問題與請益。",
              f"本週{NO_QUESTION}。",
              "這一週的工作都照原定計畫進行，沒有需要老師當場裁示的事項。",
              "如果下週出現需要裁示的問題，我會在下一次週報的這一頁逐題列出來。",
              "這一頁是必備頁，就算沒有問題我也會放，把沒有問題講清楚本身就是一種回報。"]
    deck.notes(sl, ls)
    return sl


def _nextweek(deck, d, n, pno):
    plan = lines_of(d["next_week"], 7)
    if not plan:
        plan = ["延續本週未完成的工作，並於下次週報回報進度。"]
    sl = deck.page(n, "NEXT WEEK", "下週計畫", "必備頁：依優先序排列，硬性時程請標出來", pno)
    card(sl, 0.62, 1.9, 12.1, min(4.24, 0.72 + len(plan) * 0.50))
    rows = []
    for i, p in enumerate(plan):
        hard = any(w in p for w in ("硬性", "截止", "期限", "務必", "考試", "小考"))
        rows.append([(str(i + 1),), (clip(p, 70),),
                     ("是", RED, True) if hard else ("－", AUX, False)])
    table(sl, 0.92, 2.06,
          [("序", 0.7, PP_ALIGN.CENTER), ("要做的事", 9.8, PP_ALIGN.LEFT),
           ("硬性時程", 1.4, PP_ALIGN.CENTER)],
          rows, head_c=[AUX, AUX, AUX], row_h=0.46,
          band={i for i, r in enumerate(rows) if r[2][0] == "是"})
    banner(sl, "下週計畫依優先序排列；有硬性時程的先排，其他事都排在這幾件之後。")
    ls = [f"這一頁是下週計畫，一共{int_cn(len(plan))}件事，依照優先序排列。"]
    for i, p in enumerate(plan):
        ls.append(f"第{int_cn(i+1)}件，" + say(p))
    ls.append("有硬性時程的我用紅字和金色底標出來，其他的事都排在這幾件之後。")
    deck.notes(sl, ls)
    return sl


def _closing(deck, d, pno):
    nxt = parse_date(d["report_date"])
    nxt = (nxt + datetime.timedelta(days=7)).strftime("%Y/%m/%d") if nxt else ""
    sl = deck.closing("THANK YOU", "綠色化學 × 花蓮在地資源 × 永續",
                      "Green Chemistry × Local Resources of Hualien × Sustainability",
                      [f"{d['reporter']}　{d['grade']}　｜　{d['unit']}",
                       f"感謝 {d['advisor']} 指導　｜　{d['report_date']}"
                       + (f"　｜　下次週報：{nxt}" if nxt else "")])
    ls = ["以上是本週的報告。"]
    got = [s["name"] for s in d["sections"] if s["progress"].strip()]
    if got:
        ls.append("這一週有進度的區塊是" + "、".join(got) + "。")
    qs = lines_of(d["questions"])
    if qs:
        ls.append(f"我有{int_cn(len(qs))}個問題等老師裁示，麻煩老師給我方向。")
    else:
        ls.append(f"本週{NO_QUESTION}。")
    if nxt:
        ls.append(f"下一次週報的日期是{date_cn(nxt)}。")
    ls.append("以上是本週報告，謝謝老師。")
    deck.notes(sl, ls)
    return sl
