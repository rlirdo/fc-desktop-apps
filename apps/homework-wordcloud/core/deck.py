# -*- coding: utf-8 -*-
"""
deck.py — 產生「同學作答分析-第N周.pptx」（自然科學雜誌風，16:9）

頁型：
    封面 → 總覽（各題人數／作答率表）→ 每題一頁（左文字雲、右 TOP5 關鍵詞卡）→ 結尾

版式規則：
    封面／結尾＝navy 深色頁；內容頁＝#F5F9FA 淺底
    內容頁：teal 眉標「0N / ENGLISH LABEL」→ 30pt 深色粗體標題 → 斜體灰副標
    每題頁：左白色圓角卡放文字雲＋圖說，右白色圓角卡放 TOP5 關鍵詞，底部 navy 色帶一句總結
    頁尾：左下系列名、右下「N / 總頁數」
    字級：頁標題 30pt、內文 18pt、圖說 14pt（老花眼可讀）

備忘稿：
    每一頁都寫逐字稿，一句一行（句號後換行），第一人稱口語，數字用口語唸法。
"""
import os
import json

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ------------------------------------------------------------------ 色票／字型
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
DEEP = RGBColor(0x06, 0x5A, 0x82)
TEAL = RGBColor(0x1C, 0x72, 0x93)
GREEN = RGBColor(0x2E, 0x7D, 0x5B)
MOSS = RGBColor(0x5F, 0xA9, 0x8A)
GOLD = RGBColor(0xC9, 0x9A, 0x3E)
BG = RGBColor(0xF5, 0xF9, 0xFA)
TEXT = RGBColor(0x1E, 0x29, 0x3B)
MUTED = RGBColor(0x64, 0x74, 0x8B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD8, 0xE3, 0xE8)

FONT = "微軟正黑體"
CYCLE = [TEAL, DEEP, GREEN, GOLD, NAVY]
CYCLE_ON_DARK = [TEAL, DEEP, GREEN, GOLD, MOSS]

SW, SH = 13.333, 7.5          # inches（16:9）
CM = 1 / 2.54
LOGO_CONTENT = 0.9 * CM
LOGO_COVER = 2.6 * CM

CN_DIGITS = "〇一二三四五六七八九"
CN_WEEK = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
           "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"]


# ------------------------------------------------------------------ 數字口語化
def cn_num(n):
    """把整數轉成中文口語唸法（0–9999 夠用）。四十二 / 一百零五 / 二十。"""
    n = int(n)
    if n < 0:
        return "負" + cn_num(-n)
    if n < 10:
        return CN_DIGITS[n]
    if n < 20:
        return "十" + (CN_DIGITS[n % 10] if n % 10 else "")
    if n < 100:
        return CN_DIGITS[n // 10] + "十" + (CN_DIGITS[n % 10] if n % 10 else "")
    if n < 1000:
        r = n % 100
        head = CN_DIGITS[n // 100] + "百"
        if r == 0:
            return head
        if r < 10:
            return head + "零" + CN_DIGITS[r]
        return head + cn_num(r)
    r = n % 1000
    head = CN_DIGITS[n // 1000] + "千"
    if r == 0:
        return head
    if r < 100:
        return head + "零" + cn_num(r)
    return head + cn_num(r)


def cn_pct(x):
    """百分比口語：85.0 -> 百分之八十五；85.5 -> 百分之八十五點五。"""
    x = float(x)
    i = int(x)
    frac = round(x - i, 1)
    s = "百分之" + cn_num(i)
    if frac >= 0.05:
        s += "點" + CN_DIGITS[int(round(frac * 10))]
    return s


def cn_week(n):
    n = int(n)
    return CN_WEEK[n] if 0 <= n < len(CN_WEEK) else str(n)


# ------------------------------------------------------------------ 排版工具
def set_bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color


def rect(slide, x, y, w, h, fill=None, shape=MSO_SHAPE.RECTANGLE,
         line=None, radius=None, shadow=False):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    s.shadow.inherit = shadow
    if s.has_text_frame:
        s.text_frame.text = ""
    return s


def textbox(slide, x, y, w, h, text, size=18, color=TEXT, bold=False,
            italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.2, space_after=0, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = FONT
    return tb


def band(slide, text, y=6.34, size=15):
    """底部 navy 圓角橫幅，一句總結。"""
    b = rect(slide, 0.55, y, 12.23, 0.62, fill=NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
    tf = b.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = WHITE
    r.font.name = FONT
    return b


def disp_width(s):
    """中文字算 1.0、英數算 0.55 的顯示寬度（單位：字寬）。"""
    w = 0.0
    for ch in s:
        w += 0.55 if ord(ch) < 0x2E80 else 1.0
    return w


def fit_one_line(s, width_in, size_pt, reserve=0.0):
    """把字串截到能在 width_in 吋、size_pt 字級內單行放完，超出則加刪節號。

    PowerPoint 的中文字寬約等於字級；換算成吋 = size_pt / 72。
    多留 reserve 個字寬給行尾，避免剛好壓線而折行溢出卡片。
    """
    budget = width_in / (size_pt / 72.0) - reserve
    if disp_width(s) <= budget:
        return s
    out = ""
    used = 0.0
    for ch in s:
        cw = 0.55 if ord(ch) < 0x2E80 else 1.0
        if used + cw > budget - 1.0:
            break
        out += ch
        used += cw
    return out + "…"


def fit_image(slide, png, bx, by, bw, bh):
    """等比縮放置中放入方框，絕不變形。"""
    with Image.open(png) as im:
        iw, ih = im.size
    ar = iw / ih
    w, h = bw, bw / ar
    if h > bh:
        h, w = bh, bh * ar
    x = bx + (bw - w) / 2
    y = by + (bh - h) / 2
    slide.shapes.add_picture(png, Inches(x), Inches(y), Inches(w), Inches(h))
    return x, y, w, h


def notes(slide, lines):
    """逐字稿寫入備忘稿：一句一行。"""
    text = "\n".join(x for x in lines if x)
    tf = slide.notes_slide.notes_text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(12)
            r.font.name = FONT


def footer(slide, series, page, total):
    textbox(slide, 0.55, 7.02, 7.0, 0.3, series, size=10, color=MUTED)
    textbox(slide, 10.8, 7.02, 1.98, 0.3, f"{page} / {total}", size=10,
            color=MUTED, align=PP_ALIGN.RIGHT)


def logo_corner(slide, logo):
    if logo and os.path.exists(logo):
        slide.shapes.add_picture(logo, Inches(SW - 0.55 - LOGO_CONTENT),
                                 Inches(0.36), Inches(LOGO_CONTENT), Inches(LOGO_CONTENT))


def page_head(slide, n, label, title, subtitle, logo=None):
    textbox(slide, 0.55, 0.42, 8.0, 0.3, f"{n:02d} / {label}", size=12,
            color=TEAL, bold=True)
    textbox(slide, 0.55, 0.74, 11.5, 0.62,
            fit_one_line(title, 11.5, 30), size=30, bold=True, color=TEXT)
    textbox(slide, 0.55, 1.40, 11.5, 0.34,
            fit_one_line(subtitle, 11.5, 14), size=14, italic=True, color=MUTED)
    logo_corner(slide, logo)


# ------------------------------------------------------------------ 頁型
def slide_cover(prs, ctx):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY)
    rect(s, 9.5, -1.7, 5.6, 5.6, fill=RGBColor(0x10, 0x2C, 0x4E), shape=MSO_SHAPE.OVAL)
    rect(s, 11.2, 4.4, 3.4, 3.4, fill=RGBColor(0x0E, 0x27, 0x45), shape=MSO_SHAPE.OVAL)
    rect(s, -1.1, 5.3, 3.0, 3.0, fill=RGBColor(0x0E, 0x27, 0x45), shape=MSO_SHAPE.OVAL)

    textbox(s, 0.85, 1.30, 9.0, 0.3, ctx["眉標"], size=13, color=MOSS, bold=True)
    textbox(s, 0.85, 1.75, 9.2, 1.5, ctx["主標"], size=38, bold=True, color=WHITE,
            line_spacing=1.15)
    textbox(s, 0.85, 3.30, 9.2, 0.4, ctx["英文副標"], size=16, italic=True,
            color=RGBColor(0x9F, 0xC4, 0xD6))

    x = 0.85
    for i, k in enumerate(ctx["關鍵字"]):
        w = 0.36 + 0.20 * disp_width(k)
        if x + w > 9.6:
            break
        p = rect(s, x, 3.95, w, 0.46, fill=CYCLE_ON_DARK[i % len(CYCLE_ON_DARK)],
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        tf = p.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        pr = tf.paragraphs[0]
        pr.alignment = PP_ALIGN.CENTER
        r = pr.add_run()
        r.text = k
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = WHITE
        r.font.name = FONT
        x += w + 0.18

    rect(s, 0.85, 4.95, 6.6, 0.02, fill=RGBColor(0x2B, 0x4A, 0x6E))
    textbox(s, 0.85, 5.20, 7.9, 1.4, ctx["報告人資訊"], size=15,
            color=RGBColor(0xD6, 0xE4, 0xEC), line_spacing=1.45)

    logo = ctx.get("logo")
    if logo and os.path.exists(logo):
        lx = SW - 0.85 - LOGO_COVER
        s.shapes.add_picture(logo, Inches(lx), Inches(4.92),
                             Inches(LOGO_COVER), Inches(LOGO_COVER))
    notes(s, ctx["逐字稿"])
    return s


def slide_overview(prs, ctx, qs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, BG)
    page_head(s, ctx["頁碼標"], ctx["眉標"], ctx["標題"], ctx["副標"], ctx.get("logo"))

    rect(s, 0.55, 1.88, 12.23, 4.32, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
    cols = [(0.85, 0.95), (1.85, 4.55), (6.50, 1.55), (8.15, 1.45), (9.70, 2.85)]
    for (cx, cw), h in zip(cols, ["題號", "題目", "作答/全班", "作答率", "作答率長條"]):
        textbox(s, cx, 2.10, cw, 0.3, h, size=13, bold=True, color=DEEP)
    rect(s, 0.85, 2.45, 11.65, 0.015, fill=LINE)

    # 題數多時自動縮小列高，確保 10 題也放得下（卡片底緣 6.20 吋）
    rows_h = min(0.345, (6.10 - 2.58) / max(len(qs), 1))
    size = 14 if rows_h >= 0.30 else 12
    y = 2.58
    for i, q in enumerate(qs):
        col = CYCLE[i % len(CYCLE)]
        textbox(s, 0.85, y, 0.95, 0.3, q["題號"], size=size, bold=True, color=col)
        textbox(s, 1.85, y, 4.55, 0.3, fit_one_line(q["題目"], 4.55, size),
                size=size, color=TEXT)
        textbox(s, 6.50, y, 1.55, 0.3, f"{q['作答人數']} / {q['全班人數']}",
                size=size, color=TEXT)
        textbox(s, 8.15, y, 1.45, 0.3, f"{q['作答率']}%", size=size, bold=True, color=col)
        rect(s, 9.70, y + 0.07, 2.80, 0.17, fill=RGBColor(0xE6, 0xEE, 0xF2))
        rect(s, 9.70, y + 0.07, max(2.80 * q["作答率"] / 100.0, 0.05), 0.17, fill=col)
        y += rows_h

    band(s, fit_one_line(ctx["總結"], 11.8, 15), size=15)
    notes(s, ctx["逐字稿"])
    return s


def slide_question(prs, ctx, wc_png, idx):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, BG)
    page_head(s, ctx["頁碼標"], ctx["眉標"], ctx["標題"], ctx["副標"], ctx.get("logo"))

    # 左欄：文字雲＋圖說
    rect(s, 0.55, 1.88, 6.30, 4.32, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    iw = 5.84
    if wc_png and os.path.exists(wc_png):
        fit_image(s, wc_png, 0.78, 2.10, iw, 2.90)
        textbox(s, 0.78, 5.18, iw, 0.90, ctx["圖說"], size=14, color=MUTED,
                line_spacing=1.28)
    else:
        textbox(s, 0.78, 2.60, iw, 1.4, ctx.get("大數字", ""), size=72, bold=True,
                color=TEAL, align=PP_ALIGN.CENTER)
        textbox(s, 0.78, 4.20, iw, 0.5, ctx.get("大數字說明", ""), size=18,
                color=TEXT, align=PP_ALIGN.CENTER)
        textbox(s, 0.78, 5.00, iw, 1.0, ctx.get("圖說", ""), size=14, color=MUTED,
                align=PP_ALIGN.CENTER, line_spacing=1.3)

    # 右欄：TOP5 關鍵詞卡
    rect(s, 7.08, 1.88, 5.70, 4.32, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    textbox(s, 7.36, 2.14, 5.14, 0.32, ctx.get("重點標題", "前五大關鍵詞"),
            size=14, bold=True, color=DEEP)
    rect(s, 7.36, 2.52, 5.14, 0.015, fill=LINE)

    # 每條 18pt 單行（行高 ≈ 18 × 1.22 × 1.3 / 72 ≈ 0.40 吋），5 條 ＋ 間距塞進 3.25 吋
    tb = s.shapes.add_textbox(Inches(7.36), Inches(2.80), Inches(5.14), Inches(3.25))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    items = ctx["重點"][:5]
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.18
        p.space_after = Pt(13)
        key = it["關鍵詞"]
        body = it["說明"]
        head = f"• {key}　"
        # 關鍵詞最多 8 個字寬，其餘留給說明；整條必須單行
        head = fit_one_line(head, 5.14 * 0.45, 18)
        r1 = p.add_run()
        r1.text = head
        r1.font.size = Pt(18)
        r1.font.bold = True
        r1.font.color.rgb = CYCLE[idx % len(CYCLE)]
        r1.font.name = FONT
        rest_w = 5.14 - disp_width(head) * (18 / 72.0)
        r2 = p.add_run()
        r2.text = fit_one_line(body, rest_w, 18, reserve=0.4)
        r2.font.size = Pt(18)
        r2.font.color.rgb = TEXT
        r2.font.name = FONT

    band(s, fit_one_line(ctx["總結"], 11.8, 15), size=15)
    notes(s, ctx["逐字稿"])
    return s


def slide_closing(prs, ctx):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY)
    rect(s, 9.9, -1.5, 5.2, 5.2, fill=RGBColor(0x10, 0x2C, 0x4E), shape=MSO_SHAPE.OVAL)
    rect(s, -1.4, 4.9, 3.6, 3.6, fill=RGBColor(0x0E, 0x27, 0x45), shape=MSO_SHAPE.OVAL)
    textbox(s, 0.9, 2.20, 9.4, 0.32, ctx["眉標"], size=13, color=MOSS, bold=True)
    textbox(s, 0.9, 2.62, 10.2, 1.3, ctx["主標"], size=38, bold=True, color=WHITE,
            line_spacing=1.15)
    rect(s, 0.9, 4.20, 6.6, 0.02, fill=RGBColor(0x2B, 0x4A, 0x6E))
    textbox(s, 0.9, 4.45, 10.0, 1.4, ctx["說明"], size=18,
            color=RGBColor(0xD6, 0xE4, 0xEC), line_spacing=1.45)
    notes(s, ctx["逐字稿"])
    return s


# ------------------------------------------------------------------ 幾何 QA
def layout_scan(prs, out_json=None):
    """文字/圖不得出界、不得互疊 >8%。深色頁角落的裝飾圓是刻意出血，另列一欄。"""
    W, H = prs.slide_width, prs.slide_height
    report = {"slide_w_in": W / 914400, "slide_h_in": H / 914400,
              "slides": len(prs.slides), "out_of_bounds": [], "overlaps": [],
              "刻意出血裝飾": []}
    TOL = Emu(9144)

    for si, s in enumerate(prs.slides, 1):
        boxes = []
        for sh in s.shapes:
            if sh.left is None or sh.top is None:
                continue
            kind = "pic" if sh.shape_type == 13 else (
                "text" if (sh.has_text_frame and sh.text_frame.text.strip()) else "deco")
            boxes.append((kind, sh.left, sh.top, sh.width, sh.height,
                          (sh.text_frame.text[:18] if sh.has_text_frame else "")))
            if (sh.left < -TOL or sh.top < -TOL or
                    sh.left + sh.width > W + TOL or sh.top + sh.height > H + TOL):
                bucket = "刻意出血裝飾" if kind == "deco" else "out_of_bounds"
                report[bucket].append({
                    "slide": si, "kind": kind,
                    "text": (sh.text_frame.text[:24] if sh.has_text_frame else ""),
                    "box_in": [round(sh.left / 914400, 2), round(sh.top / 914400, 2),
                               round(sh.width / 914400, 2), round(sh.height / 914400, 2)]})

        live = [b for b in boxes if b[0] in ("text", "pic")]
        for i in range(len(live)):
            for j in range(i + 1, len(live)):
                a, b = live[i], live[j]
                ox = max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
                oy = max(0, min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2]))
                if ox <= 0 or oy <= 0:
                    continue
                amin = min(a[3] * a[4], b[3] * b[4])
                ratio = (ox * oy) / amin if amin else 0
                if ratio > 0.08:
                    report["overlaps"].append({"slide": si, "ratio": round(ratio, 3),
                                               "a": f"{a[0]}:{a[5]}", "b": f"{b[0]}:{b[5]}"})

    if out_json:
        os.makedirs(os.path.dirname(out_json), exist_ok=True)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    return report


def export_thumbs(pptx_path, out_dir, pages=None, width=1600, log=print):
    """用 PowerPoint COM 匯出縮圖（本機沒有 PowerPoint 就跳過，不算失敗）。"""
    try:
        import win32com.client as win32
    except Exception:
        log("  [略過] 未安裝 pywin32，無法用 PowerPoint 匯出縮圖。")
        return []
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    made = []
    app = pres = None
    try:
        app = win32.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(pptx_path, WithWindow=False)
        n = pres.Slides.Count
        pages = pages or list(range(1, n + 1))
        h = int(width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
        for p in pages:
            if p < 1 or p > n:
                continue
            dst = os.path.join(out_dir, f"p{p:02d}.png")
            pres.Slides(p).Export(dst, "PNG", width, h)
            made.append(dst)
        log(f"  已匯出 {len(made)} 張縮圖 -> {out_dir}")
    except Exception as e:
        log(f"  [略過] PowerPoint COM 匯出縮圖失敗：{e}")
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
    return made


# ------------------------------------------------------------------ 內容組裝
def _keyword_items(q):
    """把 TOP 詞轉成右卡 5 條「關鍵詞 — 出現 N 次，M 位同學提到」。"""
    items = []
    for row in q.get("TOP詞明細", [])[:5]:
        w, c, ppl = row["詞"], row["次數"], row["提及人數"]
        items.append({"關鍵詞": w, "說明": f"出現 {c} 次，{ppl} 位同學提到"})
    if not items:
        for w, c in q.get("TOP詞", [])[:5]:
            items.append({"關鍵詞": w, "說明": f"出現 {c} 次"})
    return items


def _q_notes(q, idx, total_q, week_label):
    """每題頁逐字稿：一句一行、第一人稱口語、數字用口語唸法。"""
    top = q.get("TOP詞明細") or [{"詞": w, "次數": c, "提及人數": 0}
                                 for w, c in q.get("TOP詞", [])]
    L = [f"這一頁是第{cn_num(idx)}題，題目是「{q['題目']}」。",
         f"這一題全班有{cn_num(q['作答人數'])}位同學作答，班級人數是{cn_num(q['全班人數'])}位，"
         f"作答率{cn_pct(q['作答率'])}。",
         f"可以做文字分析的開放作答共{cn_num(q.get('文字列數', q['有效列數']))}列，"
         f"平均每則大約{cn_num(round(q['平均字數']))}個字。",
         "左邊這張圖是同學作答的文字雲，字愈大代表出現次數愈多。"]
    if top:
        w0 = top[0]
        L.append(f"出現最多的關鍵詞是「{w0['詞']}」，一共出現{cn_num(w0['次數'])}次。")
        if w0.get("提及人數"):
            L.append(f"有{cn_num(w0['提及人數'])}位同學在答案裡提到它。")
    if len(top) >= 3:
        L.append("接下來依序是「" + "」、「".join(t["詞"] for t in top[1:4]) + "」。")
    L.append("右邊這張卡片，我把前五個關鍵詞的出現次數和提到的人數列出來。")
    for s in q.get("子題", []):
        if s.get("選項分佈"):
            o = s["選項分佈"][0]
            L.append(f"這一題還有一個選擇題子題，最多同學選的是「{o['選項']}」，"
                     f"佔{cn_pct(o['百分比'])}。")
    if q.get("常見問題"):
        L.append("這一題也有同學提出疑問，我把它整理在摘要檔裡，等一下可以在課堂上回應。")
    L.append("以上是這一題的作答情形，接下來看下一題。")
    return L


def build_deck(analysis, out_path, meta, wc_map, log=print):
    """依 analysis.json 內容組裝簡報。meta 需含 course_name/ta_name/teacher/week_label/…"""
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    qs = analysis["題目"]
    total = 3 + len(qs)          # 封面 + 總覽 + 每題一頁 + 結尾
    series = meta["series"]
    logo = meta.get("logo")
    week_label = meta.get("week_label", "")
    course = meta.get("course_name", "")

    n_q = len(qs)
    n_ans = sum(q["作答人數"] for q in qs)
    avg_rate = round(sum(q["作答率"] for q in qs) / max(n_q, 1), 1)
    top10 = analysis.get("全班重點TOP10", [])
    top_words = [w for w, _ in top10[:5]]

    # ---- 1. 封面
    slide_cover(prs, {
        "眉標": meta.get("eyebrow", "WEEKLY HOMEWORK ANALYTICS"),
        "主標": f"{course}\n同學作答分析　{week_label}",
        "英文副標": "Student Response Word Cloud & Key-Concept Report",
        "關鍵字": (top_words or ["文字雲", "作答分析"])[:4] + ["姓名已遮罩"],
        "報告人資訊": (f"授課教師：{meta.get('teacher', '')}\n"
                   f"教學助理：{meta.get('ta_name', '')}\n"
                   f"製作日期：{meta.get('date', '')}"),
        "logo": logo,
        "逐字稿": [
            "老師好，我是這門課的教學助理。",
            f"這是{course}{week_label}的同學作答分析。",
            f"這一次一共整理了{cn_num(n_q)}題，全部同學的作答都已經做過姓名遮罩。",
            "接下來我會先報告整體作答情形，再一題一題看關鍵詞。",
        ],
    })

    # ---- 2. 總覽
    slide_overview(prs, {
        "頁碼標": 1, "眉標": "OVERVIEW",
        "標題": "整體作答情形", "副標": f"{week_label}　各題作答人數與作答率",
        "總結": (f"本週{cn_num(n_q)}題平均作答率 {avg_rate}%；"
               f"全班最常出現的詞是「{top_words[0] if top_words else '—'}」"),
        "logo": logo,
        "逐字稿": [
            "這一頁是本週的整體作答情形。",
            f"表格由左到右分別是題號、題目、作答人數比全班人數、作答率，以及右邊的作答率長條。",
            f"本週一共{cn_num(n_q)}題，平均作答率是{cn_pct(avg_rate)}。",
            f"累計的作答人次是{cn_num(n_ans)}人次。",
        ] + ([f"全班最常出現的關鍵詞是「{top_words[0]}」。"] if top_words else []) + [
            "下面的深藍色橫幅是這一頁的一句話重點。",
            "接下來我逐題說明。",
        ],
    }, qs)
    footer(prs.slides[-1], series, 2, total)

    # ---- 3. 每題一頁
    for i, q in enumerate(qs, 1):
        png = wc_map.get(q["題號"])
        cap = (f"圖 {i}　{q['題號']} 作答文字雲（{week_label}，{course}）\n"
               f"資料來源：Zuvio 下載數據；字級與出現次數成正比；姓名已遮罩")
        slide_question(prs, {
            "頁碼標": i + 1, "眉標": f"QUESTION {i:02d}",
            "標題": f"{q['題號']}　{q['題目']}",
            "副標": (f"作答 {q['作答人數']} / {q['全班人數']} 人　作答率 {q['作答率']}%　"
                   f"開放文字 {q.get('文字列數', q['有效列數'])} 列　"
                   f"平均 {q['平均字數']} 字"),
            "圖說": cap,
            "重點標題": "前五大關鍵詞（依出現次數）",
            "重點": _keyword_items(q),
            "大數字": f"{q['作答人數']}",
            "大數字說明": "位同學作答",
            "總結": (f"{q['題號']} 關鍵詞：" +
                   "、".join(w for w, _ in q["TOP詞"][:5])) if q["TOP詞"] else
                   f"{q['題號']} 本題有效作答不足，建議下週提醒同學補交",
            "logo": logo,
            "逐字稿": _q_notes(q, i, n_q, week_label),
        }, png, i)
        footer(prs.slides[-1], series, i + 2, total)

    # ---- 4. 結尾
    slide_closing(prs, {
        "眉標": "NEXT STEPS",
        "主標": "以上是本週的作答分析\n謝謝老師",
        "說明": ("下一步：把關鍵詞帶回課堂回應同學的疑問。\n"
               "所有輸出檔的姓名皆已遮罩（第 2 字改 O），原始 xlsx 僅留在本機。"),
        "逐字稿": [
            "這是最後一頁。",
            f"本週一共分析了{cn_num(n_q)}題，平均作答率{cn_pct(avg_rate)}。",
            "我會把同學提出的疑問整理成清單，在下一堂課回應。",
            "所有輸出的檔案姓名都已經遮罩，原始檔只留在我的電腦裡。",
            "以上是本週報告，謝謝老師。",
        ],
    })

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    log(f"  已產生簡報：{out_path}（{total} 頁）")
    return prs, total
