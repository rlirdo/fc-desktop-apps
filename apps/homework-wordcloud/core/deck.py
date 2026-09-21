# -*- coding: utf-8 -*-
"""
deck.py — 產生「同學作答分析-第N周.pptx」（自然科學雜誌風，16:9）。（2.2）

頁型：
    封面 → 總覽（各題作答率＋整體覆蓋率）→ **一題一頁**
    → 附錄概念矩陣（每頁 2 題，最多 3 頁）→ 結尾

題頁版面（2.2 新版，開放文字題；必守）：
    左欄（0.55–5.55）白色圓角卡：文字雲（等比縮放）＋三行圖說「圖 N …」
    右上（5.78–12.78，1.86–4.46）白色圓角卡「六個重點」：
          一列一個概念｜提及 n 人｜覆蓋率 %，共 6 列；底線下一行整體覆蓋率／≥3 個比例
    下方（5.78–12.78，4.86–6.20）四張小卡（2×2）：四類提問各一張，
          第一行「類別｜提問數｜提問人數」，第二行代表句（11pt）
    底部   navy 圓角橫幅：前三大重點＋「另 N 個重點見右上表」（單行）
    頁尾   左下系列名、右下「N / 總頁數」；右上角 logo

    選擇題／測驗題沒有文字雲，整頁換一套版面（2.1 原樣不動）：
    左＝作答人數大數字＋各子題最多人選的選項；右上＝各子題選項分佈；右下＝答對率／完成率。

版式規則：
    封面／結尾＝navy 深色頁；內容頁＝#F5F9FA 淺底
    內容頁：teal 眉標「0N / ENGLISH LABEL」→ 30pt 深色粗體標題（動詞在前）→ 斜體灰副標
    字級：頁標題 30pt、卡片內文 14pt、圖說與代表句 11pt
    所有單行文字都過 fit_one_line 自動縮字；卡片高度一律由「內容行高」推算，
    行高以 PowerPoint 實測值估算：LINE_H = 字級 × 1.32 × line_spacing ＋ space_after。
    幾何 QA（layout_scan）要求出界 0、重疊 0；另用 PowerPoint COM 縮圖目視。

備忘稿：
    每一頁都寫逐字稿，一句一行（句號後換行），第一人稱口語，數字用口語唸法，
    題頁 150-350 字並且一定要講六個重點（前三個講數字、後三個合併一句）、
    三種覆蓋率、四類提問的則數與一則代表句
    （超過字數時由 _trim_notes 依優先順序拿掉可省略的句子）。
"""
import os
import re
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


PPT_LINE = 1.32         # PowerPoint 實測行高 ≈ 字級 × 1.3～1.35（微軟正黑體）


def line_h(size_pt, line_spacing=1.0, space_after=0.0):
    """一個段落實際佔多高（吋）。卡片高度一律用它推算，不可只靠幾何框目測。"""
    return (size_pt * PPT_LINE * line_spacing + space_after) / 72.0


def block_h(n, size_pt, line_spacing=1.0, space_after=0.0):
    return n * line_h(size_pt, line_spacing, space_after)


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
    # 標題寬度 11.0 吋（0.55–11.55）刻意讓開右上角 logo（11.88 起），避免幾何重疊
    textbox(slide, 0.55, 0.42, 8.0, 0.3, f"{n:02d} / {label}", size=12,
            color=TEAL, bold=True)
    textbox(slide, 0.55, 0.74, 11.0, 0.62,
            fit_one_line(title, 11.0, 30), size=30, bold=True, color=TEXT)
    textbox(slide, 0.55, 1.40, 11.0, 0.34,
            fit_one_line(subtitle, 11.0, 14), size=14, italic=True, color=MUTED)
    logo_corner(slide, logo)


def column(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT):
    """欄式文字框：lines = [(文字, 字級, 顏色, 粗體, 行距, 段後)]。

    同一張表的每一欄都用相同的行結構，行高一致 → 各欄自然對齊，
    而且一欄只佔一個 shape，幾何 QA 不會互相判定重疊。
    """
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, item in enumerate(lines):
        txt, size, color, bold, ls, sa = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = ls
        p.space_after = Pt(sa)
        r = p.add_run()
        # 空字串的段落 PowerPoint 會退回預設 18pt，行高就跟旁邊那一欄對不齊，
        # 因此空白行一律塞一個不換行空格，讓字級真的生效。
        r.text = txt if txt else " "
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = FONT
    return tb


# ------------------------------------------------------------------ 頁型
def slide_cover(prs, ctx):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY)
    rect(s, 9.5, -1.7, 5.6, 5.6, fill=RGBColor(0x10, 0x2C, 0x4E), shape=MSO_SHAPE.OVAL)
    rect(s, 11.2, 4.4, 3.4, 3.4, fill=RGBColor(0x0E, 0x27, 0x45), shape=MSO_SHAPE.OVAL)
    rect(s, -1.1, 5.3, 3.0, 3.0, fill=RGBColor(0x0E, 0x27, 0x45), shape=MSO_SHAPE.OVAL)

    textbox(s, 0.85, 1.30, 9.0, 0.3, ctx["眉標"], size=13, color=MOSS, bold=True)
    # 主標最多兩行、每行都先縮到 9.2 吋內：課程名或週次標籤太長時若讓它自動折成第三行，
    # 會直接壓到下面的英文副標（2.1 的封面就踩過這個坑）。
    head = [fit_one_line(ln, 9.2, 38) for ln in str(ctx["主標"]).split("\n")][:2]
    textbox(s, 0.85, 1.75, 9.2, 1.5, "\n".join(head), size=38, bold=True, color=WHITE,
            line_spacing=1.15, wrap=False)
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
    cols = [(0.85, 0.95), (1.85, 3.85), (5.85, 1.35), (7.30, 1.05),
            (8.45, 1.45), (10.05, 2.45)]
    for (cx, cw), h in zip(cols, ["題號", "題目", "作答/全班", "作答率",
                                  "整體覆蓋率", "覆蓋率長條"]):
        textbox(s, cx, 2.10, cw, 0.3, h, size=13, bold=True, color=DEEP)
    rect(s, 0.85, 2.45, 11.65, 0.015, fill=LINE)

    # 題數多時自動縮小列高，確保 10 題也放得下（卡片底緣 6.20 吋）
    rows_h = min(0.345, (6.10 - 2.58) / max(len(qs), 1))
    size = 14 if rows_h >= 0.30 else 12
    y = 2.58
    for i, q in enumerate(qs):
        col = CYCLE[i % len(CYCLE)]
        # 選擇題／測驗題沒有概念覆蓋率可言，欄位寫「—」而且不畫長條，
        # 免得看起來像「全班都沒抓到重點」。
        has_cov = bool(q.get("重點概念"))
        cov = round(q.get("整體覆蓋率", 0) * 100)
        textbox(s, 0.85, y, 0.95, 0.3, q["題號"], size=size, bold=True, color=col)
        textbox(s, 1.85, y, 3.85, 0.3, fit_one_line(q["題目"], 3.85, size),
                size=size, color=TEXT)
        textbox(s, 5.85, y, 1.35, 0.3, f"{q['作答人數']} / {q['全班人數']}",
                size=size, color=TEXT)
        textbox(s, 7.30, y, 1.05, 0.3, f"{q['作答率']}%", size=size, color=TEXT)
        textbox(s, 8.45, y, 1.45, 0.3, f"{cov}%" if has_cov else "—（選擇題）",
                size=size if has_cov else size - 2, bold=has_cov,
                color=col if has_cov else MUTED)
        rect(s, 10.05, y + 0.07, 2.45, 0.17, fill=RGBColor(0xE6, 0xEE, 0xF2))
        if has_cov:
            rect(s, 10.05, y + 0.07, max(2.45 * cov / 100.0, 0.05), 0.17, fill=col)
        y += rows_h

    band(s, fit_one_line(ctx["總結"], 11.8, 15), size=15)
    notes(s, ctx["逐字稿"])
    return s


# ---- 開放文字題（2.2）：右上「六個重點」卡、右下 2×2「四類提問」小卡 --------
# 幾何全部寫成常數，改版時只要動這一區；每個數字都由 line_h() 推算過。
OPEN_LEFT_X, OPEN_LEFT_W = 0.55, 5.00          # 左欄文字雲卡
OPEN_IMG = (0.75, 2.04, 4.60, 2.92)            # 圖片方框（等比縮放置中）
OPEN_CAP = (0.75, 5.04, 4.60, 1.04)            # 圖說（3 行 × 11pt）
OPEN_RX, OPEN_RW = 5.78, 7.00                  # 右欄
OPEN_PAD = 0.22                                # 卡片左右內距
CARD_TOP = 1.86
CONCEPT_ROW_SA = 2.4                           # 重點列的段後（pt）
ASK_GRID_Y, ASK_CARD_H, ASK_CARD_GAP = 4.86, 0.65, 0.04


def _concept_card(s, ctx, col):
    """右上「六個重點」卡：6 列（概念｜提及 n 人｜覆蓋率 %）＋底列覆蓋率總結。

    卡高由內容推算：內距 0.10 ＋ 標題 ＋ 分隔線 ＋ 6 列 ＋ 分隔線 ＋ 總結 ＋ 內距。
    """
    focus = (ctx.get("重點") or [])[:6]
    n_row = max(len(focus), 1)
    x, w = OPEN_RX, OPEN_RW
    ix, iw = x + OPEN_PAD, w - OPEN_PAD * 2                 # 6.00 / 6.56

    rows_h = block_h(6, 14, 1.0, CONCEPT_ROW_SA)            # 固定留 6 列的高度
    y_title = CARD_TOP + 0.10
    y_line1 = y_title + 0.30
    y_rows = y_line1 + 0.06
    y_line2 = y_rows + rows_h + 0.04
    y_sum = y_line2 + 0.04
    card_h = (y_sum + 0.22 + 0.12) - CARD_TOP

    rect(s, x, CARD_TOP, w, card_h, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    textbox(s, ix, y_title, iw, 0.26,
            fit_one_line(ctx.get("重點標題", "六個重點（依提及人數排序）"), iw, 14),
            size=14, bold=True, color=DEEP)
    rect(s, ix, y_line1, iw, 0.015, fill=LINE)

    # 三欄：概念名（左）／提及人數（右）／覆蓋率（右）
    w_name, w_ppl, w_cov = 3.60, 1.30, 1.54
    x_ppl = ix + w_name + 0.06
    x_cov = x_ppl + w_ppl + 0.06
    marks = "①②③④⑤⑥"
    c_name, c_ppl, c_cov = [], [], []
    for i, f in enumerate(focus):
        c_name.append((fit_one_line(f"{marks[i]} {f['概念']}", w_name, 14, reserve=0.4),
                       14, TEXT, True, 1.0, CONCEPT_ROW_SA))
        c_ppl.append((f"{f['提及人數']} 人", 14, TEXT, False, 1.0, CONCEPT_ROW_SA))
        c_cov.append((f"{round(f['覆蓋率'] * 100)}%", 14, col, True, 1.0, CONCEPT_ROW_SA))
    if not focus:
        c_name = [("本題沒有可統計的重點概念", 14, MUTED, False, 1.0, CONCEPT_ROW_SA)]
        c_ppl = [(" ", 14, MUTED, False, 1.0, CONCEPT_ROW_SA)]
        c_cov = [(" ", 14, MUTED, False, 1.0, CONCEPT_ROW_SA)]
    box_h = block_h(n_row, 14, 1.0, CONCEPT_ROW_SA)
    column(s, ix, y_rows, w_name, box_h, c_name)
    column(s, x_ppl, y_rows, w_ppl, box_h, c_ppl, align=PP_ALIGN.RIGHT)
    column(s, x_cov, y_rows, w_cov, box_h, c_cov, align=PP_ALIGN.RIGHT)

    rect(s, ix, y_line2, iw, 0.015, fill=LINE)
    textbox(s, ix, y_sum, iw, 0.22,
            fit_one_line(ctx.get("覆蓋率總結", ""), iw, 12, reserve=0.5),
            size=12, bold=True, color=GREEN)
    return CARD_TOP + card_h


def _ask_grid(s, ctx, y_top):
    """右下四張小卡（2×2）：四類提問各一張。

    每張小卡兩行：① 類別｜提問數｜提問人數（14pt 粗體）② 代表句（11pt 灰）。
    卡高 = 內距 0.09 ＋ 14pt 行 ＋ 11pt 行 ＋ 內距 ＝ 0.65 吋（line_h 推算後取整）。
    """
    x, w = OPEN_RX, OPEN_RW
    textbox(s, x, y_top, w, 0.24,
            fit_one_line(ctx.get("提問標題", "四類提問"), w, 12, reserve=0.5),
            size=12, bold=True, color=TEAL)

    asks = list(ctx.get("提問") or [])[:4]
    while len(asks) < 4:
        asks.append({"類別": "—", "提問數": 0, "提問人數": 0, "代表句": ""})

    cw = (w - 0.12) / 2.0                       # 3.44
    tw = cw - 0.32                              # 內距後的文字寬 3.12
    for i, a in enumerate(asks):
        cx = x + (i % 2) * (cw + 0.12)
        cy = ASK_GRID_Y + (i // 2) * (ASK_CARD_H + ASK_CARD_GAP)
        rect(s, cx, cy, cw, ASK_CARD_H, fill=WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        head = f"{a['類別']}　{a['提問數']} 則／{a['提問人數']} 人"
        rep = a.get("代表句") or ""
        rep = ("「" + rep + "」") if rep else "本週這一類沒有提問"
        textbox(s, cx + 0.16, cy + 0.09, tw, 0.26,
                fit_one_line(head, tw, 14, reserve=0.3), size=14, bold=True,
                color=CYCLE[i % len(CYCLE)])
        textbox(s, cx + 0.16, cy + 0.37, tw, 0.21,
                fit_one_line(rep, tw, 11, reserve=0.8), size=11, color=MUTED,
                wrap=False)


def _right_concept_cards(s, ctx, col):
    """開放文字題的右欄：上＝六個重點＋覆蓋率，下＝四類提問 2×2 小卡。"""
    bottom = _concept_card(s, ctx, col)
    _ask_grid(s, ctx, bottom + 0.10)


def _right_choice_cards(s, ctx, col):
    """選擇題／測驗題的右欄：上＝各子題選項分佈，下＝答對率（或作答完成率）。

    開放文字題那一套（三個重點／覆蓋率／提問）在這裡全是 0，
    放上去只是空洞的版面，所以整個換掉。
    """
    # ---- 右上：各子題選項分佈（最多 4 個子題；子題多就每題少列一個選項）
    rect(s, 7.08, 1.88, 5.70, 3.30, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    subs = ctx.get("子題選項") or []
    n_show = min(len(subs), 4)
    per = 3 if n_show <= 2 else 2
    textbox(s, 7.36, 2.06, 5.14, 0.30,
            fit_one_line(ctx.get("選項標題", "各子題選項分佈"), 5.14, 14),
            size=14, bold=True, color=DEEP)
    rect(s, 7.36, 2.42, 5.14, 0.015, fill=LINE)

    left, right = [], []
    for i, sub in enumerate(subs[:n_show]):
        left.append((fit_one_line(sub["子題"], 5.14, 11), 11, DEEP, True, 1.15, 2))
        right.append(("", 11, MUTED, False, 1.15, 2))
        for o in sub["選項"][:per]:
            left.append((fit_one_line("　" + o["選項"], 3.40, 10), 10, TEXT, False, 1.15, 3))
            right.append((f"{o['人數']} 人（{o['百分比']}%）", 10, col, True, 1.15, 3))
    if len(subs) > n_show:
        left.append((f"其餘 {len(subs) - n_show} 個子題見輸出資料夾的 CSV",
                     10, MUTED, True, 1.15, 0))
        right.append(("", 10, MUTED, False, 1.15, 0))
    if not left:
        left = [("本題沒有可統計的選項", 12, MUTED, False, 1.15, 0)]
        right = [("", 12, MUTED, False, 1.15, 0)]
    column(s, 7.36, 2.48, 3.40, 2.60, left)
    column(s, 10.86, 2.48, 1.64, 2.60, right, align=PP_ALIGN.RIGHT)

    # ---- 右下：答對率（測驗型）／作答完成率
    rect(s, 7.08, 5.32, 5.70, 0.88, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    textbox(s, 7.36, 5.44, 5.14, 0.26,
            fit_one_line(ctx.get("答對標題", "作答完成率"), 5.14, 12),
            size=12, bold=True, color=DEEP)
    textbox(s, 7.36, 5.76, 5.14, 0.30,
            fit_one_line(ctx.get("答對內容", "—"), 5.14, 13, reserve=0.5),
            size=13, bold=True, color=GREEN)


def slide_question(prs, ctx, wc_png, idx):
    """一題一頁。

    開放文字題（2.2 新版面）：左＝較窄的文字雲卡＋三行圖說；
                              右上＝六個重點卡；右下＝四類提問 2×2 小卡。
    選擇題／測驗題（2.1 版面不變）：左＝大數字＋最多人選的選項；
                                    右上＝各子題選項分佈；右下＝答對率／完成率。
    """
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, BG)
    page_head(s, ctx["頁碼標"], ctx["眉標"], ctx["標題"], ctx["副標"], ctx.get("logo"))
    col = CYCLE[idx % len(CYCLE)]
    focus = (ctx.get("重點") or [])[:6]

    if focus:
        # ---------------- 左欄（窄）：文字雲＋圖說；右欄讓出來放六列重點與四張提問卡
        rect(s, OPEN_LEFT_X, CARD_TOP, OPEN_LEFT_W, 6.20 - CARD_TOP, fill=WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        ix, iy, iw, ih = OPEN_IMG
        if wc_png and os.path.exists(wc_png):
            fit_image(s, wc_png, ix, iy, iw, ih)
        else:
            textbox(s, ix, iy + 0.90, iw, 0.40, "（本題詞數不足，未產生文字雲）",
                    size=14, color=MUTED, align=PP_ALIGN.CENTER)
        cx, cy, cw, ch = OPEN_CAP
        cap = [fit_one_line(ln, cw, 11, reserve=0.3)
               for ln in str(ctx.get("圖說", "")).split("\n")][:4]
        textbox(s, cx, cy, cw, ch, "\n".join(cap), size=11, color=MUTED,
                line_spacing=1.20, wrap=False)
        _right_concept_cards(s, ctx, col)
    else:
        # ---------------- 選擇題／測驗題：維持 2.1 版面
        rect(s, 0.55, 1.88, 6.30, 4.32, fill=WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        iw = 5.84
        textbox(s, 0.78, 2.20, iw, 1.10, ctx.get("大數字", ""), size=60, bold=True,
                color=TEAL, align=PP_ALIGN.CENTER)
        textbox(s, 0.78, 3.40, iw, 0.36, ctx.get("大數字說明", ""), size=16,
                color=TEXT, align=PP_ALIGN.CENTER)
        opts = ctx.get("選項分佈", [])[:4]
        if opts:
            column(s, 0.78, 3.88, 3.70, 1.10,
                   [(fit_one_line(o["選項"], 3.70, 13), 13, TEXT, False, 1.25, 4)
                    for o in opts])
            column(s, 4.60, 3.88, 2.02, 1.10,
                   [(f"{o['人數']} 人（{o['百分比']}%）", 13, col, True, 1.25, 4)
                    for o in opts], align=PP_ALIGN.RIGHT)
        textbox(s, 0.78, 5.18, iw, 0.90, ctx.get("圖說", ""), size=12, color=MUTED,
                align=PP_ALIGN.LEFT, line_spacing=1.30)
        _right_choice_cards(s, ctx, col)

    band(s, fit_one_line(ctx["總結"], 11.8, 15), size=15)
    notes(s, ctx["逐字稿"])
    return s


APX_MARKS = "①②③④⑤⑥"
APX_PER_PAGE = 2            # 2.2：每頁最多 2 題（原本 3 題，欄位從 3 個變 6 個）
APX_MAX_PAGES = 3           # 題數再多也只出 3 頁附錄，其餘見 CSV


def slide_appendix(prs, ctx, tables):
    """附錄：概念矩陣小表（每題 6 個概念欄 ①–⑥ × 前 10 位學生編號，每頁 2 題）。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, BG)
    page_head(s, ctx["頁碼標"], ctx["眉標"], ctx["標題"], ctx["副標"], ctx.get("logo"))
    rect(s, 0.55, 1.88, 12.23, 4.32, fill=WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)

    TW = 5.60                                   # 單一小表寬度
    LEG_SA, ROW_SA = 2.0, 1.8
    leg_h = block_h(3, 11, 1.12, LEG_SA)        # 圖例 3 行（兩個子欄各 3 行）
    row_h = block_h(11, 11, 1.05, ROW_SA)       # 表頭 1 行 ＋ 10 位學生
    y_title, y_leg = 1.98, 2.30
    y_head = y_leg + leg_h + 0.08

    for ti, t in enumerate(tables[:APX_PER_PAGE]):
        x0 = 0.85 + ti * 5.95
        colr = CYCLE[ti % len(CYCLE)]
        textbox(s, x0, y_title, TW, 0.28, fit_one_line(t["標題"], TW, 13),
                size=13, bold=True, color=colr)

        # 圖例 ①–⑥ 排成兩個子欄各 3 行，省下一半高度給矩陣本體
        concepts = list(t["概念"])[:6]
        legs = [[], []]
        for i in range(6):
            txt = (fit_one_line(f"{APX_MARKS[i]} {concepts[i]}", 2.72, 11, reserve=0.3)
                   if i < len(concepts) else "")
            legs[i // 3].append((txt, 11, MUTED, False, 1.12, LEG_SA))
        column(s, x0, y_leg, 2.72, leg_h, legs[0])
        column(s, x0 + 2.84, y_leg, 2.72, leg_h, legs[1])

        codes = t["學生編號"][:10]
        c0 = [("學生編號", 11, DEEP, True, 1.05, ROW_SA)] + \
             [(c, 11, TEXT, False, 1.05, ROW_SA) for c in codes]
        column(s, x0, y_head, 1.50, row_h, c0)
        for j in range(6):
            lines = [(APX_MARKS[j] if j < len(concepts) else "—", 11, DEEP, True,
                      1.05, ROW_SA)]
            for c in codes:
                v = t["矩陣"].get(c, [])
                on = j < len(v) and v[j]
                lines.append(("1" if on else "0", 11, (colr if on else MUTED),
                              False, 1.05, ROW_SA))
            column(s, x0 + 1.54 + j * 0.68, y_head, 0.64, row_h, lines,
                   align=PP_ALIGN.CENTER)

    y_note = y_head + row_h + 0.08
    textbox(s, 0.85, y_note, 11.60, 0.22,
            fit_one_line(ctx.get("附註", ""), 11.60, 11, reserve=0.5),
            size=11, color=MUTED)
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
NOTE_MAX = 350          # 逐字稿上限（字），約 80–90 秒
NOTE_MIN = 150


def _trim_notes(lines, optional_idx, limit=NOTE_MAX, protect=()):
    """逐字稿超過字數上限時，依優先順序把「可省略的句子」拿掉。

    lines        = [句子]
    optional_idx = [可省略句子的索引]，**排在前面的先被拿掉**
    protect      = [不可被砍掉的句子索引]。規格要求題頁一定要有「代表句」與結尾句，
                   2.1 版單純從尾端 pop，長題目一超字就把這兩句吃掉了，所以加這層保護。
    """
    protect = set(protect)
    drop = set()
    for i in optional_idx:
        keep = [s for k, s in enumerate(lines) if k not in drop]
        if len("\n".join(keep)) <= limit:
            break
        drop.add(i)
    idx = [k for k in range(len(lines)) if k not in drop]

    def total(ix):
        return len("\n".join(lines[k] for k in ix))

    # 可省的都拿掉了還是超過 → 從後面往前砍「沒有被保護」的句子
    while len(idx) > 1 and total(idx) > limit:
        cand = [k for k in idx[1:] if k not in protect]
        if not cand:
            break
        idx.remove(cand[-1])

    out = [lines[k] for k in idx]
    # 最後手段：把最長的那一句截短，保證 ≤ limit（含標點）
    for _ in range(20):
        over = len("\n".join(out)) - limit
        if over <= 0 or not out:
            break
        j = max(range(len(out)), key=lambda k: len(out[k]))
        cut = len(out[j]) - over - 1
        if cut < 8:
            break
        out[j] = out[j][:cut] + "…"
    txt = "\n".join(out)
    if len(txt) > limit:                       # 理論上到不了，留作保險
        out = [txt[:limit - 1] + "…"]
    return out


def _q_notes(q, idx, total_q, week_label):
    """每題頁逐字稿：一句一行、第一人稱口語、數字用口語唸法（150–350 字）。

    必講：題目、作答率、**六個重點**（前三個逐一講數字、後三個合併一句）、
          三種覆蓋率、**四類提問的則數**與一則代表句。
    可省：開放作答則數、文字雲說明、建議句、其他提問 —— 字數超過時依序拿掉。
    """
    focus = q.get("重點概念") or []
    asks = q.get("提問統計") or []
    L, opt, keep_idx = [], [], []

    def add(text, optional=0, must=False):
        """optional=0 必講；>0 代表可省略，數字小的先被拿掉。must=True 不可被砍。"""
        if optional:
            opt.append((optional, len(L)))
        if must:
            keep_idx.append(len(L))
        L.append(text)

    def pct(x):
        """逐字稿的百分比一律講整數，少掉「點幾」可以省下不少額度。"""
        return cn_pct(round(float(x)))

    # 題目太長會把 350 字的額度吃光，先截短（報告時老師看得到完整標題）
    title = str(q["題目"])
    title = (title[:16] + "…") if len(title) > 16 else title

    add(f"這一頁是第{cn_num(idx)}題，題目是「{title}」。")
    add(f"這一題有{cn_num(q['作答人數'])}位同學作答，全班{cn_num(q['全班人數'])}位，"
        f"作答率{pct(q['作答率'])}。")
    if focus:
        add("左邊是同學作答的文字雲，字愈大代表出現次數愈多。", optional=6)
        add(f"可以做文字分析的開放作答共{cn_num(q.get('文字列數', q['有效列數']))}則，"
            f"平均每則大約{cn_num(round(q['平均字數']))}個字。", optional=1)
        add(f"右上角是這一題的{cn_num(len(focus))}個重點。", optional=5)
        # 前三個重點逐一講數字、後三個合併一句、三種覆蓋率 —— 規格要求，一律保護
        for i, f in enumerate(focus[:3], 1):
            add(f"第{cn_num(i)}個重點是「{f['概念']}」，"
                f"{cn_num(f['提及人數'])}位同學提到，覆蓋率{pct(f['覆蓋率'] * 100)}。",
                must=True)
        rest = focus[3:6]
        if rest:
            add(f"另外{cn_num(len(rest))}個重點是"
                + "、".join(f"「{f['概念']}」" for f in rest)
                + "，分別有"
                + "、".join(cn_num(f["提及人數"]) for f in rest)
                + "位同學提到。", must=True)
        add(f"整體覆蓋率{pct(q.get('整體覆蓋率', 0) * 100)}，"
            f"提到三個以上的佔{pct(q.get('提到三個以上比例', 0) * 100)}，"
            f"{cn_num(len(focus))}個都提到的佔{pct(q.get('全部提到比例', 0) * 100)}。",
            must=True)
        if q.get("整體覆蓋率", 0) < 0.6:
            add("覆蓋率還不到六成，建議下一堂課再把這幾個重點講一次。", optional=3)
        else:
            add("覆蓋率還不錯，代表大部分同學都抓到重點了。", optional=3)
    else:
        # 選擇題／測驗題：講選項分佈與答對率，不講三個重點與提問（那些都是 0）
        add("這一題是選擇題或測驗題，回答都是固定選項，所以不做文字雲。")
        add("左邊的大數字是本題的作答人數。", optional=5)
        add("右上角這張卡片是各個子題的選項分佈。")
        for sub in (q.get("子題") or [])[:3]:
            dist = sub.get("選項分佈") or []
            if not dist:
                continue
            lab = re.sub(r":.*$", "", sub.get("子題號", "")) or "本題"
            o = dist[0]
            add(f"{lab}最多同學選的是「{o['選項']}」，"
                f"有{cn_num(o['人數'])}位，佔{cn_pct(o['百分比'])}。")
        rates = [(re.sub(r":.*$", "", s.get("子題號", "")), s["答對率"])
                 for s in (q.get("子題") or []) if s.get("答對率") is not None]
        if rates:
            add("右下角是每個子題的答對率。")
            add("、".join(f"{lab}{cn_pct(r)}" for lab, r in rates[:4]) + "。")
        else:
            add(f"這一題的作答完成率是{pct(q.get('作答率', 0))}。", optional=4)
        add("以上是這一題的作答情形，接下來看下一題。", must=True)
        return _trim_notes(L, [i for _pri, i in sorted(opt)], protect=keep_idx)

    # 四類提問：則數合成一句講完，再挑「則數最多的那一類」講一則代表句
    if asks:
        add("右下角的四類提問分別是"
            + "、".join(f"{a['類別']}{cn_num(a['提問數'])}則" for a in asks[:4])
            + "。", must=True)
        top = max(asks[:4], key=lambda a: a["提問數"])
        if top["提問數"] and top.get("代表句"):
            rep = top["代表句"]
            rep = (rep[:18] + "…") if len(rep) > 18 else rep
            add(f"其中{top['類別']}最多，代表問題是「{rep}」。", must=True)
        elif not any(a["提問數"] for a in asks[:4]):
            add("這一週這一題沒有同學提問。", must=True)
    if q.get("其他提問數"):
        add(f"另外還有{cn_num(q['其他提問數'])}則沒辦法歸類，放在提問清單裡。", optional=2)
    add("以上是這一題的作答情形，接下來看下一題。", must=True)

    # 可省略句子的拿掉順序：先拿資訊量最低的；四類提問／代表句／結尾句不可被砍
    return _trim_notes(L, [i for _pri, i in sorted(opt)], protect=keep_idx)


def build_deck(analysis, out_path, meta, wc_map, log=print):
    """依 analysis.json 內容組裝簡報。meta 需含 course_name/ta_name/teacher/week_label/…"""
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    qs = analysis["題目"]
    series = meta["series"]
    logo = meta.get("logo")
    week_label = meta.get("week_label", "")
    course = meta.get("course_name", "")

    n_q = len(qs)
    n_ans = sum(q["作答人數"] for q in qs)
    avg_rate = round(sum(q["作答率"] for q in qs) / max(n_q, 1), 1)
    # 平均覆蓋率只算「有開放文字作答」的題；選擇題／測驗題沒有覆蓋率，
    # 算進去只會把平均拉低成一個沒意義的數字。
    cov_qs = [q for q in qs if q.get("重點概念")]
    avg_cov = round(sum(q.get("整體覆蓋率", 0) for q in cov_qs) * 100
                    / max(len(cov_qs), 1), 1)
    n_ask = sum(q.get("提問總數", 0) for q in qs)
    top10 = analysis.get("全班重點TOP10", [])
    top_words = [w for w, _ in top10[:5]]

    # 附錄頁：有概念矩陣才做（config.json 的 appendix_matrix 預設 true）
    # 2.2：概念欄從 3 個變 6 個 → 每頁只放 2 題，最多 3 頁，其餘請看 CSV
    appendix_qs = [q for q in qs if q.get("重點概念")]
    want_appendix = bool(meta.get("appendix_matrix", True)) and bool(appendix_qs)
    apx_pages = []
    if want_appendix:
        for i in range(0, len(appendix_qs), APX_PER_PAGE):
            if len(apx_pages) >= APX_MAX_PAGES:
                break
            apx_pages.append(appendix_qs[i:i + APX_PER_PAGE])
    apx_shown = sum(len(p) for p in apx_pages)
    total = 3 + n_q + len(apx_pages)

    # ---- 1. 封面
    slide_cover(prs, {
        "眉標": meta.get("eyebrow", "WEEKLY HOMEWORK ANALYTICS"),
        "主標": f"{course}　同學作答分析\n{week_label}",
        "英文副標": "Concept Coverage, Word Cloud & Question Taxonomy",
        "關鍵字": (top_words or ["文字雲", "作答分析"])[:3] + ["概念矩陣", "學生編號"],
        "報告人資訊": (f"授課教師：{meta.get('teacher', '')}\n"
                   f"教學助理：{meta.get('ta_name', '')}\n"
                   f"製作日期：{meta.get('date', '')}"),
        "logo": logo,
        "逐字稿": [
            "老師好，我是這門課的教學助理。",
            f"這是{course}{week_label}的同學作答分析。",
            f"這一次一共整理了{cn_num(n_q)}題。",
            "報告裡的同學一律用學生編號表示，姓名已經遮罩，學號也全部改成大寫的 O。",
            "每一題我會報告六個重點的覆蓋率，還有同學提問的四種分類。",
            "接下來先看整體作答情形。",
        ],
    })

    # ---- 2. 總覽
    slide_overview(prs, {
        "頁碼標": 1, "眉標": "OVERVIEW",
        "標題": "檢視本週整體作答與概念覆蓋",
        "副標": f"{week_label}　各題作答率與「至少提到一個重點」的整體覆蓋率",
        "總結": (f"本週{cn_num(n_q)}題平均作答率 {avg_rate}%；"
               f"其中 {len(cov_qs)} 題開放作答的平均整體覆蓋率 {avg_cov}%；"
               f"同學共提出 {n_ask} 則問題"),
        "logo": logo,
        "逐字稿": [
            "這一頁是本週的整體作答情形。",
            "表格由左到右分別是題號、題目、作答人數比全班人數、作答率、整體覆蓋率，"
            "最右邊是覆蓋率長條。",
            f"本週一共{cn_num(n_q)}題，平均作答率是{cn_pct(avg_rate)}。",
            f"累計的作答人次是{cn_num(n_ans)}人次。",
            f"整體覆蓋率的意思是，至少提到六個重點其中一個的同學佔多少比例，"
            f"本週{cn_num(len(cov_qs))}題開放作答的平均是{cn_pct(avg_cov)}。",
            f"同學這一週一共提出{cn_num(n_ask)}則問題。",
        ] + ([f"全班最常出現的關鍵詞是「{top_words[0]}」。"] if top_words else []) + [
            "下面的深藍色橫幅是這一頁的一句話重點。",
            "接下來我逐題說明。",
        ],
    }, qs)
    footer(prs.slides[-1], series, 2, total)

    # ---- 3. 每題一頁
    for i, q in enumerate(qs, 1):
        png = wc_map.get(q["題號"])
        focus = q.get("重點概念") or []
        asks = q.get("提問統計") or []
        # 開放文字題的左欄變窄（5.00 吋），圖說改成三行短句，每行都會過 fit_one_line
        cap = (f"圖 {i}　{q['題號']} 作答文字雲（{course}）\n"
               f"資料來源：Zuvio 下載數據，{week_label}\n"
               f"字級與出現次數成正比；同學以學生編號表示，已去識別化。")
        if not png or not os.path.exists(png):
            cap = (f"圖 {i}　{q['題號']} 選項分佈（{week_label}，{course}）\n"
                   f"資料來源：Zuvio 下載數據（已去識別化）；本題為選擇題／測驗題，不做文字雲。")
        # 選擇題／測驗題：每個子題只列「最多人選的那個選項」，並標上子題號，
        # 免得多個子題的選項混在一起看不出誰是誰。
        opts, sub_opts, rates = [], [], []
        for s in q.get("子題", []):
            dist = s.get("選項分佈") or []
            lab = re.sub(r":.*$", "", s.get("子題號", "")) or "本題"
            if dist:
                opts.append({"選項": f"{lab}　{dist[0]['選項']}",
                             "人數": dist[0]["人數"], "百分比": dist[0]["百分比"]})
                sub_opts.append({"子題": lab, "選項": dist})
            if s.get("答對率") is not None:
                rates.append((lab, s["答對率"]))
        if rates:
            ans_title = "答對率（各子題）"
            ans_body = "　".join(f"{lab} {r}%" for lab, r in rates[:5])
        else:
            ans_title = "作答完成率"
            ans_body = (f"{q['作答人數']} / {q['全班人數']} 人作答，"
                        f"完成率 {q['作答率']}%")
        if focus:
            # 底部橫幅：前三大重點列出來，其餘請看右上角那張六列的表（必須單行）
            summary = (f"{q['題號']} 前三大重點：" +
                       "、".join(f"{f['概念']}（{round(f['覆蓋率'] * 100)}%）"
                                for f in focus[:3]))
            if len(focus) > 3:
                summary += f"；另{cn_num(len(focus) - 3)}個重點見右上表"
        else:
            summary = (f"{q['題號']} 為選擇題／測驗題，改看選項分佈與"
                       + ("答對率" if rates else "作答完成率"))
        slide_question(prs, {
            "頁碼標": i + 1, "眉標": f"QUESTION {i:02d}",
            "標題": f"解讀 {q['題號']}：{q['題目']}",
            "副標": (f"作答 {q['作答人數']} / {q['全班人數']} 人　作答率 {q['作答率']}%　"
                   f"開放文字 {q.get('文字列數', q['有效列數'])} 則　"
                   f"平均 {q['平均字數']} 字"),
            "圖說": cap,
            "重點標題": (f"{cn_num(len(focus))}個重點（依提及人數排序）"
                    f"　概念｜提及人數｜覆蓋率" if focus else "重點概念"),
            "重點": focus,
            "覆蓋率總結": (f"整體覆蓋率 {round(q.get('整體覆蓋率', 0) * 100)}%"
                     f"（至少 1 個）　提到 ≥3 個 "
                     f"{round(q.get('提到三個以上比例', 0) * 100)}%　"
                     f"全部提到 {round(q.get('全部提到比例', 0) * 100)}%"),
            "提問標題": (f"四類提問　共 {q.get('提問總數', 0)} 則／"
                    f"{q.get('提問人數', 0)} 人（其他 {q.get('其他提問數', 0)} 則）"),
            "提問": asks,
            "選項分佈": opts,
            "選項標題": f"各子題選項分佈（共 {len(sub_opts)} 個子題）",
            "子題選項": sub_opts,
            "答對標題": ans_title,
            "答對內容": ans_body,
            "大數字": f"{q['作答人數']}",
            "大數字說明": "位同學作答",
            "總結": summary,
            "logo": logo,
            "逐字稿": _q_notes(q, i, n_q, week_label),
        }, png, i)
        footer(prs.slides[-1], series, i + 2, total)

    # ---- 4. 附錄：概念矩陣小表（學生一律依**編號數字**排序，115-1_EC_2 在 _10 前面）
    def _code_num(code):
        m = re.search(r"_(\d+)$", str(code or ""))
        return (0, int(m.group(1))) if m else (1, 0)

    n_apx = len(apx_pages)
    for pi, group in enumerate(apx_pages, 1):
        tables = [{"標題": f"{q['題號']}（前十位）",
                   "概念": [f["概念"] for f in q["重點概念"]],
                   "學生編號": sorted(q.get("概念矩陣", {}).keys(),
                                  key=_code_num)[:10],
                   "矩陣": q.get("概念矩陣", {})} for q in group]
        n_show = sum(len(t["學生編號"]) for t in tables)
        rest = len(appendix_qs) - apx_shown
        note = ("完整矩陣見輸出資料夾 Q0N_concept_matrix.csv；"
                "提問明細見 Q0N_questions.csv；表內一律使用學生編號。")
        if rest > 0 and pi == n_apx:
            note = f"另有 {rest} 題的概念矩陣受版面限制未列，請看 CSV；" + note
        page_tag = f"（{pi} / {n_apx}）" if n_apx > 1 else ""
        slide_appendix(prs, {
            "頁碼標": n_q + 1 + pi, "眉標": "APPENDIX　CONCEPT MATRIX",
            "標題": f"對照概念矩陣（前十位學生編號）{page_tag}",
            "副標": "①–⑥ 是這一題的六個重點；1 代表該生的回答有提到，0 代表沒有提到",
            "附註": note,
            "總結": (f"附錄第 {pi} 頁列出 {len(tables)} 題、{n_show} 列概念矩陣節錄，"
                   f"完整資料在輸出資料夾的 CSV"),
            "logo": logo,
            "逐字稿": [
                f"這一頁是附錄的第{cn_num(pi)}頁，我把概念矩陣節錄出來給老師看。",
                "每一張小表的左邊是學生編號，右邊六欄分別對應這一題的六個重點。",
                "圈一到圈六各代表哪一個重點，寫在表格上面的圖例。",
                "一代表這位同學的回答有提到這個重點，零代表沒有提到。",
                f"因為版面有限，這一頁放{cn_num(len(tables))}題、每題前十位，"
                f"一共{cn_num(n_show)}列。",
                "完整的矩陣在輸出資料夾裡，檔名是 Q 零 N 底線 concept matrix 點 csv。",
                "表格裡只會出現學生編號，不會出現姓名和學號。",
            ],
        }, tables)
        footer(prs.slides[-1], series, n_q + 2 + pi, total)

    # ---- 5. 結尾
    slide_closing(prs, {
        "眉標": "NEXT STEPS",
        "主標": "以上是本週的作答分析\n謝謝老師",
        "說明": ("下一步：把覆蓋率偏低的重點帶回課堂補強，並回應同學的四類提問。\n"
               "所有輸出檔一律使用學生編號，姓名已遮罩、學號已全部改成 O。"),
        "逐字稿": [
            "這是最後一頁。",
            f"本週一共分析了{cn_num(n_q)}題，平均作答率{cn_pct(avg_rate)}，"
            f"開放作答題的平均整體覆蓋率{cn_pct(avg_cov)}。",
            f"同學一共提出{cn_num(n_ask)}則問題，"
            "我已經分成計算數據、操作應用、延伸探究和概念理解四類。",
            "我會先回應提問人數比較多的那一類，再補強覆蓋率偏低的重點。",
            "所有輸出的檔案都只有學生編號，原始檔和對照表只留在我的電腦裡。",
            "以上是本週報告，謝謝老師。",
        ],
    })

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    log(f"  已產生簡報：{out_path}（{total} 頁）")
    return prs, total
