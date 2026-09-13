# -*- coding: utf-8 -*-
"""
仿生與環境工作坊・個人週報產生器　共用版式函式庫（自然科學雜誌風）

來源：D:\\02-aitt_claude_agent\\weekly-report-workshop\\core.py（複製後重構）
重構重點（桌面版必要）：
  1. import 時**不產生任何副作用**：不建立 Presentation、不建資料夾、不寫死輸出路徑。
     全部收斂進 Deck 類別，可以在同一個行程裡重複呼叫、產生多份週報。
  2. 圖號計數器、總頁數、系列名改成 Deck 的屬性（原本是模組全域）。
  3. 快取資料夾改成可設定，預設放在使用者暫存資料夾。
  4. 字型：Windows 微軟正黑體／macOS PingFang TC／其他退回 sans-serif；**不打包字型檔**。

規格來源：YCYLabWKReport.md（週報規格・唯一真相來源）
授權：CC BY-NC-SA　｜　仿生與環境工作坊
"""
import os
import sys
import hashlib
import tempfile

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------- 海洋科學色票（十色，不要自己改 hex） ----------
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
DEEP = RGBColor(0x06, 0x5A, 0x82)
TEAL = RGBColor(0x1C, 0x72, 0x93)
GREEN = RGBColor(0x2E, 0x7D, 0x5B)
MOSS = RGBColor(0x5F, 0xA9, 0x8A)
GOLD = RGBColor(0xC9, 0x9A, 0x3E)
PAGEBG = RGBColor(0xF5, 0xF9, 0xFA)
BODY = RGBColor(0x1E, 0x29, 0x3B)
AUX = RGBColor(0x64, 0x74, 0x8B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xA8, 0x4A, 0x3C)
LINE = RGBColor(0xE2, 0xE8, 0xEC)
DARK1 = RGBColor(0x10, 0x2C, 0x4E)
DARK2 = RGBColor(0x0E, 0x27, 0x45)
SUBTXT = RGBColor(0x9E, 0xC5, 0xD8)

W, H = 13.333, 7.5                      # 16:9
CIRC = [TEAL, DEEP, GREEN, GOLD, NAVY]  # 淺底頁面：編號圓的輪替色
CHIP = [TEAL, DEEP, GREEN, GOLD, MOSS]  # 深底頁面：膠囊標籤的輪替色（不用 NAVY，會看不見）
CN_NUM = ["壹", "貳", "參", "肆", "伍", "陸"]   # 中文序號（驗收錨點，不可改成 1.2.3.）


# ---------------------------------------------------------------- 字型
def pick_font():
    """簡報用的中文字型名稱。Windows 微軟正黑體；macOS PingFang TC；其餘 sans-serif。
    只寫「字型名稱」進 pptx，不會、也不可以把字型檔打包進 exe（授權）。"""
    if sys.platform.startswith("win"):
        return "微軟正黑體"
    if sys.platform == "darwin":
        return "PingFang TC"
    return "Noto Sans CJK TC"


def pick_pil_font_path():
    """PIL 畫佔位圖用的實體字型檔；找不到回 None（呼叫端退回 PIL 預設點陣字）。"""
    cands = [
        r"C:\Windows\Fonts\msjh.ttc", r"C:\Windows\Fonts\msjhbd.ttc",
        r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\arial.ttf",
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for p in cands:
        if os.path.exists(p):
            return p
    return None


FONT = pick_font()


def resource_path(rel):
    """PyInstaller onefile 下用 sys._MEIPASS 找 datas；開發時用原始碼資料夾。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def default_cache_dir():
    d = os.path.join(tempfile.gettempdir(), "WeeklyReportMaker_cache")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------- 照片處理
def crop(path, w_in, h_in, cache_dir=None):
    """依目標比例置中裁切，>1700px 縮圖後存 JPEG q86，結果以 md5 快取。"""
    cache_dir = cache_dir or default_cache_dir()
    key = hashlib.md5(f"{path}|{w_in:.2f}|{h_in:.2f}".encode("utf-8")).hexdigest()[:16]
    o = os.path.join(cache_dir, key + ".jpg")
    if not os.path.exists(o):
        im = Image.open(path).convert("RGB")
        tw, th = im.size
        t = w_in / h_in
        if tw / th > t:
            nw = int(th * t)
            x0 = (tw - nw) // 2
            im = im.crop((x0, 0, x0 + nw, th))
        else:
            nh = int(tw / t)
            y0 = (th - nh) // 2
            im = im.crop((0, y0, tw, y0 + nh))
        if im.size[0] > 1700:
            im = im.resize((1700, int(im.size[1] * 1700 / im.size[0])), Image.LANCZOS)
        im.save(o, "JPEG", quality=86)
    return o


def crop_top(path, w_in, h_in, cache_dir=None):
    """由頂端裁切（適合網頁／文件截圖，標題不會被切掉）。"""
    cache_dir = cache_dir or default_cache_dir()
    key = hashlib.md5(f"T{path}|{w_in:.2f}|{h_in:.2f}".encode("utf-8")).hexdigest()[:16]
    o = os.path.join(cache_dir, key + ".jpg")
    if not os.path.exists(o):
        im = Image.open(path).convert("RGB")
        tw, th = im.size
        t = w_in / h_in
        nh = int(tw / t)
        if nh <= th:
            im = im.crop((0, 0, tw, nh))
        else:
            im = im.crop((0, 0, int(th * t), th))
        if im.size[0] > 1700:
            im = im.resize((1700, int(im.size[1] * 1700 / im.size[0])), Image.LANCZOS)
        im.save(o, "JPEG", quality=88)
    return o


# ---------------------------------------------------------------- 基本積木
def set_run(r, s, size, color, bold=False, italic=False, font=None):
    """設定一段文字。a:ea／a:cs 一定要補，否則中文會 fallback 成別的字型。"""
    font = font or FONT
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = etree.SubElement(rPr, qn(tag))
        e.set("typeface", font)


def tbox(sl, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return tf


def text(sl, s, x, y, w, h, size, color, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, ls=None):
    tf = tbox(sl, x, y, w, h, anchor)
    for i, ln in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if ls:
            p.line_spacing = ls
        set_run(p.add_run(), ln, size, color, bold, italic)
    return tf


def bullets(sl, items, x, y, w, h, size=14, color=BODY, space=8, ls=1.18):
    tf = tbox(sl, x, y, w, h)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space)
        p.line_spacing = ls
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "200000")
        pPr.set("indent", "-200000")
        bf = etree.SubElement(pPr, qn("a:buFont"))
        bf.set("typeface", "Arial")
        bc = etree.SubElement(pPr, qn("a:buChar"))
        bc.set("char", "•")
        set_run(p.add_run(), it, size, color)
    return tf


def shape(sl, st, x, y, w, h, fill, radius=None, line=None, line_w=1.0):
    sp = sl.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    if radius is not None and st == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def card(sl, x, y, w, h, radius=0.05):
    return shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, WHITE, radius=radius)


def rule(sl, x, y, w, color=LINE, t=0.012):
    shape(sl, MSO_SHAPE.RECTANGLE, x, y, w, t, color)


def bg(sl, color):
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = color


def notes(sl, lines, font=None):
    """逐字稿：一句一行寫進備忘稿，每一行 12pt、字型與簡報同。"""
    font = font or FONT
    tf = sl.notes_slide.notes_text_frame
    tf.text = chr(10).join([ln for ln in lines if str(ln).strip()])
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(12)
            r.font.name = font
            rPr = r._r.get_or_add_rPr()
            for tag in ("a:ea", "a:cs"):
                e = rPr.find(qn(tag))
                if e is None:
                    e = etree.SubElement(rPr, qn(tag))
                e.set("typeface", font)


def banner(sl, s, x=0.62, y=6.28, w=12.1, h=0.6, size=15):
    shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, NAVY, radius=0.16)
    text(sl, s, x + 0.3, y, w - 0.6, h, size, WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, ls=1.15)


def table(sl, x, y, cols, rows, head_c=None, row_h=0.42, fs=12.5, hfs=12.5, band=None):
    """文字排版式表格（不用原生表格，才控得住字型與字級）。
    cols: [(欄名, 寬吋, 對齊)]；rows: [[(文字,)/(文字,色)/(文字,色,粗體), ...]]"""
    tw = sum(c[1] for c in cols)
    cx = x
    for i, (lab, cw, al) in enumerate(cols):
        text(sl, lab, cx, y, cw, 0.34, hfs, (head_c[i] if head_c else AUX), bold=True,
             align=al, anchor=MSO_ANCHOR.MIDDLE)
        cx += cw
    rule(sl, x, y + 0.36, tw, RGBColor(0xC8, 0xD4, 0xDA), 0.014)
    ry = y + 0.46
    for ri, row in enumerate(rows):
        if band is not None and ri in band:
            shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x - 0.12, ry - 0.03,
                  tw + 0.24, row_h + 0.06, RGBColor(0xFA, 0xF2, 0xDF), radius=0.14)
        cx = x
        for ci, cell in enumerate(row):
            if len(cell) == 3:
                t, col, bd = cell
            elif len(cell) == 2:
                t, col, bd = cell[0], cell[1], False
            else:
                t, col, bd = cell[0], BODY, False
            text(sl, t, cx, ry, cols[ci][1], row_h, fs, col, bold=bd,
                 align=cols[ci][2], anchor=MSO_ANCHOR.MIDDLE)
            cx += cols[ci][1]
        rule(sl, x, ry + row_h, tw)
        ry += row_h + 0.04
    return ry


def flow(sl, steps, x, y, w, h, per_row=None, gap=0.34, dsize=11.5):
    """流程圖：編號彩色圓＋› 箭頭串接。steps: [(標題, 說明)]"""
    n = len(steps)
    per_row = per_row or n
    cw = (w - gap * (per_row - 1)) / per_row
    for i, (t, d) in enumerate(steps):
        r, c = divmod(i, per_row)
        cx = x + c * (cw + gap)
        cy = y + r * (h + 0.42)
        card(sl, cx, cy, cw, h)
        shape(sl, MSO_SHAPE.OVAL, cx + (cw - 0.62) / 2, cy + 0.2, 0.62, 0.62, CIRC[i % 5])
        text(sl, f"{i+1:02d}", cx + (cw - 0.62) / 2, cy + 0.2, 0.62, 0.62, 14, WHITE,
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, t, cx + 0.08, cy + 0.94, cw - 0.16, 0.4, 13.5, NAVY, bold=True,
             align=PP_ALIGN.CENTER)
        text(sl, d, cx + 0.12, cy + 1.36, cw - 0.24, h - 1.46, dsize, AUX,
             align=PP_ALIGN.CENTER, ls=1.2)
        if c < per_row - 1 and i < n - 1:
            text(sl, "›", cx + cw, cy + 0.26, gap, 0.5, 20, MOSS, bold=True,
                 align=PP_ALIGN.CENTER)


def kpi(sl, items, x, y, w, h, cols=3, gapx=0.3, gapy=0.3):
    """關鍵數字卡：items = [(值, 單位, 說明, 色)]。"""
    cw = (w - gapx * (cols - 1)) / cols
    rows = (len(items) + cols - 1) // cols
    ch = (h - gapy * (rows - 1)) / rows
    for i, (val, unit, label, c) in enumerate(items):
        cx = x + (i % cols) * (cw + gapx)
        cy = y + (i // cols) * (ch + gapy)
        card(sl, cx, cy, cw, ch)
        shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, 0.09, ch, c, radius=0.5)
        text(sl, val, cx + 0.32, cy + 0.18, cw - 0.6, 0.62, 30, c, bold=True)
        if unit:
            text(sl, unit, cx + 0.32, cy + 0.84, cw - 0.6, 0.3, 12, AUX)
        text(sl, label, cx + 0.32, cy + 1.2, cw - 0.62, ch - 1.3, 13, BODY, ls=1.18)


# ---------------------------------------------------------------- Deck
class Deck:
    """一份簡報。所有需要「跨頁記憶」的東西（圖號、頁碼、總頁數）都掛在這裡，
    所以同一個行程可以連續產生很多份週報，彼此不會互相污染。"""

    def __init__(self, series="", total=0, cache_dir=None, font=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.series = series
        self.total = total
        self.fig = 0
        self.cache_dir = cache_dir or default_cache_dir()
        self.font = font or FONT

    # ---- 頁面骨架 ----
    def blank_slide(self, color=PAGEBG):
        sl = self.prs.slides.add_slide(self.blank)
        bg(sl, color)
        return sl

    def footer(self, sl, num):
        text(sl, self.series, 0.62, 7.02, 6.6, 0.3, 10, AUX)
        text(sl, f"{num} / {self.total}", 11.0, 7.02, 1.7, 0.3, 10, AUX,
             align=PP_ALIGN.RIGHT)

    def cover(self, headline_en, title, subtitle_en, chips, lines):
        """封面（深色）：眉標 → 38pt 大標 → 斜體英文副標 → 膠囊關鍵字 → 分隔線＋報告人。"""
        sl = self.blank_slide(NAVY)
        shape(sl, MSO_SHAPE.OVAL, 8.9, -2.3, 7.2, 7.2, DARK1)
        shape(sl, MSO_SHAPE.OVAL, 10.6, 4.2, 4.6, 4.6, DARK2)
        text(sl, headline_en, 0.92, 1.42, 6.0, 0.34, 13, MOSS, bold=True)
        text(sl, title, 0.92, 1.9, 8.6, 1.9, 38, WHITE, bold=True, ls=1.18)
        text(sl, subtitle_en, 0.92, 3.86, 9.4, 0.4, 15, SUBTXT, italic=True)
        px = 0.92
        for i, t in enumerate(chips):
            wpx = 0.36 + len(t) * 0.19
            if px + wpx > 12.6:
                break
            shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, px, 4.44, wpx, 0.46, CHIP[i % 5], radius=0.5)
            text(sl, t, px, 4.44, wpx, 0.46, 13, WHITE, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            px += wpx + 0.18
        shape(sl, MSO_SHAPE.RECTANGLE, 0.92, 5.34, 5.6, 0.02, RGBColor(0x2B, 0x4B, 0x67))
        yy = 5.58
        for i, ln in enumerate(lines):
            text(sl, ln, 0.92, yy, 10.4, 0.36, 15 if i == 0 else 14,
                 WHITE if i == 0 else SUBTXT)
            yy += 0.44
        return sl

    def page(self, num, label_en, title, subtitle, pageno):
        """內容頁骨架：眉標 12pt TEAL → 標題 30pt NAVY → 斜體副標 14pt → 頁尾。"""
        sl = self.blank_slide(PAGEBG)
        text(sl, f"{num:02d}  /  {label_en}", 0.62, 0.42, 8.6, 0.32, 12, TEAL, bold=True)
        text(sl, title, 0.62, 0.78, 11.9, 0.62, 30, NAVY, bold=True)
        text(sl, subtitle, 0.62, 1.46, 11.9, 0.34, 14, AUX, italic=True)
        self.footer(sl, pageno)
        return sl

    def divider(self, part_en, title, sub, chips=None):
        """分節頁（深色）。"""
        sl = self.blank_slide(NAVY)
        shape(sl, MSO_SHAPE.OVAL, 9.4, -2.2, 6.6, 6.6, DARK1)
        shape(sl, MSO_SHAPE.OVAL, 10.9, 4.4, 4.2, 4.2, DARK2)
        text(sl, part_en, 0.92, 2.26, 7.0, 0.36, 14, MOSS, bold=True)
        text(sl, title, 0.92, 2.72, 8.4, 0.96, 38, WHITE, bold=True)
        text(sl, sub, 0.92, 3.82, 8.2, 1.0, 15, SUBTXT, ls=1.3)
        if chips:
            px = 0.92
            for t, c in chips:
                wpx = 0.36 + len(t) * 0.185
                shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, px, 4.92, wpx, 0.44, c, radius=0.5)
                text(sl, t, px, 4.92, wpx, 0.44, 12.5, WHITE, bold=True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                px += wpx + 0.16
        return sl

    def closing(self, headline_en, title, subtitle_en, lines):
        sl = self.blank_slide(NAVY)
        shape(sl, MSO_SHAPE.OVAL, -2.6, 3.6, 7.4, 7.4, DARK1)
        shape(sl, MSO_SHAPE.OVAL, 10.2, -1.8, 5.4, 5.4, DARK2)
        text(sl, headline_en, 1.1, 2.4, 6.0, 0.4, 13, MOSS, bold=True)
        text(sl, title, 1.1, 2.9, 9.4, 0.9, 38, WHITE, bold=True)
        text(sl, subtitle_en, 1.1, 3.9, 9.8, 0.4, 15, SUBTXT, italic=True)
        shape(sl, MSO_SHAPE.RECTANGLE, 1.1, 4.56, 5.6, 0.02, RGBColor(0x2B, 0x4B, 0x67))
        yy = 4.82
        for i, ln in enumerate(lines):
            text(sl, ln, 1.1, yy, 10.6, 0.36, 14 if i == 0 else 13,
                 WHITE if i == 0 else SUBTXT)
            yy += 0.44
        return sl

    # ---- 照片 ----
    def figure(self, sl, path, x, y, w, h, cap, fit=False, top=False, cap_h=0.5):
        """白色圓角相框＋等比不變形的照片＋灰色圖說「圖 N　說明（日期，地點）」。
        找不到照片時畫灰色佔位卡，不報錯——缺一張圖不該讓整份週報跑不出來。"""
        self.fig += 1
        card(sl, x, y, w + 0.3, h + 0.32 + cap_h)
        ok = bool(path) and os.path.exists(path)
        if not ok:
            shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.15, y + 0.15, w, h,
                  RGBColor(0xDD, 0xE4, 0xE8), radius=0.03)
            text(sl, "照片佔位", x + 0.15, y + 0.15, w, h, 14, AUX, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        else:
            try:
                if fit:
                    iw, ih = Image.open(path).size
                    sc = min(w / iw, h / ih)
                    pw, ph = iw * sc, ih * sc
                    sl.shapes.add_picture(path, Inches(x + 0.15 + (w - pw) / 2),
                                          Inches(y + 0.15 + (h - ph) / 2),
                                          Inches(pw), Inches(ph))
                else:
                    src = (crop_top if top else crop)(path, w, h, self.cache_dir)
                    sl.shapes.add_picture(src, Inches(x + 0.15), Inches(y + 0.15),
                                          Inches(w), Inches(h))
            except Exception:
                shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.15, y + 0.15, w, h,
                      RGBColor(0xDD, 0xE4, 0xE8), radius=0.03)
                text(sl, "照片讀取失敗", x + 0.15, y + 0.15, w, h, 14, AUX, bold=True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, f"圖 {self.fig}　{cap}", x + 0.15, y + h + 0.26, w, cap_h, 11, AUX, ls=1.15)
        return self.fig

    # ---- 備忘稿 ----
    def notes(self, sl, lines):
        notes(sl, lines, self.font)

    # ---- 輸出 ----
    def save(self, path):
        d = os.path.dirname(os.path.abspath(path))
        if d:
            os.makedirs(d, exist_ok=True)
        self.prs.save(path)
        return path

    @property
    def slide_count(self):
        return len(self.prs.slides._sldIdLst)
