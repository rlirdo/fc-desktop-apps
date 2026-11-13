# -*- coding: utf-8 -*-
"""用 PIL 產生 icon.ico / icon.png（深藍底＋白色「客」＋金色聲波，風格比照其他 app）。

執行：python make_icon.py
"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY = (11, 31, 58, 255)
TEAL = (28, 114, 147, 255)
WHITE = (255, 255, 255, 255)
GOLD = (201, 154, 62, 255)

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msjhbd.ttc",
    r"C:\Windows\Fonts\msjh.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
]


def load_font(size):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def draw_icon(px):
    img = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, px - 1, px - 1], radius=int(px * 0.20), fill=NAVY)
    d.rounded_rectangle([int(px * 0.08), int(px * 0.10), int(px * 0.62), int(px * 0.90)],
                        radius=int(px * 0.08), fill=TEAL)
    d.text((px * 0.35, px * 0.50), "客", font=load_font(int(px * 0.42)), fill=WHITE, anchor="mm")
    # 金色聲波（三道弧）
    cx, cy = px * 0.62, px * 0.50
    w = max(1, int(px * 0.045))
    for i, r in enumerate((0.12, 0.21, 0.30)):
        rr = px * r
        d.arc([cx - rr, cy - rr, cx + rr, cy + rr], start=-45, end=45, fill=GOLD, width=w)
    return img


def main():
    sizes = [256, 128, 64, 48, 32, 16]
    imgs = [draw_icon(s) for s in sizes]
    ico = os.path.join(HERE, "icon.ico")
    imgs[0].save(ico, format="ICO", sizes=[(s, s) for s in sizes])
    png = os.path.join(HERE, "icon.png")
    imgs[0].save(png, format="PNG")
    print("已產生：" + ico)
    print("已產生：" + png)


if __name__ == "__main__":
    main()
