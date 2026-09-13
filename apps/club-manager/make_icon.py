# -*- coding: utf-8 -*-
"""用 PIL 產生 icon.ico / icon.png（深藍底＋青藍上緣＋白色「社」與金色星點）。

執行：python make_icon.py
"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY = (11, 31, 58, 255)
TEAL = (28, 114, 147, 255)
WHITE = (255, 255, 255, 255)
GOLD = (201, 154, 62, 255)
MOSS = (95, 169, 138, 255)

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
    r = int(px * 0.20)
    d.rounded_rectangle([0, 0, px - 1, px - 1], radius=r, fill=NAVY)
    # 上緣青藍帶＝「仰望」的天空
    d.rounded_rectangle([int(px * 0.08), int(px * 0.08), int(px * 0.92), int(px * 0.38)],
                        radius=int(px * 0.06), fill=TEAL)
    # 三顆星（仰望）
    for cx, cy, rr in ((0.26, 0.21, 0.035), (0.50, 0.16, 0.05), (0.74, 0.24, 0.03)):
        d.ellipse([px * (cx - rr), px * (cy - rr), px * (cx + rr), px * (cy + rr)], fill=GOLD)
    # 下半：白色「社團」
    f = load_font(int(px * 0.30))
    d.text((px * 0.5, px * 0.63), "社團", font=f, fill=WHITE, anchor="mm")
    # 底線：森綠
    d.rounded_rectangle([int(px * 0.28), int(px * 0.82), int(px * 0.72), int(px * 0.86)],
                        radius=int(px * 0.02), fill=MOSS)
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
