# -*- coding: utf-8 -*-
"""用 PIL 產生 icon.ico / icon.png（深藍底＋白色「姓O」意象）。

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
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjhbd.ttc",
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
    d.rounded_rectangle([int(px * 0.08), int(px * 0.08), int(px * 0.92), int(px * 0.46)],
                        radius=int(px * 0.06), fill=TEAL)
    # 上半：白色「姓」；下半：金色大 O（遮罩符號）
    f1 = load_font(int(px * 0.30))
    f2 = load_font(int(px * 0.40))
    d.text((px * 0.5, px * 0.27), "姓名", font=f1, fill=WHITE, anchor="mm")
    d.text((px * 0.5, px * 0.71), "O", font=f2, fill=GOLD, anchor="mm")
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
