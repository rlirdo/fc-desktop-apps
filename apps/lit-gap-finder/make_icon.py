# -*- coding: utf-8 -*-
"""用 PIL 產生 icon.ico / icon.png（深藍底＋青藍「文獻」標籤＋金色放大鏡，鏡中一格空缺＝缺口）。

執行：python make_icon.py
"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY = (11, 31, 58, 255)
TEAL = (28, 114, 147, 255)
MOSS = (95, 169, 138, 255)
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
    s = 4  # 超取樣後縮小，邊緣較平滑
    P = px * s
    img = Image.new("RGBA", (P, P), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, P - 1, P - 1], radius=int(P * 0.20), fill=NAVY)
    d.rounded_rectangle([int(P * 0.08), int(P * 0.07), int(P * 0.92), int(P * 0.40)],
                        radius=int(P * 0.06), fill=TEAL)
    d.text((P * 0.5, P * 0.235), "文獻", font=load_font(int(P * 0.24)), fill=WHITE, anchor="mm")
    # 放大鏡：鏡內 3×2 格子，其中一格留空（缺口）
    cx, cy, r = P * 0.42, P * 0.66, P * 0.19
    lw = max(2, int(P * 0.05))
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=GOLD, width=lw)
    g = r * 0.42
    for i, (dx, dy) in enumerate([(-1, -0.5), (0, -0.5), (1, -0.5), (-1, 0.5), (0, 0.5), (1, 0.5)]):
        x0 = cx + dx * g * 1.05 - g * 0.45
        y0 = cy + dy * g * 1.05 - g * 0.45
        if i == 4:
            d.rectangle([x0, y0, x0 + g * 0.9, y0 + g * 0.9], outline=MOSS, width=max(1, lw // 3))
        else:
            d.rectangle([x0, y0, x0 + g * 0.9, y0 + g * 0.9], fill=WHITE)
    hx, hy = cx + r * 0.72, cy + r * 0.72
    d.line([hx, hy, P * 0.86, P * 0.92], fill=GOLD, width=int(lw * 1.6))
    return img.resize((px, px), Image.LANCZOS)


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
