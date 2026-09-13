# -*- coding: utf-8 -*-
"""
make_icon.py — 用 PIL 畫出 icon.ico（Windows）與 icon.png（macOS 打包用）。

圖案：深藍圓底，上面散落大小不一的青／綠／金色圓點，像一團文字雲。
純程式產生，不使用任何外部圖檔或字型（避免授權問題）。

執行：python make_icon.py
"""
import os
import math
import random

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
SIZE = 1024

NAVY = (11, 31, 58, 255)
DEEP = (6, 90, 130, 255)
TEAL = (28, 114, 147, 255)
GREEN = (46, 125, 91, 255)
MOSS = (95, 169, 138, 255)
GOLD = (201, 154, 62, 255)
WHITE = (255, 255, 255, 255)

BLOBS = [
    # (相對 x, 相對 y, 半徑, 顏色)
    (0.50, 0.44, 0.150, WHITE),
    (0.28, 0.36, 0.085, MOSS),
    (0.72, 0.34, 0.070, GOLD),
    (0.30, 0.62, 0.075, TEAL),
    (0.70, 0.63, 0.090, GREEN),
    (0.50, 0.74, 0.055, MOSS),
    (0.50, 0.20, 0.052, TEAL),
    (0.18, 0.50, 0.045, GOLD),
    (0.83, 0.50, 0.048, MOSS),
    (0.40, 0.28, 0.036, WHITE),
    (0.62, 0.78, 0.032, WHITE),
]


def draw_icon(size=SIZE):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = int(size * 0.03)
    d.ellipse([m, m, size - m, size - m], fill=NAVY)
    d.ellipse([m, m, size - m, size - m], outline=DEEP, width=int(size * 0.02))

    random.seed(915)
    for cx, cy, r, color in BLOBS:
        x, y, rr = cx * size, cy * size, r * size
        # 橢圓一點，比較像一個「詞」而不是一顆球
        d.ellipse([x - rr * 1.35, y - rr * 0.72, x + rr * 1.35, y + rr * 0.72], fill=color)

    return img


def main():
    img = draw_icon()
    png = os.path.join(HERE, "icon.png")
    img.resize((512, 512), Image.LANCZOS).save(png)
    ico = os.path.join(HERE, "icon.ico")
    img.save(ico, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64),
                         (128, 128), (256, 256)])
    print("已產生：", png)
    print("已產生：", ico)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
