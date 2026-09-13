# -*- coding: utf-8 -*-
"""
用 PIL 畫出程式圖示（純幾何圖形，不使用任何字型，避免字型授權問題）。
執行：  python make_icon.py
產出：  icon.png（512×512）、icon.ico（16/32/48/64/128/256）
圖案：深藍圓角底 ＋ 白色投影片卡 ＋ 青藍標題列 ＋ 三根長短不一的進度條 ＋ 金色圓點（本週重點）。
"""
import os
from PIL import Image, ImageDraw

NAVY = (0x0B, 0x1F, 0x3A, 255)
TEAL = (0x1C, 0x72, 0x93, 255)
GREEN = (0x2E, 0x7D, 0x5B, 255)
GOLD = (0xC9, 0x9A, 0x3E, 255)
WHITE = (0xFF, 0xFF, 0xFF, 255)
GREY = (0xC8, 0xD4, 0xDA, 255)
BASE = os.path.dirname(os.path.abspath(__file__))


def draw_icon(size=512):
    s = size
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22), fill=NAVY)
    # 右上角裝飾圓（與簡報封面同一套語彙）
    d.ellipse([int(s * 0.62), int(-s * 0.18), int(s * 1.18), int(s * 0.38)],
              fill=(0x10, 0x2C, 0x4E, 255))
    # 白色投影片卡
    x0, y0, x1, y1 = int(s * 0.16), int(s * 0.22), int(s * 0.84), int(s * 0.78)
    d.rounded_rectangle([x0, y0, x1, y1], radius=int(s * 0.05), fill=WHITE)
    # 標題列
    d.rounded_rectangle([x0 + int(s * 0.06), y0 + int(s * 0.08),
                         x0 + int(s * 0.38), y0 + int(s * 0.12)],
                        radius=int(s * 0.02), fill=TEAL)
    # 三根進度條（長短不一 = 各區塊進度）
    bar_y = y0 + int(s * 0.20)
    for w, c in ((0.52, TEAL), (0.40, GREEN), (0.22, GREY)):
        d.rounded_rectangle([x0 + int(s * 0.06), bar_y,
                             x0 + int(s * 0.06) + int(s * w), bar_y + int(s * 0.045)],
                            radius=int(s * 0.022), fill=c)
        bar_y += int(s * 0.085)
    # 金色圓點（本週重點）
    r = int(s * 0.055)
    cx, cy = x1 - int(s * 0.10), y1 - int(s * 0.10)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=GOLD)
    return im


def main():
    im = draw_icon(512)
    png = os.path.join(BASE, "icon.png")
    ico = os.path.join(BASE, "icon.ico")
    im.save(png, "PNG")
    im.save(ico, "ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("icon.png / icon.ico 已產生：", BASE)


if __name__ == "__main__":
    main()
