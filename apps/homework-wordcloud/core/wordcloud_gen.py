# -*- coding: utf-8 -*-
"""
wordcloud_gen.py — 每題產一張文字雲 PNG。

規格：
  - 字型：Windows 微軟正黑體（msjhbd.ttc → msjh.ttc）；macOS PingFang／STHeiti；
          Linux Noto Sans CJK；都找不到再退回 wordcloud 內建字型（中文會缺字，會寫警告）。
          ※ 一律用「系統既有字型」，不打包微軟正黑體（授權限制）。
  - 頻率＝該題全班答案詞頻（高 → 字愈大）
  - 配色：綠色化學色票（navy / deep / teal / green / moss / gold），白底，頻率高者用深色
  - 輸出：<out_dir>/Q0N_wc.png，1600×900
"""
import os
import re
import csv
import glob
import sys
from collections import Counter

import wordcloud as _wc_pkg
from wordcloud import WordCloud

from . import analyze as AN

WIN_FONTS = [
    r"C:\Windows\Fonts\msjhbd.ttc",
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjh.ttf",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\NotoSansTC-Regular.otf",
    r"C:\Windows\Fonts\NotoSansCJKtc-Regular.otf",
    r"C:\Windows\Fonts\mingliu.ttc",
    r"C:\Windows\Fonts\kaiu.ttf",
    r"C:\Windows\Fonts\simhei.ttf",
]

MAC_FONTS = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
]

LINUX_FONTS = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
]

PALETTE = ["#0B1F3A", "#065A82", "#1C7293", "#2E7D5B", "#5FA98A", "#C99A3E"]
W, H = 1600, 900

FONT_HELP = (
    "找不到系統中文字型，文字雲改用內建英文字型，中文會變成空格或方框。\n"
    "  Windows：請確認 C:\\Windows\\Fonts\\msjh.ttc（微軟正黑體）存在。\n"
    "  macOS：請確認 /System/Library/Fonts/PingFang.ttc 存在。\n"
    "  其他：可到 Google Fonts 下載 Noto Sans TC 安裝後重跑。"
)

_FONT_CACHE = None          # (路徑, 是否支援中文)


def candidates():
    if sys.platform.startswith("win"):
        return WIN_FONTS
    if sys.platform == "darwin":
        return MAC_FONTS
    return LINUX_FONTS


def builtin_font():
    """wordcloud 套件自帶的 DroidSansMono.ttf（沒有中文字，只作最後保險）。"""
    return os.path.join(os.path.dirname(os.path.abspath(_wc_pkg.__file__)),
                        "DroidSansMono.ttf")


def font_path(log=print):
    """回傳可用字型路徑；找不到中文字型時退回內建字型並寫警告，不丟例外。"""
    global _FONT_CACHE
    if _FONT_CACHE is None:
        picked = ""
        for p in candidates():
            if os.path.exists(p):
                picked = p
                break
        if picked:
            _FONT_CACHE = (picked, True)
        else:
            _FONT_CACHE = (builtin_font(), False)
            for line in ("  [警告] " + FONT_HELP).splitlines():
                log(line)
    return _FONT_CACHE[0]


def make_color_func(freqs):
    """頻率愈高 → 色票愈深（navy 端）。"""
    if not freqs:
        return lambda *a, **k: PALETTE[0]
    mx, mn = max(freqs.values()), min(freqs.values())
    span = max(mx - mn, 1)

    def f(word, **kw):
        v = freqs.get(word, mn)
        i = int((1 - (v - mn) / span) * (len(PALETTE) - 1))
        return PALETTE[min(max(i, 0), len(PALETTE) - 1)]
    return f


def build(freqs, out_png, log=print):
    if not freqs:
        return False
    wc = WordCloud(
        font_path=font_path(log),
        width=W, height=H,
        background_color="white",
        prefer_horizontal=0.92,
        max_words=90,
        min_font_size=16,
        max_font_size=260,
        relative_scaling=0.45,
        collocations=False,
        margin=6,
        random_state=42,
    ).generate_from_frequencies(freqs)
    wc.recolor(color_func=make_color_func(freqs), random_state=42)
    wc.to_file(out_png)
    return True


def generate_all(work_dir, out_dir=None, min_words=5, log=print):
    """對 work_dir 內每個 Q*.csv 產一張文字雲，回傳 {題號: png 路徑}。"""
    out_dir = out_dir or os.path.join(work_dir, "wc")
    os.makedirs(out_dir, exist_ok=True)
    made = {}
    for csv_path in sorted(glob.glob(os.path.join(work_dir, "Q*.csv"))):
        qno = os.path.basename(csv_path)[:3]
        freqs = Counter()
        with open(csv_path, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if not AN.is_open_text(r["子題號"]):
                    continue                      # 選擇題的回答是選項文字，不進文字雲
                t = AN.clean(r["作答內容"])
                if not AN.is_valid(t):
                    continue
                freqs.update(AN.tokens(t))
        freqs = {w: c for w, c in freqs.items()
                 if len(w) >= 2 or re.match(r"^[A-Za-z]{2,}$", w)}
        if len(freqs) < min_words:
            log(f"  {qno} 詞數 {len(freqs)} 不足 {min_words}，略過文字雲（該頁改用純排版）")
            continue
        png = os.path.join(out_dir, f"{qno}_wc.png")
        build(freqs, png, log=log)
        made[qno] = png
        log(f"  {qno} -> wc\\{os.path.basename(png)}（{len(freqs)} 個詞）")
    return made
