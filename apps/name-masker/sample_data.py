# -*- coding: utf-8 -*-
"""合成測試資料（零真實個資，全部為虛構人名）。

供 --selftest 使用，也用來產生 selftest/ 資料夾內的示範檔。
"""
import openpyxl

SHEET_TITLE = "作答名單"

# (序號, 姓名備註, 姓名, 分數)
ROWS = [
    (1, "甲班同學", "王小明", 90),
    (2, "乙班同學", "歐陽小花", 85),
    (3, "丙班同學", "李明", 78),
    (4, "外籍生", "Mary Chen", 88),
    (5, "未填", None, 60),
    (6, "系統預設", "匿名作答者", 55),
    (7, "以學號代替", 12345, 70),
]

HEADERS = ("序號", "姓名備註", "姓名", "分數")

# 第 3 欄（C 欄＝姓名）遮罩後的期望值；None 代表該格應維持空白
EXPECTED_NAMES = [
    "王O明",
    "歐O小花",
    "李O",
    "Mary Chen",
    None,
    "匿O作答者",
    12345,
]

EXPECTED_MASKED_COUNT = 4
NAME_COL = 3          # C 欄
NOTE_COL = 2          # B 欄（姓名備註，不可被遮罩）


def write_sample(path):
    """產生合成測試用的 .xlsx，回傳路徑。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_TITLE
    ws.append(list(HEADERS))
    for r in ROWS:
        ws.append(list(r))
    wb.save(path)
    return path


if __name__ == "__main__":
    import os
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "selftest", "樣本_合成資料.xlsx")
    os.makedirs(os.path.dirname(target), exist_ok=True)
    write_sample(target)
    print("已產生合成測試檔：" + target)
