# -*- coding: utf-8 -*-
"""合成測試資料（零真實個資，全部為虛構人名與虛構學號）。

供 `app.py --selftest` 使用，也用來產生 `selftest\\` 資料夾裡的示範檔。
資料格式仿照 Zuvio「下載數據」xlsx（多個表頭區塊、匿名作答者、未作答名單、
測驗題目型答案欄、自由文字提到同學姓名），以及名冊型 xlsx（含電子郵件欄、
含 None 表頭）與通用 CSV。
"""
import csv

import openpyxl

BLANK = None

# --------------------------------------------------------------------------
# 1) Zuvio 題組問答＋測驗題目 合成檔（5 個表頭區塊）
# --------------------------------------------------------------------------
ZUVIO_SHEET = "sheet1"

# 每一列就是實際要寫進去的內容；None = 空白儲存格
ZUVIO_ROWS = [
    ["資料夾名稱", "合成測試資料夾", BLANK, BLANK, BLANK, BLANK],                 # R1
    ["問題題型", "題組問答", BLANK, BLANK, BLANK, BLANK],                         # R2
    ["是否分組", "否", BLANK, BLANK, BLANK, BLANK],                               # R3
    ["是否匿名", "否", BLANK, BLANK, BLANK, BLANK],                               # R4
    ["題幹", "題組問答_合成資料_不含真實個資", BLANK, BLANK, BLANK, BLANK],        # R5
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R6
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R7
    ["問題類型", BLANK, BLANK, "第1題:問答題", "第2題:單選題", BLANK],             # R8
    ["問題敘述", BLANK, BLANK, "請描述一個與課程相關的生活例子。", "請選一個選項。", BLANK],  # R9
    ["正解", BLANK, BLANK, "無正解", "無正解", BLANK],                             # R10
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R11
    ["學號", "姓名", "作答時間", "回答", "回答", BLANK],                           # R12  ← 區塊 1
    ["匿名作答者", "匿名作答者", "2026-09-09 15:00:00", "這是匿名的回答", "(2)乙", BLANK],   # R13
    [411354043, "王小明", "2026-09-07 16:00:00", "我和林大同一起討論了這題", "(1)甲", BLANK],  # R14  int 學號
    [411354044, "林大同", "2026-09-08 10:00:00", "參考了陳曉華學姊的報告", "(3)丙", BLANK],    # R15  int 學號
    ["411354045", "陳曉華", "2026-09-08 11:00:00", "自己查資料完成", "(2)乙", BLANK],          # R16  str 學號
    [411354046, "歐陽小花", "2026-09-09 09:00:00", "和王小明討論過", "(1)甲", BLANK],          # R17
    [411354047, "李明", "2026-09-10 08:00:00", "沒有跟別人討論", "(2)乙", BLANK],              # R18
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R19
    ["未作答學生", BLANK, BLANK, BLANK, BLANK, BLANK],                            # R20
    ["學號", "姓名", BLANK, BLANK, BLANK, BLANK],                                 # R21  ← 區塊 2
    [411354048, "張三豐", BLANK, BLANK, BLANK, BLANK],                            # R22
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R23
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R24
    ["第1題:問答題", BLANK, BLANK, BLANK, BLANK, BLANK],                          # R25
    ["問題敘述", "請描述一個與課程相關的生活例子。", BLANK, BLANK, BLANK, BLANK],   # R26
    ["學號", "姓名", "回答", BLANK, BLANK, BLANK],                                # R27  ← 區塊 3
    [411354043, "王小明", "我和林大同一起討論了這題", BLANK, BLANK, BLANK],        # R28
    [411354044, "林大同", "參考了陳曉華學姊的報告", BLANK, BLANK, BLANK],          # R29
    [411354046, "歐陽小花", "和王小明討論過", BLANK, BLANK, BLANK],                # R30
    ["匿名作答者", "匿名作答者", "這是匿名的回答", BLANK, BLANK, BLANK],           # R31
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R32
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R33
    ["測驗題目", "合成小考", BLANK, BLANK, BLANK, BLANK],                         # R34
    ["學號", "姓名", "作答時間", "總分", "第1題", "第2題"],                        # R35  ← 區塊 4
    [411354043, "王小明", "2026-09-12 09:00:00", 80, "答對", "答錯"],              # R36
    ["411354045", "陳曉華", "2026-09-12 09:05:00", 100, "答對", "答對"],           # R37
    [411354049, "林明", "2026-09-12 09:10:00", 60, "答錯", "答對"],                # R38  2 字姓名
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R39
    [BLANK, BLANK, BLANK, BLANK, BLANK, BLANK],                                   # R40
    ["分組討論彙總", BLANK, BLANK, BLANK, BLANK, BLANK],                          # R41
    ["學號", "姓名", "回答", "同組成員", BLANK, BLANK],                            # R42  ← 區塊 5
    [411354049, "林明", "討論主題是水質檢測", "李明", BLANK, BLANK],               # R43
    [411354043, "王小明", "我和歐陽小花同組", "林明", BLANK, BLANK],               # R44
]

# ---- 期望值（selftest 逐項比對用） ---------------------------------------
ZUVIO_EXPECT = {
    "blocks": 5,
    "students": 7,
    "anon_rows": 2,
    "id_masked": 16,
    "email_masked": 0,
    "name_masked": 14,
    "text_replacements": 9,
    "link_rows": 7,
}

# 學生編號（依首次出現順序）→ 對應的合成姓名／首次出現列
ZUVIO_ORDER = [
    (1, "王小明", "411354043", 14),
    (2, "林大同", "411354044", 15),
    (3, "陳曉華", "411354045", 16),
    (4, "歐陽小花", "411354046", 17),
    (5, "李明", "411354047", 18),
    (6, "張三豐", "411354048", 22),
    (7, "林明", "411354049", 38),
]

# 表頭列（原始列號，插欄後列號不變）
ZUVIO_HEADER_ROWS = [12, 21, 27, 35, 42]

# (列, 期望的 A 欄值)：n → 第 n 號；"匿名"；"" 代表空白
ZUVIO_A_COL = {
    12: "學生編號", 13: "匿名", 14: 1, 15: 2, 16: 3, 17: 4, 18: 5,
    21: "學生編號", 22: 6,
    27: "學生編號", 28: 1, 29: 2, 30: 4, 31: "匿名",
    35: "學生編號", 36: 1, 37: 3, 38: 7,
    42: "學生編號", 43: 7, 44: 1,
}

# 姓名遮罩後（B 欄以後皆右移 1 欄，姓名欄 = 原 2 → 3）
ZUVIO_MASKED_NAMES = {
    14: "王O明", 15: "林O同", 16: "陳O華", 17: "歐O小花", 18: "李O",
    22: "張O豐", 28: "王O明", 29: "林O同", 30: "歐O小花",
    36: "王O明", 37: "陳O華", 38: "林O", 43: "林O", 44: "王O明",
    13: "匿名作答者", 31: "匿名作答者",     # 匿名列不遮罩
}

# 文字替換後的期望內容（列, 插欄後的欄索引）→ 期望字串（{n} 代表第 n 號學生編號）
ZUVIO_TEXT = {
    (14, 5): "我和{2}一起討論了這題",
    (15, 5): "參考了{3}學姊的報告",
    (17, 5): "和{1}討論過",
    (28, 4): "我和{2}一起討論了這題",
    (29, 4): "參考了{3}學姊的報告",
    (30, 4): "和{1}討論過",
    (43, 4): "討論主題是水質檢測",      # 不含姓名，不可被動到
    (43, 5): "{5}",                     # 同組成員「李明」整段中文＝姓名 → 換編號
    (44, 4): "我和{4}同組",
    (44, 5): "{7}",
    (36, 6): "答對",                    # 測驗答案欄不可被當成姓名
    (36, 7): "答錯",
    (38, 6): "答錯",
    (13, 5): "這是匿名的回答",
}


def write_zuvio_sample(path):
    """產生 Zuvio 格式合成檔，回傳路徑。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ZUVIO_SHEET
    for row in ZUVIO_ROWS:
        ws.append(list(row))
    wb.save(path)
    wb.close()
    return path


# --------------------------------------------------------------------------
# 2) 名冊型合成檔（第 1 列就是表頭、含電子郵件欄、含 None 表頭）
# --------------------------------------------------------------------------
ROSTER_SHEET = "名冊"
ROSTER_ROWS = [
    [BLANK, "學號", "姓名", "電子郵件", BLANK],                                    # R1（第 1、5 欄表頭是 None）
    [BLANK, 411354043, "王小明", "wang.ming@gms.ndhu.edu.tw", "環境化學修課生甲班A"],
    [BLANK, 411354044, "林大同", "lin.datong@gms.ndhu.edu.tw", "環境化學修課生乙班B"],
    [BLANK, "411354045", "陳曉華", "chen.hsiaohua@gms.ndhu.edu.tw", "環境化學修課生丙班C"],
]
ROSTER_EXPECT = {
    "blocks": 1,
    "students": 3,
    "id_masked": 3,
    "email_masked": 3,
    "name_masked": 3,
    "text_replacements": 0,
    "link_rows": 3,
}


def write_roster_sample(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ROSTER_SHEET
    for row in ROSTER_ROWS:
        ws.append(list(row))
    wb.save(path)
    wb.close()
    return path


# --------------------------------------------------------------------------
# 3) 通用 CSV
# --------------------------------------------------------------------------
CSV_ROWS = [
    ["姓名", "學號", "作答"],
    ["王小明", "411354043", "我和林大同一起做實驗"],
    ["林大同", "411354044", "自己完成"],
]
CSV_EXPECT = {"blocks": 1, "students": 2, "id_masked": 2, "text_replacements": 1, "link_rows": 2}


def write_csv_sample(path):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(CSV_ROWS)
    return path


if __name__ == "__main__":
    import os
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "selftest")
    os.makedirs(folder, exist_ok=True)
    a = write_zuvio_sample(os.path.join(folder, "樣本_Zuvio合成資料.xlsx"))
    b = write_roster_sample(os.path.join(folder, "樣本_名冊合成資料.xlsx"))
    c = write_csv_sample(os.path.join(folder, "樣本_通用格式.csv"))
    for p in (a, b, c):
        print("已產生合成測試檔：" + p)
