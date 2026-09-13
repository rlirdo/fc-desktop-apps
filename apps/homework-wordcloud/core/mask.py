# -*- coding: utf-8 -*-
"""
mask.py — 姓名遮罩（隱私鐵律）

規則（與 excel-name-masker\\mask_names.py 相同）：
    把中文姓名的第 2 個字改成 O
        王小明 -> 王O明
        李明   -> 李O
        歐陽小花 -> 歐O小花

兩道防線：
    1) mask_name()    ：「姓名」欄位一律遮罩。
    2) mask_in_text() ：自由文字（作答內容）裡若出現名冊上的真實姓名，也一併換掉。
                        自我介紹類題目很常把名字寫進答案，這道防線不可省。

任何寫到硬碟、簡報、網路的資料都必須是遮罩後的版本；
原始 xlsx 只留在 input\\，永遠不進 git、不上傳。
"""
import re

# 連續 2 個以上中文字（含擴充 A 區），用來辨識「這一段是姓名」
NAME_RUN = re.compile("[一-鿿㐀-䶿]{2,}")

# 系統佔位字串，不是真人姓名，不遮罩
NOT_A_NAME = {"匿名作答者", "未作答學生", "匿名", "無", "測試帳號", "示範帳號"}


def mask_name(text):
    """把字串中第一段連續中文字的第 2 個字改為 O；無中文則原樣返回。"""
    if not isinstance(text, str) or not text:
        return text
    if text.strip() in NOT_A_NAME:
        return text
    m = NAME_RUN.search(text)
    if not m:
        return text
    run = m.group(0)
    masked = run[0] + "O" + run[2:]
    return text[: m.start()] + masked + text[m.start() + len(run):]


def mask_in_text(ans, roster):
    """把名冊中的真實姓名在自由文字裡換成遮罩版（長名優先，避免部分覆蓋）。"""
    if not isinstance(ans, str) or not ans or not roster:
        return ans
    for real, fake in roster:
        if real and real in ans:
            ans = ans.replace(real, fake)
    return ans


def build_roster(*name_lists):
    """回傳 [(真實姓名, 遮罩姓名)]，依長度由長到短排序。"""
    names = set()
    for lst in name_lists:
        for n in lst:
            n = (n or "").strip()
            if len(n) >= 2 and n not in NOT_A_NAME and NAME_RUN.fullmatch(n):
                names.add(n)
    return sorted(((n, mask_name(n)) for n in names), key=lambda t: -len(t[0]))


def roster_names(roster):
    """只取真實姓名清單（供 --check-privacy 掃描用）。"""
    return [real for real, _ in roster]
