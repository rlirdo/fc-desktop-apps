# -*- coding: utf-8 -*-
"""官方拼音 ↔ 語音拼音比對（純函式、離線）。

用途：老師從客委會翻譯網頁一起貼回「官方拼音」，工具就能自動檢查合成語音有沒有唸錯字。
- 第一層「音節比對」（所有腔調）：只比字母，用 difflib.SequenceMatcher 對齊。
- 第二層「聲調提示」（只有四縣腔）：字母相同時，把官方聲調符號換成數字再比。
- 對回漢字：客語漢字（去標點空白）逐字編號；字數＝兩邊音節數時才寫「第 N 字『X』」。
這是比對規則，不會修改任何文字或語音。
"""
import difflib
import re

# 官方拼音：字母＋非字母尾巴（聲調符號 ˊ ˇ ˋ ˆ ^ + 等）；以空白與標點切開
_PUNCT = r"\s,.;:!?，。、；：！？「」『』（）()\"'…—\-"
_OFF_RE = re.compile(r"[A-Za-z]+[^A-Za-z%s]*" % _PUNCT)
# 語音拼音：字母＋數字
_VOICE_RE = re.compile(r"([A-Za-z]+)(\d*)")
# 漢字比對時要去掉的標點與空白
_HAN_STRIP = re.compile(r"[\s,.;:!?，。、；：！？「」『』（）()\"'…—\-~～]")

# 四縣腔：官方符號 → 語音調值
SIXIAN_TONES = {"ˊ": "24", "ˇ": "11", "ˋ": "31", "": "55"}
SIXIAN_CHECKED = {"ˋ": "2", "": "5"}          # 入聲（以 b、d、g 結尾；-ng 除外）


def split_official(s):
    """→ [(字母小寫, 尾巴, 原樣)]"""
    out = []
    for m in _OFF_RE.finditer(s or ""):
        tok = m.group(0)
        letters = re.match(r"[A-Za-z]+", tok).group(0)
        out.append((letters.lower(), tok[len(letters):], tok))
    return out


def split_voice(s):
    """→ [(字母小寫, 數字, 原樣)]"""
    return [(m.group(1).lower(), m.group(2), m.group(0)) for m in _VOICE_RE.finditer(s or "")]


def han_chars(s):
    """去標點與空白後的逐字清單（U+FFFF 以上的字也算一個字）。"""
    return list(_HAN_STRIP.sub("", s or ""))


def sixian_tone(letters, tail):
    """官方尾巴 → 四縣調值；無法判定（未知符號）回傳 None。"""
    t = "".join(ch for ch in tail if ch in "ˊˇˋ")
    rest = "".join(ch for ch in tail if ch not in "ˊˇˋ")
    if rest.strip() or len(t) > 1:
        return None
    if letters[-1:] in ("b", "d", "g") and not letters.endswith("ng"):   # -ng 是鼻音韻尾，不是入聲
        return SIXIAN_CHECKED.get(t)
    return SIXIAN_TONES.get(t)


def compare(official, voice, hakka="", accent="sixian"):
    """回傳 dict：
    provided, n_official, n_voice, n_chars, char_mapped, tone_checked,
    syllable_issues [{kind, pos, char, official, voice}], tone_hints [...], lines [str], ok
    kind：不同／缺少／多出；pos 為 1 起算（以官方音節位置為主，多出時以語音位置）。
    """
    res = {"provided": bool((official or "").strip()), "syllable_issues": [], "tone_hints": [],
           "tone_checked": accent == "sixian", "lines": [], "ok": True}
    off = split_official(official)
    voi = split_voice(voice)
    chars = han_chars(hakka)
    res.update(n_official=len(off), n_voice=len(voi), n_chars=len(chars))
    if not res["provided"]:
        res["lines"] = ["沒有貼官方拼音，無法自動檢查；請老師聽一遍確認"]
        return res
    mapped = len(chars) == len(off) == len(voi) and len(off) > 0
    res["char_mapped"] = mapped

    def ch(i):
        return chars[i] if mapped and 0 <= i < len(chars) else ""

    sm = difflib.SequenceMatcher(a=[o[0] for o in off], b=[v[0] for v in voi], autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            if res["tone_checked"]:
                for k in range(i2 - i1):
                    o, v = off[i1 + k], voi[j1 + k]
                    want = sixian_tone(o[0], o[1])
                    if want is not None and v[1] and want != v[1]:
                        res["tone_hints"].append({"pos": i1 + k + 1, "char": ch(i1 + k),
                                                  "official": o[2], "voice": v[2],
                                                  "official_tone": want})
            continue
        if tag == "replace" and (i2 - i1) == (j2 - j1):
            for k in range(i2 - i1):
                res["syllable_issues"].append({"kind": "不同", "pos": i1 + k + 1, "char": ch(i1 + k),
                                               "official": off[i1 + k][2], "voice": voi[j1 + k][2]})
            continue
        for k in range(i1, i2):
            res["syllable_issues"].append({"kind": "缺少" if tag == "delete" else "不同",
                                           "pos": k + 1, "char": ch(k),
                                           "official": off[k][2], "voice": ""})
        for k in range(j1, j2):
            res["syllable_issues"].append({"kind": "多出" if tag == "insert" else "不同",
                                           "pos": k + 1, "char": "", "official": "",
                                           "voice": voi[k][2]})
    res["ok"] = not res["syllable_issues"]
    res["lines"] = report_lines(res)
    return res


def _where(it, mapped):
    if mapped and it["char"]:
        return "第 %d 字『%s』" % (it["pos"], it["char"])
    return "第 %d 個音節" % it["pos"]


def report_lines(res):
    if not res["provided"]:
        return ["沒有貼官方拼音，無法自動檢查；請老師聽一遍確認"]
    mapped = res.get("char_mapped")
    lines = []
    if res["syllable_issues"]:
        lines.append("✗ 有 %d 處音節和官方拼音不同：" % len(res["syllable_issues"]))
        for it in res["syllable_issues"]:
            if it["kind"] == "缺少":
                lines.append("  %s：官方 %s／語音缺少（缺少）" % (_where(it, mapped), it["official"]))
            elif it["kind"] == "多出":
                lines.append("  語音第 %d 個音節 %s：官方沒有（多出）" % (it["pos"], it["voice"]))
            else:
                lines.append("  %s：官方 %s／語音 %s" % (_where(it, mapped), it["official"] or "（無對應）", it["voice"] or "（無對應）"))
        lines.append("這幾個字語音可能唸錯。建議這一句改用客委會語音合成網頁產生。")
    else:
        head = "✓ %d 個音節全部一致" % res["n_official"]
        if res["tone_checked"] and not res["tone_hints"]:
            head += "，聲調也一致"
        lines.append(head)
    if res["tone_checked"]:
        if res["tone_hints"]:
            lines.append("聲調不同（提示）：")
            for it in res["tone_hints"]:
                lines.append("  %s：官方 %s（%s）／語音 %s" % (_where(it, mapped), it["official"],
                                                        it["official_tone"], it["voice"]))
    else:
        lines.append("此腔調只比對音節，不比聲調")
    if not mapped:
        lines.append("字數與音節數對不上，請人工核對（漢字 %d 字、官方 %d 個音節、語音 %d 個音節）"
                     % (res["n_chars"], res["n_official"], res["n_voice"]))
    return lines


def astral_positions(s):
    """回傳 U+FFFF 以上字元在「去標點空白後」的位置（1 起算）與字元。"""
    return [(i + 1, c) for i, c in enumerate(han_chars(s)) if ord(c) > 0xFFFF]
