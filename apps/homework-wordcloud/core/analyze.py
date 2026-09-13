# -*- coding: utf-8 -*-
"""
analyze.py — 吃某週資料夾內的 Q*.csv，產出 analysis.json ＋ summary.md

流程：
  1. 清理：去掉網址、空白、純符號、以及「無／未作答」這類佔位答案。
  2. 斷詞：jieba ＋ 課程自訂詞（config.json 的 user_words）＋ 停用詞表。
  3. 詞頻：每題 TOP-N 關鍵詞、每個子題 TOP-N、全班 TOP10。
  4. 統計：每題作答人數／全班人數／作答率／有效列數／平均字數。
  5. 常見問題：答案裡帶問號或「不懂／為什麼」等字樣的句子。

所有輸入都已經是遮罩版 CSV，這裡不再碰任何真實姓名。
"""
import os
import re
import csv
import json
import glob
from collections import Counter, defaultdict

import jieba

# ---------------------------------------------------------------- 內建詞典
# 通用版只放「理工／環境課程常見」的基本詞，其餘請寫進 config.json 的 user_words。
BASE_USER_WORDS = [
    "綠色化學", "十二原則", "綠色化學十二原則", "原子經濟", "原子經濟性", "綠色溶劑",
    "催化劑", "再生原料", "可分解", "即時分析", "廢棄物", "減量", "廢棄物減量",
    "能源效率", "本質安全", "衍生物", "循環經濟", "資源回收", "減塑", "永續",
    "永續發展", "淨零碳排", "碳循環", "碳足跡", "溫室氣體", "二氧化碳", "溫室效應",
    "氣候變遷", "全球暖化", "再生能源", "太陽能", "水力發電", "火力發電", "綠能",
    "水化學", "水污染", "水質", "廢水", "硬度", "溶氧", "優養化", "混凝", "絮凝",
    "重金屬", "微塑膠", "塑膠微粒", "生物累積", "生物放大", "食物鏈", "污染物",
    "毒性", "低毒性", "無毒", "空氣污染", "土壤污染", "酸雨", "分配係數",
    "氧化還原", "官能基", "化學鍵", "共價鍵", "離子鍵", "莫耳", "prompt",
    "生成式AI", "人工智慧", "提示詞", "查證", "幻覺", "文獻回顧",
    "ChatGPT", "Gemini", "Copilot", "Claude", "NotebookLM", "Perplexity",
    "學習金字塔", "主動學習", "被動學習", "小組討論", "動手做",
]

BASE_STOPWORDS = set("""
的 了 是 在 我 我們 你 你們 他 她 它 這 那 有 和 與 就 也 都 而 及 或 等 很 更 最 會 能 可 可以
覺得 認為 應該 因為 所以 但是 但 不過 然後 還有 還是 一個 一些 什麼 怎麼 為什麼 如何
自己 自我 沒有 不是 不會 這個 那個 其實 真的 比較 非常 可能 需要 已經 透過 由於 對於
以及 之後 之前 時候 東西 事情 方面 部分 問題 方式 方法 感覺 知道 了解 學習 課程 老師
上課 同學 內容 使用 進行 讓 把 被 從 到 向 於 之 其 此 該 個 們 吧 呢 啊 喔 嗎 呀 哦
a b c d e f 1 2 3 4 5 6 7 8 9 0 無 沒 不 要 想 做 看 用 說 好 多 少 大 小 中 上 下 來 去
無正解 未作答 空白 以上 例如 像是 尤其 甚至 只是 而且 如果 雖然 因此 這些 那些 一樣 一起
今天 現在 目前 之類 等等 這樣 那樣 一下 有點 覺 得 我的 他的 她的 它的 自身 本身
我用 我會 我想 我請 我跟 我和 我在 我有 讓我 他們 以下 一項 這一 那一 其他 另外
""".split())

# 學生常把題目代號寫在答案開頭（1a、C2、3.、(1)…），這些 token 會污染文字雲
NUMBERING = re.compile(r"^[0-9A-Za-z.]{1,4}$")
BASE_KEEP_SHORT = {"AI", "pH", "CO2", "UV", "DNA", "6R", "CFC", "HFC", "GPT", "PM"}

# 大小寫歸一：ai / Ai / chatgpt 併成同一個詞再計頻
CANON = {"ai": "AI", "Ai": "AI", "ph": "pH", "PH": "pH", "6r": "6R", "co2": "CO2",
         "chatgpt": "ChatGPT", "Chatgpt": "ChatGPT", "gemini": "Gemini",
         "claude": "Claude", "notebooklm": "NotebookLM", "copilot": "Copilot"}

PUNCT_ONLY = re.compile(r"^[\s\W_]*$", re.UNICODE)
URL = re.compile(r"https?://\S+")
PLACEHOLDER = {"無", "沒有", "未作答", "無正解", "空白", "已作答", "不知道", "略", "x", "X", "-"}

QUESTION_MARK = ["?", "？", "不懂", "不太懂", "為什麼", "不確定", "不知道", "想知道",
                 "好奇", "困惑", "不清楚", "是不是", "會不會", "怎麼辦"]

# 選擇題的「回答」其實是選項文字，全班都一樣，放進文字雲只會洗版。
# 因此文字雲與詞頻只吃開放文字題；選擇題改成統計選項分佈。
CHOICE_SUB = re.compile(r"(單選|多選|是非|排序|評分|順序)")


def is_open_text(sub_label):
    """子題號像「第2題:單選題」→ False；「第1題:問答題」→ True。"""
    return not CHOICE_SUB.search(sub_label or "")

# 模組層狀態（load_dict 會依 config 覆寫）
STOPWORDS = set(BASE_STOPWORDS)
KEEP_SHORT = set(BASE_KEEP_SHORT)
_LOADED = False


def load_dict(user_words=None, stopwords_extra=None, keep_short=None):
    """把課程自訂詞灌進 jieba，並合併停用詞／短詞白名單。可重複呼叫。"""
    global STOPWORDS, KEEP_SHORT, _LOADED
    for w in list(BASE_USER_WORDS) + list(user_words or []):
        if w:
            jieba.add_word(w, freq=100000)
    STOPWORDS = set(BASE_STOPWORDS) | set(stopwords_extra or [])
    KEEP_SHORT = set(BASE_KEEP_SHORT) | set(keep_short or [])
    _LOADED = True


def clean(t):
    return URL.sub("", str(t or "")).strip()


def is_valid(t):
    if not t or PUNCT_ONLY.match(t):
        return False
    return t.strip() not in PLACEHOLDER


def tokens(t):
    if not _LOADED:
        load_dict()
    out = []
    for w in jieba.cut(t):
        w = w.strip()
        if not w or PUNCT_ONLY.match(w):
            continue
        w = CANON.get(w, w)
        if w in STOPWORDS or w.lower() in STOPWORDS:
            continue
        if len(w) < 2 and not re.match(r"^[A-Za-z]{2,}$", w):
            continue
        if NUMBERING.match(w) and w not in KEEP_SHORT:
            continue
        out.append(w)
    return out


def freq_of(texts):
    f = Counter()
    for t in texts:
        f.update(tokens(t))
    return f


def speakers_by_word(rows):
    """{詞: 提到這個詞的同學人數}。用斷詞結果計算，才會和詞頻同一套標準。"""
    d = defaultdict(set)
    for r in rows:
        sid = r["學號"] or r["姓名"]
        for w in set(tokens(r["作答內容"])):
            d[w].add(sid)
    return {w: len(v) for w, v in d.items()}


def analyse_dir(work_dir, index=None, top_n=15, log=print):
    """讀 work_dir 下的 Q*.csv，寫出 analysis.json 與 summary.md，回傳 result dict。"""
    meta = {r["題號"]: r for r in (index or [])}
    # 只存資料夾名稱，不存完整本機路徑（輸出可能會分享出去）
    result = {"資料夾": os.path.basename(os.path.normpath(work_dir)), "題目": []}
    all_freq = Counter()
    all_questions = []

    for path in sorted(glob.glob(os.path.join(work_dir, "Q*.csv"))):
        qno = os.path.basename(path)[:3]
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        valid_rows, texts = [], []
        per_sub = defaultdict(list)
        students = set()
        for r in rows:
            t = clean(r["作答內容"])
            if not is_valid(t):
                continue
            r = dict(r, 作答內容=t)
            per_sub[(r["子題號"], r["子題題目"])].append(r)
            students.add(r["學號"] or r["姓名"])
            if is_open_text(r["子題號"]):     # 選擇題不進文字雲／詞頻
                valid_rows.append(r)
                texts.append(t)

        freq = freq_of(texts)
        all_freq.update(freq)
        questions = [t[:60] for t in texts if any(k in t for k in QUESTION_MARK)]
        all_questions += questions

        subs = []
        for (label, text), items in per_sub.items():
            itexts = [r["作答內容"] for r in items]
            sub = {
                "子題號": label,
                "子題題目": text,
                "型別": "開放文字" if is_open_text(label) else "選擇題",
                "有效作答數": len(items),
                "作答人數": len({r["學號"] or r["姓名"] for r in items}),
                "平均字數": round(sum(len(t) for t in itexts) / max(len(itexts), 1), 1),
            }
            if is_open_text(label):
                sub["TOP詞"] = freq_of(itexts).most_common(12)
            else:
                c = Counter(itexts)
                sub["TOP詞"] = []
                sub["選項分佈"] = [{"選項": o, "人數": k,
                                 "百分比": round(k * 100.0 / max(len(itexts), 1), 1)}
                                for o, k in c.most_common()]
            subs.append(sub)

        m = meta.get(qno, {})
        answered = m.get("作答人數") or len(students)
        klass = m.get("全班人數") or len(students)
        spk = speakers_by_word(valid_rows)
        top_rows = [{"詞": w, "次數": c, "提及人數": spk.get(w, 0)}
                    for w, c in freq.most_common(top_n)]
        result["題目"].append({
            "題號": qno,
            "題目": m.get("題目", qno),
            "題型": m.get("題型", ""),
            "來源檔": m.get("來源檔", ""),
            "作答人數": answered,
            "全班人數": klass,
            "作答率": round(answered * 100.0 / max(klass, 1), 1),
            "原始列數": len(rows),
            "有效列數": sum(len(v) for v in per_sub.values()),
            "文字列數": len(texts),
            "平均字數": round(sum(len(t) for t in texts) / max(len(texts), 1), 1),
            "總詞數": sum(freq.values()),
            "相異詞數": len(freq),
            "TOP詞": freq.most_common(top_n),
            "TOP詞明細": top_rows,
            "常見問題": questions[:5],
            "子題": subs,
        })
        log(f"  {qno} 原始 {len(rows)} 列 → 有效 {sum(len(v) for v in per_sub.values())} 列"
            f"（其中開放文字 {len(texts)} 列），子題 {len(subs)}，"
            f"TOP 詞 = {[w for w, _ in freq.most_common(5)]}")

    result["全班重點TOP10"] = all_freq.most_common(10)
    result["常見問題TOP5"] = all_questions[:5]
    return result


def write_outputs(result, work_dir, course_name="", week_label=""):
    with open(os.path.join(work_dir, "analysis.json"), "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    L = [f"# {course_name}　學生作答分析摘要{('（' + week_label + '）') if week_label else ''}", "",
         f"- 資料夾：`{result.get('資料夾', '')}`",
         f"- 題數：{len(result['題目'])}",
         "",
         "## 各題作答率", "",
         "| 題號 | 題目 | 題型 | 作答/全班 | 作答率 | 有效列數 | 平均字數 |",
         "|---|---|---|---|---|---|---|"]
    for q in result["題目"]:
        L.append(f"| {q['題號']} | {q['題目']} | {q['題型']} | "
                 f"{q['作答人數']}/{q['全班人數']} | {q['作答率']}% | "
                 f"{q['有效列數']} | {q['平均字數']} |")
    L += ["", "## 全班最常出現重點 TOP10", ""]
    for i, (w, c) in enumerate(result["全班重點TOP10"], 1):
        L.append(f"{i}. {w}（{c} 次）")
    L += ["", "## 各題重點", ""]
    for q in result["題目"]:
        L.append(f"### {q['題號']}　{q['題目']}")
        L.append("- TOP 詞：" + "、".join(f"{w}({c})" for w, c in q["TOP詞"][:8]))
        for s in q["子題"]:
            head = f"  - {s['子題號']} {s['子題題目'][:40]}　有效 {s['有效作答數']} 列；"
            if s.get("選項分佈"):
                L.append(head + "選項分佈：" +
                         "、".join(f"{o['選項']} {o['人數']}人({o['百分比']}%)"
                                  for o in s["選項分佈"]))
            else:
                L.append(head + "TOP：" + "、".join(w for w, _ in s["TOP詞"][:6]))
        L.append("")
    if result["常見問題TOP5"]:
        L += ["## 同學的疑問（節錄）", ""]
        for i, s in enumerate(result["常見問題TOP5"], 1):
            L.append(f"{i}. {s}")
    L += ["", "---", "", "> 本檔所有姓名皆已遮罩（第 2 字改 O）。原始 xlsx 僅留在本機 input\\，不得上傳。"]

    with open(os.path.join(work_dir, "summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(L))
