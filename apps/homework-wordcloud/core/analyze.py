# -*- coding: utf-8 -*-
"""
analyze.py — 吃某週資料夾內的 Q*.csv，產出每題分析。（2.2）

2.2 的改變：三個重點 → **六個重點**、兩類提問 → **四類提問**。

決定性演算法，離線可跑，不依賴任何 LLM／API：
  1. 六個重點與概念矩陣
     - 概念候選＝jieba 詞頻 TOP（沿用 user_words／stopwords）。
     - 重點＝前 6 名「提及人數」最高的概念；同分以詞頻決，再同分以字典序決。
       候選不足 6 個時**有幾個列幾個**（不補空字串、不報錯）。
     - 概念矩陣：列＝學生編號、欄＝6 個重點概念，值 1／0（含 config.json 的 synonyms），
       最後一欄「提到重點數」。輸出 `Q0N_concept_matrix.csv`。
     - 覆蓋率：每個概念＝提及人數／作答人數；整體覆蓋率＝至少提到 1 個重點的比例；
       另算「提到三個以上比例」（≥3 個重點）與「全部提到比例」（六個都提到）。
  2. 提問抽取與四類分類
     - **提問句判定（寧缺勿濫，規則與 2.1 相同）**：句尾是 ？／?，或句尾是 嗎／呢，
       或句子**開頭**就是疑問詞（為什麼／如何／怎麼／是否／什麼是／哪些…，
       英文 Why／How／What／Is／Can…）。
       句中含疑問詞但屬陳述句的**不算**，例如
       「了解物質是由原子組成，以及原子如何組成不同物質」。
     - 四類（規則式，關鍵詞可在 config.json 的 question_types 改，
       **dict 的先後順序就是判定優先序**，命中數相同時排前面的贏）：
         計算數據類：怎麼算／計算／公式／單位／數值／濃度／莫耳／pH／平衡常數／
                     幾次方／換算／多少／比例／百分比／有效數字
         操作應用類：怎麼做／如何（做、用、看、觀察、測量）／步驟／方法／實驗／
                     儀器／操作／應用／例子／處理／How
         延伸探究類：未來／影響／如果／會不會／能不能／可否／為什麼不／生活／產業／
                     政策／案例／比較好／替代／永續／環境衝擊／風險
         概念理解類：為什麼／為何／原理／定義／意義／差別／不同／關係／機制／
                     是什麼／什麼是／何者／哪一項／本質／特性／Why／What
       四類都沒中時走後備規則（含 什麼／嗎／為／Why → 概念理解類），
       再沒中才計入「其他」（目標 ≤ 兩成）。
     - 統計各類提問數、提問人數、代表句（≤40 字，已去識別化）。
       輸出 `Q0N_questions.csv`（學生編號、提問句、分類）。

所有輸入都已經是去識別化 CSV（學生編號＋遮罩姓名＋全 O 學號），這裡不再碰任何真實姓名。
"""
import os
import re
import csv
import json
import glob
from collections import Counter, defaultdict

import jieba

# ---------------------------------------------------------------- 內建詞典
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
開始 影響 減少 想到 提到 變成 用在 一定 大家 重要 有效 印象 深刻 最多 一直 地方
情形 樣子 之類 有關 相關 直接 常常 幫我 拿來 看到 聽到 發現 注意 特別 完全 到底
不要 就是 一開始 很多 有些 盡量 造成 用來 這一項 不如 出來 起來 下來 過來 很難
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

# ---------------------------------------------------------------- 提問抽取
SENT_END = "。！!？?；;\n"

# 判定「這句是不是提問」——寧可漏抓也不要把陳述句當提問。
# 反例（**不算**提問）：「了解物質是由原子組成，以及原子如何組成不同物質」
#                       「理解物質的微觀結構如何決定性質」
# 這兩句都含疑問詞，但句子是陳述句，所以只靠「句中含疑問詞」會大量誤判。
#
# 規則（三選一才算提問）：
#   1. 句尾是 ？ 或 ?
#   2. 句尾是 嗎／呢（可以沒有問號）
#   3. 句子**開頭**就是疑問詞（前面可以有「請問」「我想知道」這類引導語）
QUESTION_END = re.compile(r"[？?]\s*$")
QUESTION_TAIL = re.compile(r"(嗎|呢)[\s。.!！~～]*$")

Q_LEAD_PREFIX = r"(?:我?(?:很|想|好)?(?:請問|想問|想知道|好奇|不懂|不清楚|不確定)[，,、：:\s]*)?"
Q_WORDS = ("為什麼|為何|如何|怎麼|怎樣|怎麼樣|是否|什麼是|什麼叫|甚麼是|"
           "哪些|哪個|哪一|哪裡|可否|能不能|可不可以|有沒有|是不是|會不會")
QUESTION_LEAD = re.compile(r"^" + Q_LEAD_PREFIX + r"(?:" + Q_WORDS + r")")
QUESTION_LEAD_EN = re.compile(
    r"^(why|how|what|which|who|whom|whose|when|where|is|are|am|was|were|"
    r"can|could|do|does|did|should|would|will|shall|may|might|have|has)\b", re.I)

# 四類提問（2.2）。**dict 的先後順序就是判定優先序**：命中數一樣時，排前面的贏。
DEFAULT_QUESTION_TYPES = {
    "計算數據類": ["怎麼算", "怎樣算", "如何算", "計算", "算出", "公式", "單位", "數值", "濃度", "莫耳", "pH", "平衡常數", "幾次方", "換算", "多少", "比例", "百分比", "有效數字"],
    "操作應用類": ["怎麼做", "怎麼用", "怎麼看", "怎麼觀察", "怎麼測量", "怎樣做", "如何做", "如何用", "如何看", "如何觀察", "如何測量", "步驟", "方法", "辦法", "實驗", "儀器", "操作", "應用", "例子", "處理", "How", "如何", "怎麼", "怎樣", "方式"],
    "延伸探究類": ["未來", "影響", "如果", "會不會", "能不能", "可否", "可不可以", "為什麼不", "為何不", "生活", "產業", "政策", "案例", "比較好", "替代", "永續", "環境衝擊", "風險"],
    "概念理解類": ["為什麼", "為何", "原理", "定義", "意義", "差別", "差異", "不同", "關係", "機制", "是什麼", "什麼是", "何者", "哪一項", "哪一個", "哪項", "哪一種", "本質", "特性", "區別", "Why", "What", "哪些", "哪種", "哪個", "哪裡", "何種", "何謂"],
}
# 後備規則：已判定為提問、但四類關鍵詞都沒中的句子，含這些字就歸概念理解類，
# 其餘才真的算「其他」。（目標：其他 ≤ 總提問的兩成）
FALLBACK_CONCEPT = ["什麼", "嗎", "為", "why", "甚麼"]
FALLBACK_TYPE = "概念理解類"
OTHER_TYPE = "其他"

# 選擇題／測驗題的「回答」是選項文字或「答對／答錯」，全班都一樣，
# 放進文字雲只會洗版；因此文字雲、詞頻與概念矩陣只吃開放文字題。
CHOICE_SUB = re.compile(r"(單選|多選|是非|排序|評分|順序|測驗)")

# 測驗型子題的答案字樣（用來算答對率）
CORRECT_WORDS = {"答對", "正確", "對", "O", "o", "✓", "V", "v"}
WRONG_WORDS = {"答錯", "錯誤", "錯", "X", "x", "✗"}

# 模組層狀態（load_dict 會依 config 覆寫）
STOPWORDS = set(BASE_STOPWORDS)
KEEP_SHORT = set(BASE_KEEP_SHORT)
SYNONYMS = {}
QUESTION_TYPES = dict(DEFAULT_QUESTION_TYPES)
_LOADED = False


def load_dict(user_words=None, stopwords_extra=None, keep_short=None,
              synonyms=None, question_types=None):
    """把課程自訂詞灌進 jieba，並合併停用詞／短詞白名單／同義詞／提問關鍵詞。"""
    global STOPWORDS, KEEP_SHORT, SYNONYMS, QUESTION_TYPES, _LOADED
    for w in list(BASE_USER_WORDS) + list(user_words or []):
        if w:
            jieba.add_word(w, freq=100000)
    for c, syns in (synonyms or {}).items():
        jieba.add_word(c, freq=100000)
        for s in syns or []:
            if s:
                jieba.add_word(s, freq=100000)
    STOPWORDS = set(BASE_STOPWORDS) | set(stopwords_extra or [])
    KEEP_SHORT = set(BASE_KEEP_SHORT) | set(keep_short or [])
    SYNONYMS = {k: list(v or []) for k, v in (synonyms or {}).items()}
    qt = {k: list(v) for k, v in DEFAULT_QUESTION_TYPES.items()}
    for k, v in (question_types or {}).items():
        if v:
            qt[k] = list(v)
    QUESTION_TYPES = qt
    _LOADED = True


def clean(t):
    return URL.sub("", str(t or "")).strip()


def is_valid(t):
    if not t or PUNCT_ONLY.match(t):
        return False
    s = t.strip()
    # 遮罩殘影（OOO…）與長數字串（學號、電話）不當概念
    if "OOO" in s or re.search(r"\d{8,}", s):
        return False
    return s not in PLACEHOLDER


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


def student_key(r):
    """學生主鍵：學生編號優先，否則退回學號／姓名。"""
    return (r.get("學生編號") or "").strip() or (r.get("學號") or "").strip() \
        or (r.get("姓名") or "").strip()


def speakers_by_word(rows):
    """{詞: 提到這個詞的同學人數}。用斷詞結果計算，才會和詞頻同一套標準。"""
    d = defaultdict(set)
    for r in rows:
        for w in set(tokens(r["作答內容"])):
            d[w].add(student_key(r))
    return {w: len(v) for w, v in d.items()}


# ---------------------------------------------------------------- 題型判斷
def is_open_text(sub_label):
    """只看子題號：「第2題:單選題」→ False；「第1題:問答題」→ True。"""
    return not CHOICE_SUB.search(sub_label or "")


def sub_is_open(sub_label, texts):
    """看子題號＋實際答案內容決定是不是「開放文字」。

    測驗題目型匯出的子題號只是「第1題」，但答案是「答對／答錯」；
    選擇題的答案是少數幾個固定選項。這兩種都不做文字雲與概念矩陣。
    """
    if not is_open_text(sub_label):
        return False
    vals = [str(t or "").strip() for t in texts]
    vals = [v for v in vals if v]
    if not vals:
        return False
    uniq = set(vals)
    if len(uniq) <= 6 and max(len(v) for v in uniq) <= 12:
        return False                       # 答對/答錯、(1)男/(2)女 這種
    if sum(len(v) for v in vals) / len(vals) < 8:
        return False                       # 平均不到 8 個字，不值得做文字分析
    return True


def open_sub_labels(rows):
    """回傳這一題裡「算開放文字」的子題號集合。CSV 列 → set(子題號)。"""
    per = defaultdict(list)
    for r in rows:
        per[r.get("子題號", "")].append(r.get("作答內容", ""))
    return {lab for lab, texts in per.items() if sub_is_open(lab, texts)}


# ---------------------------------------------------------------- 概念矩陣
def canon_map():
    """{任一寫法: 正式概念名}。"""
    m = {}
    for c, syns in SYNONYMS.items():
        m[c] = c
        for s in syns:
            m[s] = c
    return m


def variants_of(concept):
    return [concept] + [s for s in SYNONYMS.get(concept, []) if s]


def code_sort_key(code):
    """115-1_EC_2 要排在 115-1_EC_10 前面。"""
    m = re.search(r"_(\d+)$", code or "")
    return (0, int(m.group(1))) if m else (1, 0, code or "")


def pick_concepts(rows, top_n=6, pool=60):
    """回傳 (concepts, per_student, cfreq, speakers)。

    concepts   = [概念名]，依「提及人數」由多到少（同分比詞頻，再比字典序），
                 最多 top_n 個；**候選不足 top_n 時有幾個列幾個，不補空字串**。
    per_student= {學生編號: (token集合, 原文)}
    speakers   = {概念: 提到它的學生編號集合}（候選池全體，可用來算候選數）
    """
    per_student = {}
    for r in rows:
        k = student_key(r)
        toks, text = per_student.get(k, (set(), ""))
        t = r["作答內容"]
        per_student[k] = (toks | set(tokens(t)), text + " " + t)

    freq = Counter()
    for t in (r["作答內容"] for r in rows):
        freq.update(tokens(t))

    cm = canon_map()
    cfreq = Counter()
    for w, c in freq.items():
        cfreq[cm.get(w, w)] += c

    cands = [w for w, _ in cfreq.most_common(pool)]
    speakers = {}
    for c in cands:
        vs = variants_of(c)
        s = set()
        for k, (toks, text) in per_student.items():
            if any((v in toks) or (v and v in text) for v in vs):
                s.add(k)
        speakers[c] = s

    ranked = sorted(cands, key=lambda w: (-len(speakers[w]), -cfreq[w], w))
    return ranked[:top_n], per_student, cfreq, speakers


def concept_matrix(concepts, per_student):
    """回傳 (學生編號排序清單, {學生編號: [0/1,...]})。"""
    keys = sorted(per_student.keys(), key=code_sort_key)
    mat = {}
    for k in keys:
        toks, text = per_student[k]
        mat[k] = [1 if any((v in toks) or (v and v in text) for v in variants_of(c)) else 0
                  for c in concepts]
    return keys, mat


def write_concept_matrix(path, concepts, keys, mat):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["學生編號"] + list(concepts) + ["提到重點數"])
        for k in keys:
            row = mat[k]
            w.writerow([k] + row + [sum(row)])


# ---------------------------------------------------------------- 提問
def sentences(text):
    out, buf = [], ""
    for ch in str(text or ""):
        buf += ch
        if ch in SENT_END:
            if buf.strip():
                out.append(buf.strip())
            buf = ""
    if buf.strip():
        out.append(buf.strip())
    return out


def is_question_sentence(s):
    """句尾問號／句尾嗎呢／句首疑問詞，三者之一才算提問（陳述句不算）。"""
    t = str(s or "").strip()
    if not t:
        return False
    if QUESTION_END.search(t):
        return True
    if QUESTION_TAIL.search(t):
        return True
    if QUESTION_LEAD.match(t):
        return True
    return bool(QUESTION_LEAD_EN.match(t))


def _kw_hits(s, keywords):
    """關鍵詞命中數；英數關鍵詞不分大小寫。"""
    low = s.lower()
    n = 0
    for k in keywords:
        if not k:
            continue
        n += 1 if ((k.lower() in low) if k.isascii() else (k in s)) else 0
    return n


def fallback_type():
    """後備類別：預設「概念理解類」；若使用者自訂的類別名沒有它，就用最後一類。"""
    if FALLBACK_TYPE in QUESTION_TYPES:
        return FALLBACK_TYPE
    names = list(QUESTION_TYPES.keys())
    return names[-1] if names else OTHER_TYPE


def classify_question(s):
    """規則式四類分類；四類都沒中就走後備規則，再沒中才算「其他」。

    QUESTION_TYPES 是保序 dict，**先後順序就是判定優先序**：
    命中數嚴格大於目前最佳才換人，所以命中數相同時排前面的類別勝出。
    """
    best, best_hits = OTHER_TYPE, 0
    for name in QUESTION_TYPES:                     # dict 保序 → 順序即優先序
        hits = _kw_hits(s, QUESTION_TYPES[name])
        if hits > best_hits:
            best, best_hits = name, hits
    if best_hits:
        return best, best_hits
    if _kw_hits(s, FALLBACK_CONCEPT):
        # 後備：是問句但沒中關鍵詞，含「什麼／嗎／為／Why」→ 概念理解類
        return fallback_type(), 0
    return OTHER_TYPE, 0


def tidy_question(s, limit=40):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    s = s.lstrip("，,、。 ")
    return s[:limit] + ("…" if len(s) > limit else "")


def extract_questions(rows):
    """回傳 (records, stats)。

    records = [{"學生編號", "提問句", "分類"}]
    stats   = {類別: {"提問數", "提問人數", "代表句"}}（含「其他」）
    """
    recs = []
    for r in rows:
        k = student_key(r)
        for s in sentences(r["作答內容"]):
            if not is_question_sentence(s):
                continue
            cat, hits = classify_question(s)
            recs.append({"學生編號": k, "提問句": tidy_question(s, 120),
                         "分類": cat, "_hits": hits, "_len": len(s)})

    stats = {}
    for name in list(QUESTION_TYPES.keys()) + [OTHER_TYPE]:
        sub = [x for x in recs if x["分類"] == name]
        rep = ""
        if sub:
            pick = sorted(sub, key=lambda x: (-x["_hits"], x["_len"], x["提問句"]))[0]
            rep = tidy_question(pick["提問句"], 40)
        stats[name] = {"提問數": len(sub),
                       "提問人數": len({x["學生編號"] for x in sub}),
                       "代表句": rep}
    for x in recs:
        x.pop("_hits", None)
        x.pop("_len", None)
    return recs, stats


def write_questions(path, recs):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["學生編號", "提問句", "分類"])
        for x in recs:
            w.writerow([x["學生編號"], x["提問句"], x["分類"]])


# ---------------------------------------------------------------- 主流程
def analyse_dir(work_dir, index=None, top_n=15, top_concepts=6, log=print):
    """讀 work_dir 下的 Q*.csv，寫出 analysis.json、summary.md 與每題兩份 CSV。"""
    meta = {r["題號"]: r for r in (index or [])}
    result = {"資料夾": os.path.basename(os.path.normpath(work_dir)), "題目": []}
    all_freq = Counter()
    all_questions = []

    paths = [p for p in sorted(glob.glob(os.path.join(glob.escape(work_dir), "Q*.csv")))
             if not os.path.basename(p).endswith(("_concept_matrix.csv", "_questions.csv"))]

    for path in paths:
        qno = os.path.basename(path)[:3]
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))

        cleaned = []
        for r in rows:
            t = clean(r.get("作答內容", ""))
            if not is_valid(t):
                continue
            cleaned.append(dict(r, 作答內容=t))

        open_labels = open_sub_labels(cleaned)
        valid_rows, texts = [], []
        per_sub = defaultdict(list)
        students = set()
        for r in cleaned:
            per_sub[(r["子題號"], r.get("子題題目", ""))].append(r)
            students.add(student_key(r))
            if r["子題號"] in open_labels:
                valid_rows.append(r)
                texts.append(r["作答內容"])

        freq = freq_of(texts)
        all_freq.update(freq)

        # ---- 概念矩陣與覆蓋率
        concepts, per_student, cfreq, cspk = pick_concepts(valid_rows, top_n=top_concepts)
        keys, mat = concept_matrix(concepts, per_student)
        n_ans = len(keys)
        cm_name = f"{qno}_concept_matrix.csv"
        if concepts:
            write_concept_matrix(os.path.join(work_dir, cm_name), concepts, keys, mat)
        else:
            write_concept_matrix(os.path.join(work_dir, cm_name), [], keys, mat)

        focus = []
        for i, c in enumerate(concepts):
            hit = sum(mat[k][i] for k in keys)
            focus.append({"概念": c, "提及人數": hit, "次數": int(cfreq.get(c, 0)),
                          "覆蓋率": round(hit / n_ans, 3) if n_ans else 0.0})
        # 覆蓋率三件事：至少 1 個、≥3 個、全部（六個）都提到
        any_hit = sum(1 for k in keys if any(mat[k])) if concepts else 0
        three_hit = sum(1 for k in keys if sum(mat[k]) >= 3) if concepts else 0
        all_hit = sum(1 for k in keys if concepts and all(mat[k])) if concepts else 0
        overall = round(any_hit / n_ans, 3) if n_ans else 0.0
        three_up = round(three_hit / n_ans, 3) if n_ans else 0.0
        all_cov = round(all_hit / n_ans, 3) if n_ans else 0.0

        # ---- 提問抽取與分類
        qrecs, qstats = extract_questions(valid_rows)
        q_name = f"{qno}_questions.csv"
        write_questions(os.path.join(work_dir, q_name), qrecs)

        questions = [t[:60] for t in texts if any(k in t for k in QUESTION_MARK)]
        all_questions += questions

        subs = []
        for (label, text), items in per_sub.items():
            itexts = [r["作答內容"] for r in items]
            is_open = label in open_labels
            sub = {
                "子題號": label,
                "子題題目": text,
                "型別": "開放文字" if is_open else "選擇題／測驗",
                "有效作答數": len(items),
                "作答人數": len({student_key(r) for r in items}),
                "平均字數": round(sum(len(t) for t in itexts) / max(len(itexts), 1), 1),
            }
            if is_open:
                sub["TOP詞"] = freq_of(itexts).most_common(12)
            else:
                c = Counter(itexts)
                sub["TOP詞"] = []
                sub["選項分佈"] = [{"選項": o, "人數": k,
                                 "百分比": round(k * 100.0 / max(len(itexts), 1), 1)}
                                for o, k in c.most_common()]
                # 測驗型（答案只有 答對／答錯）才算得出答對率
                right = sum(1 for t in itexts if t.strip() in CORRECT_WORDS)
                wrong = sum(1 for t in itexts if t.strip() in WRONG_WORDS)
                if right + wrong == len(itexts) and len(itexts):
                    sub["答對率"] = round(right * 100.0 / len(itexts), 1)
                    sub["答對人數"] = right
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
            "測驗型": bool(m.get("測驗型")) or not open_labels,
            "來源檔": m.get("來源檔", ""),
            "作答人數": answered,
            "全班人數": klass,
            "作答率": round(answered * 100.0 / max(klass, 1), 1),
            "原始列數": len(rows),
            "有效列數": sum(len(v) for v in per_sub.values()),
            "文字列數": len(texts),
            "文字作答人數": n_ans,
            "平均字數": round(sum(len(t) for t in texts) / max(len(texts), 1), 1),
            "總詞數": sum(freq.values()),
            "相異詞數": len(freq),
            "TOP詞": freq.most_common(top_n),
            "TOP詞明細": top_rows,
            "重點概念": focus,
            "候選概念數": len(cspk),
            "整體覆蓋率": overall,
            "提到三個以上比例": three_up,
            "全部提到比例": all_cov,
            "概念矩陣檔": cm_name,
            "概念矩陣": {k: mat[k] for k in keys},
            "提問統計": [{"類別": name, **qstats[name]} for name in QUESTION_TYPES],
            "其他提問數": qstats[OTHER_TYPE]["提問數"],
            "提問總數": len(qrecs),
            "提問人數": len({x["學生編號"] for x in qrecs}),
            "提問檔": q_name,
            "常見問題": questions[:5],
            "子題": subs,
        })
        log(f"  {qno} 原始 {len(rows)} 列 → 有效 {sum(len(v) for v in per_sub.values())} 列"
            f"（開放文字 {len(texts)} 列／{n_ans} 人），子題 {len(subs)}")
        if concepts:
            log(f"       {len(concepts)} 個重點 = " +
                "、".join(f"{d['概念']}({d['提及人數']}人/{round(d['覆蓋率'] * 100)}%)"
                         for d in focus) +
                f"；整體覆蓋率 {round(overall * 100)}%；"
                f"提到三個以上 {round(three_up * 100)}%；"
                f"全部提到 {round(all_cov * 100)}%")
        else:
            log("       本題無開放文字作答（選擇題／測驗型），不做概念矩陣與文字雲。")
        log(f"       提問 {len(qrecs)} 則／{len({x['學生編號'] for x in qrecs})} 人：" +
            "、".join(f"{name} {qstats[name]['提問數']} 則" for name in QUESTION_TYPES) +
            f"、其他 {qstats[OTHER_TYPE]['提問數']} 則")

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
         "## 各題作答率與概念覆蓋率", "",
         "| 題號 | 題目 | 作答/全班 | 作答率 | 整體覆蓋率 | 提到 ≥3 個 | 六個都提到 | 提問數 |",
         "|---|---|---|---|---|---|---|---|"]
    for q in result["題目"]:
        L.append(f"| {q['題號']} | {q['題目']} | {q['作答人數']}/{q['全班人數']} | "
                 f"{q['作答率']}% | {round(q['整體覆蓋率'] * 100, 1)}% | "
                 f"{round(q.get('提到三個以上比例', 0) * 100, 1)}% | "
                 f"{round(q.get('全部提到比例', 0) * 100, 1)}% | {q['提問總數']} |")
    L += ["", "## 全班最常出現重點 TOP10", ""]
    for i, (w, c) in enumerate(result["全班重點TOP10"], 1):
        L.append(f"{i}. {w}（{c} 次）")
    L += ["", "## 各題六個重點與四類提問", ""]
    for q in result["題目"]:
        L.append(f"### {q['題號']}　{q['題目']}")
        if q["重點概念"]:
            L.append(f"- 六個重點（提及人數／覆蓋率，實際 {len(q['重點概念'])} 個）：" +
                     "、".join(f"{d['概念']} {d['提及人數']} 人（{round(d['覆蓋率'] * 100, 1)}%）"
                              for d in q["重點概念"]))
            L.append(f"- 整體覆蓋率（至少提到 1 個重點）：{round(q['整體覆蓋率'] * 100, 1)}%；"
                     f"提到 ≥3 個重點：{round(q.get('提到三個以上比例', 0) * 100, 1)}%；"
                     f"六個都提到：{round(q.get('全部提到比例', 0) * 100, 1)}%")
            L.append(f"- 概念矩陣：`{q['概念矩陣檔']}`（{len(q['重點概念'])} 個概念欄）")
        else:
            L.append("- 本題沒有開放文字作答（選擇題／測驗型），不做概念矩陣。")
        for t in q["提問統計"]:
            L.append(f"- {t['類別']}：{t['提問數']} 則／{t['提問人數']} 人"
                     + (f"；代表句「{t['代表句']}」" if t["代表句"] else ""))
        L.append(f"- 其他提問：{q['其他提問數']} 則；提問明細：`{q['提問檔']}`")
        L.append("- TOP 詞：" + "、".join(f"{w}({c})" for w, c in q["TOP詞"][:8]))
        for s in q["子題"]:
            head = f"  - {s['子題號']} {s['子題題目'][:40]}　有效 {s['有效作答數']} 列；"
            if s.get("選項分佈"):
                L.append(head + "選項分佈：" +
                         "、".join(f"{o['選項']} {o['人數']}人({o['百分比']}%)"
                                  for o in s["選項分佈"][:6]))
            else:
                L.append(head + "TOP：" + "、".join(w for w, _ in s["TOP詞"][:6]))
        L.append("")
    if result["常見問題TOP5"]:
        L += ["## 同學的疑問（節錄）", ""]
        for i, s in enumerate(result["常見問題TOP5"], 1):
            L.append(f"{i}. {s}")
    L += ["", "---", "",
          "> 本檔一律使用學生編號（例 115-1_EC_3），姓名已遮罩、學號已全部改成 O。",
          "> 「學生編號連結姓名」對照表只能留在本機，不得上傳、不得外流。"]

    with open(os.path.join(work_dir, "summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(L))
