"""APA 7 參考文獻格式（期刊／會議／專書章節／專書／專利／網頁／中文）。

format_apa7(paper) -> str             純文字
format_apa7_runs(paper) -> [(text, style)]  style: ""、"i"（斜體）、"b"（粗體）——docx／pdf 用
in_text(paper) -> str                 文內引用
sort_key(paper)                       先中文（筆畫）後英文（字母）
"""
from __future__ import annotations

import ctypes
import html
import re
import sys
import unicodedata
from typing import List, Tuple

from .models import Paper, has_cjk

Run = Tuple[str, str]

# ---------------------------------------------------------------- 字串工具
_TAG = re.compile(r"<[^>]+>")


def clean_text(s: str) -> str:
    s = html.unescape(_TAG.sub("", s or ""))
    return re.sub(r"\s+", " ", s).strip()


def dash_range(pages: str) -> str:
    p = (pages or "").strip()
    p = re.sub(r"\s*[-‐‑‒–—]+\s*", "–", p)
    return p


_ELEMENTS = set("""H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr
Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os
Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr""".split())
# 這些元素符號同時是常見英文字，不視為元素
_ELEMENT_WORDS = {"He", "Be", "In", "As", "At", "No", "Am", "Es", "I", "B", "C", "S", "W", "Y", "U", "K", "N", "O",
                  "F", "H", "V", "P", "Ho", "Ba", "La", "Pa", "Re", "Ra", "Po", "Co", "Ca", "Ga", "Ge"}
_KEEP_ELEMENTS = _ELEMENTS - _ELEMENT_WORDS | {"Ca", "Co", "Fe"}

# 常見需保留大寫之專有名詞（標題轉句首大寫時不改）
PROPER = {"python", "numpy", "taiwan", "hualien", "europe", "european", "asia", "asian", "pacific", "atlantic",
          "china", "chinese", "japan", "japanese", "korea", "korean", "english", "american", "african", "india",
          "indian", "earth", "mars", "covid-19", "sars-cov-2", "raman", "fourier", "langmuir", "freundlich",
          "pourbaix", "bragg", "planck", "kubelka-munk", "mössbauer", "scherrer", "brunauer-emmett-teller",
          "atlantic", "mediterranean", "arctic", "antarctic", "ajzen", "piaget", "vygotsky", "bloom"}
_SMALL = {"a", "an", "the", "and", "or", "of", "in", "on", "at", "to", "for", "by", "with", "from", "as", "via",
          "into", "nor", "but", "vs", "vs.", "per", "over", "under", "upon"}


def _is_title_case(title: str) -> bool:
    words = [w for w in re.findall(r"[A-Za-z][A-Za-z'’]*", title)]
    long = [w for w in words if len(w) > 3 and w.lower() not in _SMALL]
    if len(long) < 3:
        return False
    caps = sum(1 for w in long if w[0].isupper())
    return caps / len(long) >= 0.6


def _keep_word(w: str) -> bool:
    """保留原樣：縮寫、含數字、內部大寫、化學式、元素符號、專有名詞。"""
    core = w.strip("()[]{}\"'“”‘’,.;:!?")
    if not core:
        return True
    if any(ch.isdigit() for ch in core):
        return True
    if sum(1 for ch in core[1:] if ch.isupper()) >= 1:     # FIR、NumPy、TPGS、pH 型
        return True
    if core in _KEEP_ELEMENTS:
        return True
    if core.lower() in PROPER:
        return True
    return False


def _sentence_word(w: str, first: bool) -> str:
    if any(ch.isdigit() for ch in w) or (w.isupper() and len(w.strip("()[].,:;")) >= 2):
        return w
    parts = re.split(r"([-‐–—/])", w)
    out = []
    for i, part in enumerate(parts):
        if i % 2 == 1 or not part:
            out.append(part)
            continue
        if _keep_word(part):
            out.append(part)
            continue
        lead = re.match(r"^([\(\[\"'“‘]*)(.*)$", part)
        pre, body = lead.group(1), lead.group(2)
        if first and i == 0:
            out.append(pre + body[:1].upper() + body[1:].lower())
        else:
            out.append(pre + body.lower())
    return "".join(out)


def sentence_case(title: str) -> str:
    """APA 7：題名句首大寫，冒號／問號後首字大寫；縮寫、化學式、專有名詞保留。"""
    t = clean_text(title)
    if not t or has_cjk(t) and not re.search(r"[A-Za-z]{4,}", t):
        return t
    if _is_title_case(t):
        tokens = t.split(" ")
        out = []
        cap_next = True
        for tok in tokens:
            out.append(_sentence_word(tok, cap_next))
            cap_next = bool(re.search(r"[:?!.]$", tok)) and not re.search(r"\b[A-Z]\.$", tok)
        t = " ".join(out)
    # 冒號後首字一律大寫
    t = re.sub(r"([:?!]\s+)([\"'“‘(]*)([a-z])", lambda m: m.group(1) + m.group(2) + m.group(3).upper(), t)
    if t and t[0].islower():
        t = t[0].upper() + t[1:]
    return t


def title_case_journal(name: str) -> str:
    return clean_text(name)


# ---------------------------------------------------------------- 作者
def _initials(given: str) -> str:
    given = clean_text(given)
    if not given:
        return ""
    out = []
    for tok in given.replace(".", ". ").split():
        if "-" in tok:
            subs = [s for s in tok.split("-") if s]
            out.append("-".join(s[0].upper() + "." for s in subs))
        else:
            t = tok.strip(".")
            if not t:
                continue
            out.append(t[0].upper() + ".")
    return " ".join(out)


def _en_author(a: dict) -> str:
    fam = clean_text(a.get("family", ""))
    giv = a.get("given", "")
    if not giv:
        return fam
    ini = _initials(giv)
    return f"{fam}, {ini}" if fam else ini


def _en_author_list(authors: list) -> str:
    names = [_en_author(a) for a in authors if (a.get("family") or a.get("given"))]
    n = len(names)
    if n == 0:
        return ""
    if n == 1:
        return names[0]
    if n == 2:
        return f"{names[0]}, & {names[1]}"
    if n <= 20:
        return ", ".join(names[:-1]) + ", & " + names[-1]
    return ", ".join(names[:19]) + ", . . . " + names[-1]


def _zh_name(a: dict) -> str:
    fam, giv = clean_text(a.get("family", "")), clean_text(a.get("given", ""))
    if has_cjk(fam + giv):
        return fam + giv
    return _en_author(a)


def _zh_author_list(authors: list) -> str:
    names = [_zh_name(a) for a in authors if (a.get("family") or a.get("given"))]
    if len(names) > 20:
        names = names[:19] + ["…", names[-1]]
    return "、".join(names)


def _editor_list(eds: list) -> str:
    names = []
    for a in eds:
        ini = _initials(a.get("given", ""))
        fam = clean_text(a.get("family", ""))
        names.append(f"{ini} {fam}".strip())
    if not names:
        return ""
    if len(names) == 1:
        return names[0] + " (Ed.)"
    if len(names) == 2:
        return f"{names[0]} & {names[1]} (Eds.)"
    return ", ".join(names[:-1]) + ", & " + names[-1] + " (Eds.)"


def _end(s: str) -> str:
    s = s.rstrip()
    return s if s.endswith((".", "?", "!")) else s + "."


def _doi_or_url(p: Paper) -> str:
    if p.doi:
        return f"https://doi.org/{p.doi}"
    return p.url or ""


def _year(p: Paper, zh: bool = False) -> str:
    y = str(p.year) if p.year else ("無日期" if zh else "n.d.")
    if p.type == "webpage" and p.month_day and p.year and not zh:
        y = f"{p.year}, {p.month_day}"
    return y


def _locator_en(p: Paper) -> List[Run]:
    """卷(期), 頁／Article N"""
    runs: List[Run] = []
    if p.volume:
        runs.append((", ", ""))
        runs.append((p.volume, "i"))
        if p.issue:
            runs.append((f"({p.issue})", ""))
    elif p.issue:
        runs.append((f", ({p.issue})", ""))
    if p.pages:
        runs.append((", " + dash_range(p.pages), ""))
    elif p.article_number:
        runs.append((f", Article {p.article_number}", ""))
    return runs


# ---------------------------------------------------------------- 主函式
def format_apa7_runs(p: Paper) -> List[Run]:
    if p.is_chinese:
        return _zh_runs(p)
    t = p.type or "journal-article"
    authors = _en_author_list(p.authors)
    title = sentence_case(p.title)
    tail = _doi_or_url(p)
    runs: List[Run] = []
    head = authors
    if head:
        if not head.endswith("."):
            head += "."
        runs.append((head + f" ({_year(p)}). ", ""))
    if t in ("journal-article", "posted-content", "other", "report", "dissertation") or not t:
        if not head:
            runs.append((_end(title) + f" ({_year(p)}). ", ""))
        else:
            runs.append((_end(title) + " ", ""))
        if p.container:
            runs.append((clean_text(p.container), "i"))
            runs += _locator_en(p)
            runs.append((". ", ""))
    elif t == "proceedings-article":
        runs.append((_end(title) + " ", ""))
        if p.container:
            runs.append(("In ", ""))
            if p.editors:
                runs.append((_editor_list(p.editors) + ", ", ""))
            runs.append((clean_text(p.container), "i"))
            if p.pages:
                runs.append((f" (pp. {dash_range(p.pages)})", ""))
            elif p.article_number:
                runs.append((f" (Article {p.article_number})", ""))
            runs.append((". ", ""))
        if p.publisher and p.publisher not in (p.container or ""):
            runs.append((_end(clean_text(p.publisher)) + " ", ""))
    elif t == "book-chapter":
        runs.append((_end(title) + " ", ""))
        book = p.book_title or p.container
        runs.append(("In ", ""))
        if p.editors:
            runs.append((_editor_list(p.editors) + ", ", ""))
        runs.append((clean_text(book), "i"))
        extra = []
        if p.edition:
            extra.append(f"{p.edition} ed.")
        if p.pages:
            extra.append(f"pp. {dash_range(p.pages)}")
        runs.append(((f" ({', '.join(extra)})" if extra else "") + ". ", ""))
        if p.publisher:
            runs.append((_end(clean_text(p.publisher)) + " ", ""))
    elif t == "book":
        runs = [(head + f" ({_year(p)}). ", "")] if head else []
        runs.append((sentence_case(p.title), "i"))
        runs.append(((f" ({p.edition} ed.)" if p.edition else "") + ". ", ""))
        if p.publisher:
            runs.append((_end(clean_text(p.publisher)) + " ", ""))
    elif t == "patent":
        num = p.patent_number or ""
        office = p.patent_office or p.publisher or "Patent office"
        runs.append((sentence_case(p.title), "i"))
        runs.append((f" ({_patent_label_en(num)}). ", ""))
        runs.append((_end(office) + " ", ""))
        if not tail and num:
            tail = f"https://patents.google.com/patent/{num}"
    elif t == "webpage":
        runs = []
        if head:
            runs.append((head + f" ({_year(p)}). ", ""))
            runs.append((_end(sentence_case(p.title)).rstrip(".") , "i"))
            runs.append((". ", ""))
        else:
            runs.append((sentence_case(p.title), "i"))
            runs.append((f". ({_year(p)}). ", ""))
        site = clean_text(p.container)
        if site and (not p.authors or site != _en_author(p.authors[0])):
            runs.append((_end(site) + " ", ""))
        if p.accessed and not p.year:
            runs.append((f"Retrieved {p.accessed}, from ", ""))
    else:
        runs.append((_end(title) + " ", ""))
        if p.container:
            runs.append((clean_text(p.container), "i"))
            runs.append((". ", ""))
    if tail:
        runs.append((tail, ""))
    out = _merge_runs(runs)
    return out


def _patent_label_en(num: str) -> str:
    m = re.match(r"^(US|TW|CN|JP|EP|WO|KR)\s*([\w,/.-]+)$", num.replace(" ", ""), re.I)
    if not m:
        return f"Patent No. {num}" if num else "Patent"
    cc, n = m.group(1).upper(), m.group(2)
    label = {"US": "U.S. Patent No.", "TW": "Taiwan Patent No.", "CN": "Chinese Patent No.", "JP": "Japanese Patent No.",
             "EP": "European Patent No.", "WO": "International Patent Application No.", "KR": "Korean Patent No."}[cc]
    return f"{label} {cc}{n}" if cc != "US" else f"{label} {n}"


def _zh_runs(p: Paper) -> List[Run]:
    authors = _zh_author_list(p.authors)
    y = _year(p, zh=True)
    title = clean_text(p.title)
    runs: List[Run] = []
    head = f"{authors}（{y}）。" if authors else ""
    if p.type == "patent":
        num = p.patent_number
        office = p.publisher or "經濟部智慧財產局"
        kind = p.patent_office or "中華民國專利公開"
        label = f"（{kind}第 {num} 號）" if num else ""
        runs.append((head or f"{title}（{y}）。", ""))
        if head:
            runs.append((title + label + "。", ""))
        runs.append((office.rstrip("。") + "。", ""))
    elif p.type == "book":
        runs.append((head, ""))
        runs.append((title, "b"))
        runs.append(("。", ""))
        if p.publisher:
            runs.append((p.publisher.rstrip("。") + "。", ""))
    elif p.type == "book-chapter":
        runs.append((head + title + "。", ""))
        eds = "、".join(_zh_name(a) for a in p.editors)
        runs.append((f"載於{eds}（主編），" if eds else "載於", ""))
        runs.append((clean_text(p.book_title or p.container), "b"))
        runs.append(((f"（頁 {dash_range(p.pages)}）" if p.pages else "") + "。", ""))
        if p.publisher:
            runs.append((p.publisher.rstrip("。") + "。", ""))
    elif p.type == "webpage":
        runs.append((head, ""))
        runs.append((title, "b"))
        runs.append(("。", ""))
        if p.container:
            runs.append((p.container.rstrip("。") + "。", ""))
    else:
        runs.append((head or f"{title}（{y}）。", ""))
        if head:
            runs.append((title.rstrip("。") + "。", ""))
        if p.container:
            runs.append((clean_text(p.container), "b"))
            if p.volume:
                runs.append(("，", ""))
                runs.append((p.volume, "b"))
                if p.issue:
                    runs.append((f"({p.issue})", ""))
            elif p.issue:
                runs.append((f"，({p.issue})", ""))
            if p.pages:
                runs.append(("，" + dash_range(p.pages), ""))
            elif p.article_number:
                runs.append((f"，文章 {p.article_number}", ""))
            runs.append(("。", ""))
    tail = _doi_or_url(p)
    if tail:
        runs.append((tail, ""))
    return _merge_runs(runs)


def _merge_runs(runs: List[Run]) -> List[Run]:
    out: List[Run] = []
    for text, st in runs:
        if not text:
            continue
        if out and out[-1][1] == st:
            out[-1] = (out[-1][0] + text, st)
        else:
            out.append((text, st))
    if out:
        out[-1] = (out[-1][0].rstrip(), out[-1][1])
    return out


def format_apa7(paper: Paper) -> str:
    return "".join(t for t, _ in format_apa7_runs(paper))


# ---------------------------------------------------------------- 文內引用
def in_text(paper: Paper) -> str:
    au = [a for a in paper.authors if (a.get("family") or a.get("given"))]
    if paper.is_chinese:
        y = str(paper.year) if paper.year else "無日期"
        names = [_zh_name(a) if has_cjk(_zh_name(a)) else clean_text(a.get("family", "")) for a in au]
        if not names:
            who = clean_text(paper.title)[:12]
        elif len(names) == 1:
            who = names[0]
        elif len(names) == 2:
            who = f"{names[0]}、{names[1]}"
        else:
            who = f"{names[0]}等人"
        return f"（{who}，{y}）"
    y = str(paper.year) if paper.year else "n.d."
    fams = [clean_text(a.get("family") or a.get("given", "")) for a in au]
    if not fams:
        who = sentence_case(paper.title)
        who = "“" + (who[:40] + ("…" if len(who) > 40 else "")) + "”"
    elif len(fams) == 1:
        who = fams[0]
    elif len(fams) == 2:
        who = f"{fams[0]} & {fams[1]}"
    else:
        who = f"{fams[0]} et al."
    return f"({who}, {y})"


# ---------------------------------------------------------------- 排序鍵
def _win_stroke_key(s: str):
    """Windows zh-TW 預設排序即筆畫序；以 LCMapStringEx 取 sort key。"""
    if sys.platform != "win32" or not s:
        return None
    try:
        f = ctypes.windll.kernel32.LCMapStringEx
        LCMAP_SORTKEY = 0x00000400
        n = f("zh-TW", LCMAP_SORTKEY, s, len(s), None, 0, None, None, 0)
        if n <= 0:
            return None
        buf = ctypes.create_string_buffer(n)
        f("zh-TW", LCMAP_SORTKEY, s, len(s), buf, n, None, None, 0)
        return buf.raw[:n]
    except Exception:
        return None


def _fold(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch)).lower()


def sort_key(paper: Paper):
    """APA：先中文（第一作者筆畫）後英文（第一作者姓氏字母）；同作者依年份。"""
    if paper.is_chinese:
        who = "".join(_zh_name(a) for a in paper.authors[:1]) or clean_text(paper.title)
        k = _win_stroke_key(who)
        if k is None:
            k = who.encode("utf-16-be")
        return (0, k, paper.year or 0, clean_text(paper.title))
    names = " ".join(_fold(clean_text(a.get("family", "")) + " " + _initials(a.get("given", ""))) for a in paper.authors)
    if not names.strip():
        names = _fold(re.sub(r"^(the|a|an)\s+", "", clean_text(paper.title), flags=re.I))
    return (1, names.encode("utf-8"), paper.year or 9999, _fold(clean_text(paper.title)))
