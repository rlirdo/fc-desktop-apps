"""合格指標 ①②③⑥⑦⑩ 與相關性 ⑫。

settings 鍵（皆可省略，預設見 config.DEFAULTS）：
  checks   {"doi":bool,"oa":bool,"type":bool,"year":bool,"issn":bool,"lang":bool}  勾選＝納入判定
  types    ["journal-article", "proceedings-article", "book-chapter", ...]
  year_from / year_to / classic(bool：含經典不限年)
  lang     "en" | "zh" | "any"
  weights  {"title":3,"abstract":1,"cited":1,"recency":1}
  keywords ["serpentine; 蛇紋石", "deep sea water; 深層海水"]   相關性用（每組「;」分同義詞）
  keep_mode "mark"（全部保留但標示）| "strict"（只保留全部通過）
indicators[name] = (True|False|None, 說明)；None＝待人工確認；未勾選者不寫入。
"""
from __future__ import annotations

import datetime as _dt
import math
import re
from typing import Optional

from .models import Paper, TYPE_LABELS, has_kana

# 主要出版社（⑦：ISSN 存在且在 DOAJ 或為主要出版社者視為通過）
MAJOR_PUBLISHERS = [
    "elsevier", "springer", "nature", "wiley", "taylor & francis", "taylor and francis", "informa", "sage",
    "american chemical society", "royal society of chemistry", "mdpi", "ieee", "institute of electrical",
    "oxford university press", "cambridge university press", "iop publishing", "frontiers", "plos",
    "public library of science", "de gruyter", "walter de gruyter", "american physical society", "aip publishing",
    "american institute of physics", "acm", "association for computing machinery", "emerald", "hindawi",
    "copernicus", "american geophysical union", "mineralogical society", "geological society", "cell press",
    "american association for the advancement of science", "aaas", "the royal society", "thieme", "karger",
    "bmj", "jstor", "university of chicago press", "mit press", "annual reviews", "american society",
    "american meteorological society", "spie", "optica", "the optical society", "j-stage", "ios press",
    "inderscience", "brill", "routledge", "national academy of sciences", "american psychological association",
]


def _terms(keywords) -> list:
    out = []
    for g in keywords or []:
        if isinstance(g, (list, tuple)):
            items = g
        else:
            items = re.split(r"[;；]", str(g))
        terms = [t.strip().lower() for t in items if t.strip()]
        if terms:
            out.append(terms)
    return out


def _hits(text: str, groups: list) -> int:
    t = (text or "").lower()
    n = 0
    for g in groups:
        if any(term in t for term in g):
            n += 1
    return n


def relevance(p: Paper, groups: list, weights: dict, now_year: Optional[int] = None) -> float:
    now_year = now_year or _dt.date.today().year
    w_t = float(weights.get("title", 3.0))
    w_a = float(weights.get("abstract", 1.0))
    w_c = float(weights.get("cited", 1.0))
    w_r = float(weights.get("recency", 1.0))
    score = w_t * _hits(p.title, groups) + w_a * _hits(p.abstract or p.tldr, groups)
    score += w_c * math.log1p(max(0, p.cited_by or 0))
    if p.year:
        rec = max(0.0, 1.0 - (now_year - p.year) / 10.0)
        score += w_r * rec * 3.0      # 當年＝3 分、10 年前＝0 分
    return round(score, 3)


def _is_major(publisher: str) -> bool:
    pub = (publisher or "").lower()
    return any(m in pub for m in MAJOR_PUBLISHERS)


def evaluate(papers: list, settings: Optional[dict] = None) -> list:
    from .config import DEFAULTS
    s = dict(DEFAULTS)
    s.update(settings or {})
    checks = dict(DEFAULTS["checks"])
    checks.update(s.get("checks") or {})
    types = set(s.get("types") or ["journal-article"])
    yf, yt = s.get("year_from"), s.get("year_to")
    classic = bool(s.get("classic"))
    lang = (s.get("lang") or "any").lower()
    weights = dict(DEFAULTS["weights"])
    weights.update(s.get("weights") or {})
    groups = _terms(s.get("keywords"))
    this_year = _dt.date.today().year

    for p in papers:
        ind = {}
        manual = p.source == "manual"
        if checks.get("doi"):
            if not p.doi:
                ind["doi"] = (False, "無 DOI")
            elif p.doi_ok is True:
                ind["doi"] = (True, "doi.org 可解析")
            elif p.doi_ok is False:
                ind["doi"] = (False, "doi.org 無法解析")
            else:
                ind["doi"] = (None, "有 DOI，尚未驗證")
        if checks.get("oa"):
            if p.oa_pdf_url or p.oa_status in ("gold", "green", "hybrid", "bronze", "diamond"):
                lic = f"，授權 {p.oa_license}" if p.oa_license else ""
                ind["oa"] = (True, f"OA（{p.oa_status or '有全文連結'}{lic}）")
            elif p.oa_status == "closed":
                ind["oa"] = (False, "非 OA")
            else:
                ind["oa"] = (None, "OA 狀態未知")
        if checks.get("type"):
            ind["type"] = (p.type in types, TYPE_LABELS.get(p.type, p.type or "未知"))
        if checks.get("year"):
            if classic:
                ind["year"] = (True, f"{p.year or '無年份'}（含經典不限年）")
            elif not p.year:
                ind["year"] = (None, "無出版年")
            else:
                lo = yf or 0
                hi = yt or this_year + 1
                ind["year"] = (lo <= p.year <= hi, f"{p.year}（範圍 {yf or '不限'}–{yt or '今'}）")
        if checks.get("issn"):
            if p.type not in ("journal-article", "proceedings-article"):
                ind["issn"] = (None, "非期刊，待人工確認")
            elif not p.issn:
                ind["issn"] = (None, "無 ISSN，待人工確認")
            elif p.doaj:
                ind["issn"] = (True, "ISSN＋DOAJ 收錄")
            elif _is_major(p.publisher):
                ind["issn"] = (True, f"ISSN＋主要出版社（{p.publisher}）")
            else:
                ind["issn"] = (None, "待人工確認")
        if checks.get("lang"):
            code = (p.language or ("zh" if p.is_chinese else ("ja" if has_kana(p.title) else ""))).lower()
            if lang in ("any", "", "不限"):
                ind["lang"] = (True, code or "未知")
            elif not code:
                ind["lang"] = (None, "語言未知")
            else:
                ind["lang"] = (code.startswith(lang), code)
        if manual:
            ind["manual"] = (None, "使用者自填，未經 API 驗證")
        p.relevance = relevance(p, groups, weights, this_year)
        ind["relevance"] = (None, f"{p.relevance:.2f}")
        p.indicators = ind
    if (s.get("keep_mode") or "mark") == "strict":
        for p in papers:
            if not p.passes_all():
                p.include = False
    return papers
