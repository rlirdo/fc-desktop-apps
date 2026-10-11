"""離線假資料版介面（規格第 4 節）——讓 GUI 在真模組完成前先開發。

函式名稱與簽章與真模組完全一致：
    search / enrich / evaluate / gap_matrix / format_apa7 / in_text / sort_key /
    export_all / download_oa / cache_get / cache_set / cache_clear / load / save / DEFAULTS
所有長作業接受 progress(msg, frac) 與 cancel_event（threading.Event，可省略）。
真模組完成後，GUI 改為：
    from core.search import search; from core.enrich import enrich; from core.indicators import evaluate
    from core.gap import gap_matrix; from core.apa import format_apa7, in_text, sort_key
    from core.export import export_all; from core.oa_download import download_oa
    from core import cache, config
"""
from __future__ import annotations

import copy
import json
import os
import time
from typing import Optional

from .models import Paper, Progress, load_samples, report, check_cancel, DEFAULT_CHAPTERS

from .config import DEFAULTS  # noqa: E402  與真模組同一份預設值

_CFG = copy.deepcopy(DEFAULTS)
_CACHE: dict = {}


def _sleep(cancel_event, sec=0.05):
    check_cancel(cancel_event)
    time.sleep(sec)


# ---------------- search / enrich / indicators ----------------
def search(keywords: list, year_from: Optional[int] = None, year_to: Optional[int] = None, lang: str = "any",
           limit: int = 50, progress: Progress = None, cancel_event=None) -> list:
    papers = load_samples()
    out = []
    for i, p in enumerate(papers):
        _sleep(cancel_event)
        report(progress, f"（假資料）檢索第 {i + 1}/{len(papers)} 筆", (i + 1) / len(papers))
        if year_from and p.year and p.year < year_from:
            continue
        if year_to and p.year and p.year > year_to:
            continue
        p.order = len(out)
        out.append(p)
    return out[:limit] if limit else out


def enrich(papers: list, email: str = "", progress: Progress = None, cancel_event=None) -> list:
    for i, p in enumerate(papers):
        _sleep(cancel_event)
        report(progress, f"（假資料）補充 OA／被引 {i + 1}/{len(papers)}", (i + 1) / max(1, len(papers)))
        if p.doi:
            p.doi_ok = True
            if p.oa_status in ("gold", "hybrid", "green", "bronze") and not p.oa_pdf_url:
                p.oa_pdf_url = f"https://doi.org/{p.doi}"
                p.oa_license = "cc-by" if p.oa_status == "gold" else ""
            if p.doaj is None:
                p.doaj = p.oa_status == "gold"
    return papers


def evaluate(papers: list, settings: Optional[dict] = None) -> list:
    """settings 範例：{"checks":{"doi":True,...}, "types":["journal-article"], "year_from":2016, "year_to":None,
    "classic":False, "lang":"any", "weights":{...}, "keywords":[...]}"""
    s = settings or {}
    yf = s.get("year_from", 2016)
    for p in papers:
        ind = {}
        ind["doi"] = (bool(p.doi), "有 DOI" if p.doi else "無 DOI")
        ind["oa"] = (bool(p.oa_pdf_url) or p.oa_status in ("gold", "hybrid", "green", "bronze"), p.oa_status or "未知")
        ind["type"] = (p.type in s.get("types", ["journal-article"]), p.type)
        ind["year"] = (True if s.get("classic") else (p.year is not None and (not yf or p.year >= yf)), str(p.year))
        ind["issn"] = (True if p.issn and p.doaj else (None if p.issn else False), "DOAJ" if p.doaj else "待人工確認")
        ind["lang"] = (True, p.language or "未知")
        p.indicators = ind
        p.relevance = round(1.0 + (p.cited_by or 0) ** 0.1, 2)
    return papers


# ---------------- gap ----------------
def gap_matrix(groups: list, threshold: int = 5, progress: Progress = None, cancel_event=None) -> dict:
    from itertools import combinations
    samples = load_samples()
    pairs, triples = [], []
    combos2 = list(combinations(range(len(groups)), 2))
    combos3 = list(combinations(range(len(groups)), 3))
    total = max(1, len(combos2) + len(combos3))
    k = 0
    for (i, j) in combos2:
        _sleep(cancel_event)
        k += 1
        report(progress, f"（假資料）組合 {i + 1}×{j + 1}", k / total)
        n = (i * 7 + j * 3) % 40
        pairs.append({"groups": (i, j), "count": n, "gap": n < threshold, "query": "", "top": samples[: min(n, 10)]})
    for (i, j, m) in combos3:
        _sleep(cancel_event)
        k += 1
        report(progress, f"（假資料）組合 {i + 1}×{j + 1}×{m + 1}", k / total)
        n = (i + j + m) % 6
        triples.append({"groups": (i, j, m), "count": n, "gap": n < threshold, "query": "", "top": samples[: min(n, 10)]})
    trend = {y: 10 + (y % 7) * 3 for y in range(2016, 2026)}
    return {"pairs": pairs, "triples": triples, "trend": trend, "threshold": threshold, "groups": groups}


# ---------------- apa ----------------
def format_apa7(paper: Paper) -> str:
    au = ", ".join(f"{a.get('family','')}, {a.get('given','')[:1]}." for a in paper.authors) or "Anonymous"
    y = paper.year or "n.d."
    tail = f" https://doi.org/{paper.doi}" if paper.doi else ""
    return f"{au} ({y}). {paper.title}. {paper.container}, {paper.volume}({paper.issue}), {paper.pages}.{tail}"


def in_text(paper: Paper) -> str:
    return f"({paper.first_author_family}, {paper.year or 'n.d.'})"


def sort_key(paper: Paper):
    return (0 if paper.is_chinese else 1, paper.first_author_family.lower(), paper.year or 0)


# ---------------- export / OA ----------------
def export_all(papers: list, gap: Optional[dict], out_dir: str, name: str, options: Optional[dict] = None,
               progress: Progress = None, cancel_event=None) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{name}_查詢紀錄.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"stub": True, "n": len(papers)}, f, ensure_ascii=False, indent=2)
    report(progress, "（假資料）匯出完成", 1.0)
    return {path: True}


def download_oa(papers: list, out_dir: str, progress: Progress = None, cancel_event=None) -> list:
    rows = []
    for i, p in enumerate(papers):
        _sleep(cancel_event)
        report(progress, f"（假資料）OA 下載 {i + 1}/{len(papers)}", (i + 1) / max(1, len(papers)))
        rows.append({"doi": p.doi, "url": p.oa_pdf_url, "license": p.oa_license, "time": "", "status": "stub-skip", "file": ""})
    return rows


# ---------------- cache / config ----------------
def cache_get(key: str):
    return _CACHE.get(key)


def cache_set(key: str, value) -> None:
    _CACHE[key] = value


def cache_clear() -> int:
    n = len(_CACHE)
    _CACHE.clear()
    return n


def load() -> dict:
    return copy.deepcopy(_CFG)


def save(cfg: dict) -> None:
    _CFG.update(cfg)
