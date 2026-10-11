"""FC LitGap Finder — 資料模型（規格第 4 節契約）。

Paper 是全系統共用的書目物件；所有 core 模組與 GUI 都只透過它交換資料。
"""
from __future__ import annotations

import json
import os
import re
import sys
import threading
from dataclasses import dataclass, field, asdict, fields
from typing import Any, Callable, Optional

VERSION = "1.0.0"
CREDIT = "NDHU自資系游豐兆製作"

# progress(msg, frac) 回呼型別；frac 為 0..1 或 None（不定）
Progress = Optional[Callable[[str, Optional[float]], None]]

# 指標代號（GUI 勾選清單用）
INDICATOR_NAMES = {
    "doi": "①有 DOI 且可解析",
    "oa": "②有合法 OA 電子全文",
    "type": "③文獻類型符合",
    "year": "⑥出版年範圍",
    "issn": "⑦期刊有 ISSN 且在 DOAJ／主要出版社",
    "lang": "⑩語言",
    "relevance": "⑫相關性分數",
}

# 文獻類型（內部統一用 Crossref 風格字串）
TYPE_LABELS = {
    "journal-article": "期刊論文",
    "proceedings-article": "會議論文",
    "book-chapter": "專書章節",
    "book": "專書",
    "patent": "專利",
    "webpage": "網頁",
    "posted-content": "預印本",
    "dissertation": "學位論文",
    "report": "報告",
    "other": "其他",
}

DEFAULT_CHAPTERS = ["1 緒論", "2.1 文獻回顧", "2.2", "3 方法", "4 結果", "5 討論", "附錄", "未分類"]


class Cancelled(Exception):
    """長作業被 cancel_event 中止時拋出。"""


def check_cancel(cancel_event: Optional[threading.Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise Cancelled("使用者已取消")


def report(progress: Progress, msg: str, frac: Optional[float] = None) -> None:
    if progress is not None:
        try:
            progress(msg, frac)
        except Exception:
            pass


@dataclass
class Paper:
    doi: str = ""                      # 無 https://doi.org/ 前綴；保留 API 原大小寫，去重一律用 doi_key（小寫）
    title: str = ""
    authors: list = field(default_factory=list)   # [{"family": "", "given": ""}]；團體作者 {"family": "名稱", "given": ""}
    year: Optional[int] = None
    container: str = ""                # 刊名／書名／會議名／網站名
    volume: str = ""
    issue: str = ""
    pages: str = ""                    # "13-19"
    article_number: str = ""           # "e00923"、"37"
    type: str = "journal-article"
    issn: list = field(default_factory=list)
    publisher: str = ""
    language: str = ""                 # "en"、"zh"…
    abstract: str = ""
    cited_by: int = 0
    oa_status: str = ""                # gold/green/hybrid/bronze/closed/""
    oa_pdf_url: str = ""
    oa_license: str = ""
    source: str = "crossref"           # crossref | openalex | manual
    doaj: Optional[bool] = None
    indicators: dict = field(default_factory=dict)   # name -> (bool|None, note)
    relevance: float = 0.0
    chapter: str = "未分類"
    include: bool = True
    order: int = 0
    # ---- 以下為擴充欄位（契約外，皆有預設值，GUI 可忽略） ----
    url: str = ""                      # 網頁／專利網址或 doi.org 連結
    editors: list = field(default_factory=list)     # 專書章節用 [{"family","given"}]
    book_title: str = ""               # 專書章節所屬書名
    edition: str = ""
    patent_number: str = ""
    patent_office: str = ""
    month_day: str = ""                # 網頁發布日期「Month D」
    accessed: str = ""                 # 網頁擷取日期
    tldr: str = ""                     # Semantic Scholar TLDR
    oa_pdf_alts: list = field(default_factory=list)  # 其他合法 OA PDF 位置（Unpaywall oa_locations），主連結失敗時依序嘗試
    doi_ok: Optional[bool] = None      # doi.org HEAD 驗證結果
    openalex_id: str = ""
    note: str = ""                     # 備註（例如手動加入附註）

    # ------------------------------------------------------------------
    @property
    def doi_url(self) -> str:
        return f"https://doi.org/{self.doi}" if self.doi else ""

    @property
    def key(self) -> str:
        return doi_key(self.doi) if self.doi else "t:" + re.sub(r"\W+", "", (self.title or "").lower())[:80] + f":{self.year}"

    @property
    def first_author_family(self) -> str:
        if self.authors:
            return (self.authors[0].get("family") or self.authors[0].get("given") or "").strip()
        return ""

    @property
    def is_chinese(self) -> bool:
        """中文文獻（套中文 APA、排在中文區）。含日文假名或語言標為非中文者不算。"""
        if has_kana(self.title):
            return False
        lang = (self.language or "").lower()
        if lang.startswith("zh"):
            return True
        if lang and not lang.startswith("zh"):
            return False
        return has_cjk(self.title) or has_cjk(self.first_author_family)

    def indicator_summary(self) -> str:
        """結果表「指標摘要」欄：✔/✘/? 串。"""
        order = ["doi", "oa", "type", "year", "issn", "lang"]
        marks = []
        for k in order:
            if k in self.indicators:
                v = self.indicators[k][0] if isinstance(self.indicators[k], (list, tuple)) else self.indicators[k]
                marks.append("✔" if v is True else ("✘" if v is False else "?"))
        return "".join(marks)

    def passes_all(self) -> bool:
        for k, v in self.indicators.items():
            if k == "relevance":
                continue
            val = v[0] if isinstance(v, (list, tuple)) else v
            if val is False:
                return False
        return True

    def to_dict(self) -> dict:
        d = asdict(self)
        d["indicators"] = {k: list(v) if isinstance(v, tuple) else v for k, v in self.indicators.items()}
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Paper":
        names = {f.name for f in fields(cls)}
        kw = {k: v for k, v in d.items() if k in names}
        p = cls(**kw)
        p.indicators = {k: tuple(v) if isinstance(v, list) else v for k, v in (p.indicators or {}).items()}
        p.doi = normalize_doi(p.doi)
        return p


def has_cjk(s: str) -> bool:
    return any("一" <= ch <= "鿿" or "㐀" <= ch <= "䶿" for ch in (s or ""))


def has_kana(s: str) -> bool:
    return any("\u3040" <= ch <= "\u30ff" for ch in (s or ""))


def normalize_doi(doi: str) -> str:
    if not doi:
        return ""
    d = doi.strip()
    for pre in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/", "http://dx.doi.org/", "doi:", "DOI:", "DOI "):
        if d.lower().startswith(pre.lower()):
            d = d[len(pre):]
    return d.strip()


def doi_key(doi: str) -> str:
    """去重用：DOI 一律小寫比對。"""
    return normalize_doi(doi).lower()


def app_base_dir() -> str:
    """原始碼執行時為 lit-gap-finder\\；PyInstaller onefile 時為 _MEIPASS。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS  # type: ignore[attr-defined]
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_samples(path: Optional[str] = None) -> list:
    """讀 selftest/samples.json；找不到時退回內建 core/sample_papers.py。"""
    cands = [path] if path else []
    cands.append(os.path.join(app_base_dir(), "selftest", "samples.json"))
    for c in cands:
        if c and os.path.isfile(c):
            with open(c, encoding="utf-8") as f:
                data = json.load(f)
            items = data["papers"] if isinstance(data, dict) else data
            return [Paper.from_dict(x) for x in items]
    from . import sample_papers
    return [Paper.from_dict(x) for x in sample_papers.SAMPLES]
