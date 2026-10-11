"""Crossref＋OpenAlex 檢索、合併去重（DOI 小寫）。

keywords：list[str]，每個元素為一組關鍵字（組內「;」分中英同義詞，組間為 AND）。
若某元素本身是 DOI（10.xxxx/...），直接以 DOI 取回該筆。
每筆 API 原始 JSON 均存快取（raw:crossref:<doi>／raw:openalex:<id>），作為「不杜撰書目」的證據。
"""
from __future__ import annotations

import datetime as _dt
import itertools
import re
from typing import Optional
from urllib.parse import quote

from . import cache
from .apa import clean_text
from .config import load as load_config
from .models import Paper, Progress, check_cancel, report, normalize_doi, doi_key, has_cjk
from .net import get_json, ApiError

CROSSREF = "https://api.crossref.org/works"
OPENALEX = "https://api.openalex.org/works"
DOI_RE = re.compile(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)?(10\.\d{4,9}/\S+)$", re.I)

# 最近一次檢索的錯誤（GUI 可顯示「部分來源失敗」）
last_errors: list = []


def split_groups(keywords) -> list:
    groups = []
    for g in keywords or []:
        items = g if isinstance(g, (list, tuple)) else re.split(r"[;；]", str(g))
        terms = [t.strip() for t in items if t and t.strip()]
        if terms:
            groups.append(terms)
    return groups


def openalex_query(groups: list) -> str:
    parts = []
    for g in groups:
        # 多字詞與中文詞一律加引號做片語比對（OpenAlex 會把中文拆成單字，不加引號會大量誤中）
        terms = [f'"{t}"' if (" " in t or "-" in t or has_cjk(t)) else t for t in g]
        parts.append("(" + " OR ".join(terms) + ")" if len(terms) > 1 else terms[0])
    return " AND ".join(parts)


# ---------------------------------------------------------------- 快取包裝
def cached_json(url: str, params: dict, mailto: str, cancel_event=None, use_cache: bool = True) -> dict:
    key = "q:" + url + "?" + "&".join(f"{k}={params[k]}" for k in sorted(params) if k not in ("mailto",))
    if use_cache:
        hit = cache.get(key)
        if hit is not None:
            return hit
    data = get_json(url, params=params, mailto=mailto, cancel_event=cancel_event)
    cache.set(key, data)
    return data


# ---------------------------------------------------------------- Crossref
def _cr_year(item: dict) -> Optional[int]:
    for k in ("published-print", "issued", "published-online", "published", "created"):
        dp = (item.get(k) or {}).get("date-parts") or []
        if dp and dp[0] and dp[0][0]:
            try:
                return int(dp[0][0])
            except (TypeError, ValueError):
                continue
    return None


def parse_crossref(item: dict) -> Paper:
    authors = []
    for a in item.get("author") or []:
        if a.get("family") or a.get("given"):
            authors.append({"family": clean_text(a.get("family", "")), "given": clean_text(a.get("given", ""))})
        elif a.get("name"):
            authors.append({"family": clean_text(a["name"]), "given": ""})
    editors = [{"family": clean_text(a.get("family", "")), "given": clean_text(a.get("given", ""))}
               for a in item.get("editor") or [] if a.get("family")]
    titles = item.get("title") or [""]
    title = clean_text(titles[0])
    sub = item.get("subtitle") or []
    if sub and sub[0] and clean_text(sub[0]).lower() not in title.lower():
        title = f"{title}: {clean_text(sub[0])}"
    ctr = item.get("container-title") or [""]
    container = clean_text(ctr[0]) if ctr else ""
    page = (item.get("page") or "").strip()
    art = (item.get("article-number") or "").strip()
    pages = ""
    if page:
        if re.search(r"\w\s*[-–]\s*\w", page):
            pages = page
        elif not art:
            art = page          # 單一頁碼／e 編號＝文章編號（Heliyon e00923、MDPI 37）
        elif page != art:
            pages = page
    t = item.get("type") or "journal-article"
    book_title = ""
    if t == "book-chapter":
        book_title = container
    lang = (item.get("language") or "").lower()
    p = Paper(
        doi=normalize_doi(item.get("DOI", "")), title=title, authors=authors, year=_cr_year(item),
        container=container, volume=str(item.get("volume") or ""), issue=str(item.get("issue") or ""),
        pages=pages, article_number=art, type=t, issn=list(dict.fromkeys(item.get("ISSN") or [])),
        publisher=clean_text(item.get("publisher", "")), language=lang,
        abstract=clean_text(item.get("abstract", "")), cited_by=int(item.get("is-referenced-by-count") or 0),
        source="crossref", editors=editors, book_title=book_title,
    )
    if not p.language:
        p.language = "zh" if has_cjk(p.title) else ""
    return p


def crossref_by_doi(doi: str, mailto: str, cancel_event=None) -> Optional[Paper]:
    try:
        data = cached_json(f"{CROSSREF}/{quote(doi, safe='/()')}", {}, mailto, cancel_event)
    except ApiError as e:
        if e.status == 404:
            return None
        raise
    item = data.get("message") or {}
    cache.set(f"raw:crossref:{doi_key(item.get('DOI', doi))}", item)
    return parse_crossref(item)


def crossref_search(query: str, year_from, year_to, rows: int, mailto: str, cancel_event=None, title_only=False) -> list:
    params = {"rows": max(1, min(rows, 100)), "mailto": mailto,
              "select": "DOI,title,subtitle,author,editor,issued,published-print,published-online,container-title,"
                        "volume,issue,page,article-number,type,ISSN,publisher,abstract,is-referenced-by-count"}
    params["query.title" if title_only else "query.bibliographic"] = query
    filt = []
    if year_from:
        filt.append(f"from-pub-date:{year_from}")
    if year_to:
        filt.append(f"until-pub-date:{year_to}")
    if filt:
        params["filter"] = ",".join(filt)
    data = cached_json(CROSSREF, params, mailto, cancel_event)
    items = (data.get("message") or {}).get("items") or []
    out = []
    raws = []
    for it in items:
        out.append(parse_crossref(it))
        if it.get("DOI"):
            raws.append((f"raw:crossref:{doi_key(it['DOI'])}", it))
    cache.set_many(raws)
    return out


# ---------------------------------------------------------------- OpenAlex
def _abstract_from_inverted(inv: Optional[dict]) -> str:
    if not inv:
        return ""
    pos = []
    for word, idxs in inv.items():
        for i in idxs:
            pos.append((i, word))
    pos.sort()
    return " ".join(w for _, w in pos)


_OA_TYPE = {"article": "journal-article", "book-chapter": "book-chapter", "book": "book", "preprint": "posted-content",
            "dissertation": "dissertation", "report": "report", "review": "journal-article", "letter": "journal-article",
            "editorial": "journal-article", "proceedings-article": "proceedings-article"}


def parse_openalex(w: dict) -> Paper:
    authors = []
    for au in w.get("authorships") or []:
        # raw_author_name＝論文上的署名；author.display_name 為 OpenAlex 歸併後的人名，可能與原文不符
        name = (au.get("raw_author_name") or (au.get("author") or {}).get("display_name") or "").strip()
        if "," in name and not has_cjk(name):
            fam, _, giv = name.partition(",")
            authors.append({"family": fam.strip(), "given": giv.strip()})
            continue
        if not name:
            continue
        if has_cjk(name):
            authors.append({"family": name, "given": ""})
        else:
            parts = name.split()
            authors.append({"family": parts[-1], "given": " ".join(parts[:-1])} if len(parts) > 1 else {"family": name, "given": ""})
    for a in authors:      # 全大寫署名（B. H. YOO）轉為一般大小寫
        for k in ("family", "given"):
            v = a[k]
            if v and v.isupper() and len(v.replace(".", "").replace(" ", "")) > 2 and not has_cjk(v):
                a[k] = v.title()
    repo_only = False
    loc = w.get("primary_location") or {}
    src = loc.get("source") or {}
    if src.get("type") not in (None, "journal", "conference", "book series", "ebook platform"):
        # primary 是 PubMed 等典藏庫時，改用期刊來源
        for alt in w.get("locations") or []:
            s2 = alt.get("source") or {}
            if s2.get("type") in ("journal", "conference"):
                src = s2
                break
        else:
            if src.get("type") == "repository":
                src = {}
                repo_only = True
    bib = w.get("biblio") or {}
    fp, lp = bib.get("first_page") or "", bib.get("last_page") or ""
    pages, art = "", ""
    if fp and lp and fp.isdigit() and lp.isdigit() and len(lp) < len(fp):
        lp = fp[: len(fp) - len(lp)] + lp        # 175–84 → 175–184
    if fp and lp and fp != lp:
        pages = f"{fp}-{lp}"
    elif fp:
        art = fp
    t = w.get("type_crossref") or _OA_TYPE.get(w.get("type") or "", "journal-article")
    if t == "journal-article" and (src.get("type") == "conference"):
        t = "proceedings-article"
    oa = w.get("open_access") or {}
    best = w.get("best_oa_location") or {}
    pdf = best.get("pdf_url") or ""
    p = Paper(
        doi=normalize_doi(w.get("doi") or ""), title=clean_text(w.get("title") or w.get("display_name") or ""),
        authors=authors, year=w.get("publication_year"), container=clean_text(src.get("display_name") or ""),
        volume=str(bib.get("volume") or ""), issue=str(bib.get("issue") or ""), pages=pages, article_number=art,
        type=t, issn=list(src.get("issn") or []), publisher=clean_text(src.get("host_organization_name") or ""),
        language=(w.get("language") or ""), abstract=_abstract_from_inverted(w.get("abstract_inverted_index")),
        cited_by=int(w.get("cited_by_count") or 0), oa_status=(oa.get("oa_status") or ""),
        oa_pdf_url=pdf, oa_license=(best.get("license") or ""), source="openalex",
        doaj=src.get("is_in_doaj") if src else None, openalex_id=(w.get("id") or "").rsplit("/", 1)[-1],
    )
    pmid = ((w.get("ids") or {}).get("pmid") or "")
    if pmid and not p.doi:
        p.url = pmid
    if repo_only:
        p.note = "OpenAlex 僅有典藏庫紀錄（" + (pmid.rsplit("/", 1)[-1] and "PMID " + pmid.rsplit("/", 1)[-1] if pmid else "無 PMID") + "），刊名缺，待人工補正"
    return p


OA_SELECT = ("id,doi,ids,title,display_name,publication_year,type,type_crossref,language,authorships,primary_location,"
             "biblio,open_access,best_oa_location,locations,cited_by_count,abstract_inverted_index")


def openalex_search(query: str, year_from, year_to, lang: str, per_page: int, mailto: str, cancel_event=None,
                    title_only: bool = False) -> list:
    params = {"per_page": max(1, min(per_page, 200)), "mailto": mailto, "select": OA_SELECT}
    filt = []
    if title_only:
        filt.append(f"title.search:{query.replace(',', ' ')}")
    else:
        params["search"] = query
    if year_from or year_to:
        filt.append(f"publication_year:{year_from or ''}-{year_to or ''}")
    if lang in ("en", "zh"):
        filt.append(f"language:{lang}")
    if filt:
        params["filter"] = ",".join(filt)
    data = cached_json(OPENALEX, params, mailto, cancel_event)
    out, raws = [], []
    for w in data.get("results") or []:
        out.append(parse_openalex(w))
        raws.append((f"raw:openalex:{(w.get('id') or '').rsplit('/', 1)[-1]}", w))
    cache.set_many(raws)
    return out


def openalex_by_doi(doi: str, mailto: str, cancel_event=None) -> Optional[Paper]:
    try:
        w = cached_json(f"{OPENALEX}/https://doi.org/{doi}", {"mailto": mailto}, mailto, cancel_event)
    except ApiError as e:
        if e.status == 404:
            return None
        raise
    cache.set(f"raw:openalex:{(w.get('id') or '').rsplit('/', 1)[-1]}", w)
    return parse_openalex(w)


# ---------------------------------------------------------------- 合併
_PREFER_CR = ("title", "authors", "year", "container", "volume", "issue", "pages", "article_number", "type", "issn",
              "publisher", "editors", "book_title")


def merge(a: Paper, b: Paper) -> Paper:
    """a 優先（Crossref），以 b（OpenAlex）補空欄；被引取大者、OA 與 DOAJ 取 OpenAlex。"""
    if a.source != "crossref" and b.source == "crossref":
        a, b = b, a
    for f in _PREFER_CR:
        if not getattr(a, f) and getattr(b, f):
            setattr(a, f, getattr(b, f))
    if not a.abstract and b.abstract:
        a.abstract = b.abstract
    if not a.language and b.language:
        a.language = b.language
    a.cited_by = max(a.cited_by or 0, b.cited_by or 0)
    for f in ("oa_status", "oa_pdf_url", "oa_license", "openalex_id"):
        if getattr(b, f) and not getattr(a, f):
            setattr(a, f, getattr(b, f))
    if a.doaj is None and b.doaj is not None:
        a.doaj = b.doaj
    if not a.doi and b.doi:
        a.doi = b.doi
    a.note = (a.note + " " if a.note else "") + ("Crossref＋OpenAlex" if b.source == "openalex" else "")
    a.note = a.note.strip()
    return a


def dedupe(papers: list) -> list:
    by_key: dict = {}
    order = []
    title_index: dict = {}
    for p in papers:
        tkey = re.sub(r"\W+", "", (p.title or "").lower())[:80]
        k = doi_key(p.doi) if p.doi else None
        if k and k in by_key:
            by_key[k] = merge(by_key[k], p)
            continue
        tk = (tkey, p.year)
        if tkey and tk in title_index:
            kk = title_index[tk]
            by_key[kk] = merge(by_key[kk], p)
            continue
        kk = k or f"t:{tkey}:{p.year}"
        by_key[kk] = p
        order.append(kk)
        if tkey:
            title_index[tk] = kk
    return [by_key[k] for k in order]


def _hits_all(p: Paper, groups: list) -> bool:
    text = f"{p.title} {p.abstract}".lower()
    return all(any(t.lower() in text for t in g) for g in groups)


# ---------------------------------------------------------------- 主函式
def search(keywords: list, year_from: Optional[int] = None, year_to: Optional[int] = None, lang: str = "any",
           limit: int = 50, progress: Progress = None, cancel_event=None, mailto: Optional[str] = None,
           title_only: bool = False) -> list:
    """回傳合併去重後的 Paper 清單（不含指標，請再呼叫 enrich／evaluate）。"""
    global last_errors
    last_errors = []
    mailto = mailto or load_config().get("mailto") or ""
    items = [str(k).strip() for k in (keywords or []) if str(k).strip()]
    dois = []
    rest = []
    for k in items:
        m = DOI_RE.match(k)
        (dois if m else rest).append(normalize_doi(m.group(1)) if m else k)
    found: list = []
    steps = len(dois) * 2 + (2 if rest else 0)
    step = 0
    for d in dois:
        check_cancel(cancel_event)
        for fn, name in ((crossref_by_doi, "Crossref"), (openalex_by_doi, "OpenAlex")):
            step += 1
            report(progress, f"以 DOI 取回 {d}（{name}）", step / max(1, steps))
            try:
                p = fn(d, mailto, cancel_event)
                if p:
                    found.append(p)
            except ApiError as e:
                last_errors.append(f"{name}：{e}")
    if rest:
        groups = split_groups(rest)
        # Crossref 不支援布林：以各組英文（無則第一個）同義詞組合查詢，最多 4 種組合
        def pick(g):
            en = [t for t in g if not has_cjk(t)]
            return en or g[:1]
        combos = list(itertools.islice(itertools.product(*[pick(g) for g in groups]), 4))
        per = max(10, limit // max(1, len(combos)))
        if len(groups) > 1:
            per = min(100, per * 3)   # 之後會做 AND 篩選，先多取
        step += 1
        report(progress, f"Crossref 檢索（{len(combos)} 組查詢）", step / max(1, steps))
        for c in combos:
            check_cancel(cancel_event)
            try:
                got = crossref_search(" ".join(c), year_from, year_to, per, mailto, cancel_event, title_only)
                if len(groups) > 1:
                    # Crossref 為寬鬆相關性比對（無 AND）：多組時要求每組至少一個同義詞出現在題名／摘要
                    got = [p for p in got if _hits_all(p, groups)]
                found += got
            except ApiError as e:
                last_errors.append(f"Crossref：{e}")
                break
        step += 1
        q = openalex_query(groups)
        report(progress, f"OpenAlex 檢索：{q}", step / max(1, steps))
        try:
            found += openalex_search(q, year_from, year_to, lang, limit, mailto, cancel_event, title_only)
        except ApiError as e:
            last_errors.append(f"OpenAlex：{e}")
    merged = dedupe(found)
    # Crossref 無年份篩選語言：此處補做年份過濾
    out = []
    for p in merged:
        if year_from and p.year and p.year < year_from:
            continue
        if year_to and p.year and p.year > year_to:
            continue
        if lang in ("en", "zh") and p.language and not p.language.lower().startswith(lang):
            continue
        out.append(p)
    for i, p in enumerate(out):
        p.order = i
    if last_errors and not out:
        raise ApiError("所有來源皆失敗：" + "；".join(last_errors))
    report(progress, f"檢索完成：{len(out)} 筆（已去重）" + (f"；部分來源失敗：{'；'.join(last_errors)}" if last_errors else ""), 1.0)
    return out


def find_by_title(title: str, mailto: str, year: Optional[int] = None, cancel_event=None) -> Optional[Paper]:
    """回測用：以題名找回單篇（Crossref query.title＋OpenAlex title.search，比對題名相似度）。"""
    def norm(s):
        return re.sub(r"[^a-z0-9]", "", clean_text(s).lower())
    target = norm(title)
    cands = []
    try:
        cands += crossref_search(title, None, None, 5, mailto, cancel_event, title_only=True)
    except ApiError:
        pass
    try:
        cands += openalex_search(re.sub(r"[^\w\s-]", " ", title), None, None, "any", 5, mailto, cancel_event, title_only=True)
    except ApiError:
        pass
    best = None
    for c in dedupe(cands):
        n = norm(c.title)
        if not n:
            continue
        if n == target or (len(target) > 20 and (target in n or n in target)):
            if year and c.year and abs(c.year - year) > 1:
                continue
            if best is None or (best.source != "crossref" and c.source == "crossref"):
                best = c
    return best
