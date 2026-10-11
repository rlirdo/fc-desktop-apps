"""補充資料：Unpaywall（OA、DOAJ）、Semantic Scholar（被引、TLDR；≤1 req/s，失敗即降級）、doi.org HEAD 驗證。

DOAJ 判定：Unpaywall `journal_is_in_doaj`，其次 OpenAlex `source.is_in_doaj`（規格限定只呼叫第 2 節 API）。
"""
from __future__ import annotations

import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional
from urllib.parse import quote

from . import cache
from .config import load as load_config
from .models import Paper, Progress, check_cancel, report, doi_key, Cancelled
from .net import get_json, request, ApiError

UNPAYWALL = "https://api.unpaywall.org/v2/"
S2 = "https://api.semanticscholar.org/graph/v1/paper/DOI:"

# 最近一次 enrich 的狀態（GUI 狀態列可顯示）
status: dict = {}


def unpaywall(doi: str, email: str, cancel_event=None) -> Optional[dict]:
    key = f"q:unpaywall:{doi_key(doi)}"
    hit = cache.get(key)
    if hit is not None:
        return hit or None
    try:
        data = get_json(UNPAYWALL + quote(doi, safe="/()"), params={"email": email}, cancel_event=cancel_event)
    except ApiError as e:
        if e.status == 404:
            cache.set(key, {})
            return None
        raise
    cache.set(key, data)
    return data


def apply_unpaywall(p: Paper, d: dict) -> None:
    best = d.get("best_oa_location") or {}
    p.oa_status = d.get("oa_status") or p.oa_status or ("closed" if d.get("is_oa") is False else "")
    if best.get("url_for_pdf"):
        p.oa_pdf_url = best["url_for_pdf"]
    elif not d.get("is_oa"):
        p.oa_pdf_url = ""
    p.oa_license = best.get("license") or p.oa_license or ""
    alts = []
    for loc in d.get("oa_locations") or []:
        u = loc.get("url_for_pdf")
        if u and u != p.oa_pdf_url and u not in alts:
            alts.append(u)
    p.oa_pdf_alts = alts
    if d.get("journal_is_in_doaj") is not None:
        p.doaj = bool(d.get("journal_is_in_doaj")) or bool(p.doaj)
    if not p.issn and d.get("journal_issn_l"):
        p.issn = [d["journal_issn_l"]]
    if not p.publisher and d.get("publisher"):
        p.publisher = d["publisher"]


def doi_resolves(doi: str, cancel_event=None) -> Optional[bool]:
    key = f"q:doiorg:{doi_key(doi)}"
    hit = cache.get(key)
    if hit is not None:
        return hit.get("ok")
    try:
        r = request("HEAD", "https://doi.org/" + quote(doi, safe="/()"), allow_redirects=False,
                    ok_status=(200, 301, 302, 303, 307, 308), cancel_event=cancel_event, retries=2)
        ok = True
        loc = r.headers.get("Location", "")
    except ApiError as e:
        if e.status == 404:
            ok, loc = False, ""
        else:
            return None
    cache.set(key, {"ok": ok, "location": loc})
    return ok


def s2_lookup(doi: str, cancel_event=None) -> Optional[dict]:
    key = f"q:s2:{doi_key(doi)}"
    hit = cache.get(key)
    if hit is not None:
        return hit or None
    try:
        data = get_json(S2 + quote(doi, safe="/()"), params={"fields": "citationCount,tldr,abstract,journal"},
                        cancel_event=cancel_event, retries=1)
    except ApiError as e:
        if e.status == 404:
            cache.set(key, {})
            return None
        raise
    cache.set(key, data)
    return data


def apply_s2(p: Paper, d: dict) -> None:
    p.cited_by = max(p.cited_by or 0, int(d.get("citationCount") or 0))
    tl = d.get("tldr") or {}
    if tl.get("text"):
        p.tldr = tl["text"]
    if not p.abstract and d.get("abstract"):
        p.abstract = d["abstract"]
    # Crossref 與 OpenAlex 都沒有頁碼時，以 S2 journal.pages 補（只補頁碼；S2 卷號偶有錯誤，不採用）
    pg = ((d.get("journal") or {}).get("pages") or "").strip()
    if pg and not p.pages and not p.article_number and p.type in ("journal-article", "proceedings-article", "book-chapter"):
        pg = re.sub(r"\s*[-–]\s*", "-", pg)
        if re.fullmatch(r"[A-Za-z]?\d+-[A-Za-z]?\d+", pg):
            p.pages = pg
            p.note = (p.note + "；" if p.note else "") + "頁碼取自 Semantic Scholar"


def enrich(papers: list, email: str = "", progress: Progress = None, cancel_event=None,
           use_s2: Optional[bool] = None, s2_max: Optional[int] = None, verify_doi: bool = True) -> list:
    cfg = load_config()
    email = email or cfg.get("email") or ""
    use_s2 = cfg.get("s2_enabled", True) if use_s2 is None else use_s2
    s2_max = cfg.get("s2_max", 60) if s2_max is None else s2_max
    targets = [p for p in papers if p.doi]
    n = len(targets)
    status.clear()
    status.update({"unpaywall_fail": 0, "doi_fail": 0, "s2": "未使用", "errors": []})
    report(progress, f"補充 OA／DOI 驗證：{n} 筆", 0.0)
    done = [0]
    lock = threading.Lock()

    def work(p: Paper):
        check_cancel(cancel_event)
        if email:
            try:
                d = unpaywall(p.doi, email, cancel_event)
                if d:
                    apply_unpaywall(p, d)
            except ApiError as e:
                with lock:
                    status["unpaywall_fail"] += 1
                    if len(status["errors"]) < 5:
                        status["errors"].append(f"Unpaywall：{e}")
        if verify_doi:
            ok = doi_resolves(p.doi, cancel_event)
            p.doi_ok = ok
            if ok is None:
                with lock:
                    status["doi_fail"] += 1
        with lock:
            done[0] += 1
            report(progress, f"Unpaywall／doi.org {done[0]}/{n}", 0.8 * done[0] / max(1, n))

    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = [ex.submit(work, p) for p in targets]
        for f in as_completed(futs):
            try:
                f.result()
            except Cancelled:
                for g in futs:
                    g.cancel()
                raise

    # Semantic Scholar：依被引排序取前 s2_max 筆，≤1 req/s；連續失敗 2 次即降級停止
    if use_s2 and targets:
        status["s2"] = "使用中"
        s2_targets = sorted(targets, key=lambda p: -(p.cited_by or 0))[: max(0, s2_max)]
        # 缺頁碼／文章號的期刊論文一律查（補頁碼），不受 s2_max 限制
        s2_targets += [p for p in targets if p not in s2_targets and not p.pages and not p.article_number
                       and p.type == "journal-article"]
        fails = 0
        for i, p in enumerate(s2_targets):
            check_cancel(cancel_event)
            report(progress, f"Semantic Scholar 被引／TLDR {i + 1}/{len(s2_targets)}",
                   0.8 + 0.2 * (i + 1) / max(1, len(s2_targets)))
            try:
                d = s2_lookup(p.doi, cancel_event)
                fails = 0
            except ApiError as e:
                fails += 1
                if fails >= 2:
                    status["s2"] = f"已降級（{e}）"
                    status["errors"].append(f"Semantic Scholar 已降級：{e}")
                    break
                continue
            if d:
                apply_s2(p, d)
        else:
            status["s2"] = "完成"
    report(progress, "補充完成" + (f"（{'；'.join(status['errors'][:3])}）" if status["errors"] else ""), 1.0)
    return papers
