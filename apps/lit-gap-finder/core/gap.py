"""缺口矩陣：多組關鍵字兩兩／三三組合在 OpenAlex 的命中篇數，< 門檻者標為缺口。

每個組合一次請求：works?search=(A1 OR A2) AND (B1 OR B2)&per_page=10 → meta.count＋前 10 篇（相關性排序）。
趨勢：每組關鍵字近 10 年逐年篇數（group_by=publication_year），trend 為第 1 組（主題組）。
"""
from __future__ import annotations

import datetime as _dt
import time
from itertools import combinations
from typing import Optional

from .config import load as load_config
from .models import Progress, check_cancel, report
from .net import ApiError
from .search import OPENALEX, OA_SELECT, cached_json, openalex_query, parse_openalex, split_groups

last_errors: list = []


def _count_and_top(groups: list, mailto: str, top_n: int, cancel_event=None) -> tuple:
    q = openalex_query(groups)
    params = {"search": q, "per_page": max(1, top_n), "mailto": mailto, "select": OA_SELECT}
    data = cached_json(OPENALEX, params, mailto, cancel_event)
    count = int((data.get("meta") or {}).get("count") or 0)
    top = [parse_openalex(w) for w in (data.get("results") or [])][:top_n]
    return q, count, top


def yearly_counts(groups: list, mailto: str, years: int = 10, cancel_event=None) -> dict:
    last = _dt.date.today().year - 1          # 只取完整年度（今年資料不完整且 OpenAlex 當年度常有異常值）
    y0 = last - years + 1
    params = {"search": openalex_query(groups), "group_by": "publication_year", "mailto": mailto,
              "filter": f"publication_year:{y0}-{last}"}
    data = cached_json(OPENALEX, params, mailto, cancel_event)
    out = {y: 0 for y in range(y0, last + 1)}
    for g in data.get("group_by") or []:
        try:
            y = int(g.get("key"))
        except (TypeError, ValueError):
            continue
        if y in out:
            out[y] = int(g.get("count") or 0)
    return out


def find_outliers(trend: dict, factor: float = 3.0) -> list:
    """某年篇數 > 其他年份中位數×factor（且多出 100 篇以上）視為異常值（如資料庫批次匯入）。"""
    ys = sorted(trend)
    out = []
    for y in ys:
        others = sorted(v for k, v in trend.items() if k != y)
        if len(others) < 3:
            continue
        med = others[len(others) // 2]
        if trend[y] > max(med * factor, med + 100):
            out.append(y)
    return out


def gap_matrix(groups: list, threshold: int = 5, progress: Progress = None, cancel_event=None,
               mailto: Optional[str] = None, top_n: int = 10, triples: bool = True,
               trend_groups: Optional[str] = None) -> dict:
    """groups：list[list[str]] 或 list["a; b"]。回傳 {"pairs","triples","trend","trend_by_group","singles",...}"""
    global last_errors
    last_errors = []
    t0 = time.time()
    cfg = load_config()
    mailto = mailto or cfg.get("mailto") or ""
    trend_groups = trend_groups or cfg.get("gap_trend_groups", "first")
    G = split_groups(groups)
    combos2 = list(combinations(range(len(G)), 2))
    combos3 = list(combinations(range(len(G)), 3)) if triples else []
    total = len(G) + len(combos2) + len(combos3) + (len(G) if trend_groups == "all" else 1)
    k = 0
    budget_out = [False]
    singles = []
    for i, g in enumerate(G):
        check_cancel(cancel_event)
        k += 1
        report(progress, f"單組命中：{' / '.join(g)}", k / total)
        if budget_out[0]:
            singles.append({"group": i, "count": None, "query": openalex_query([g]), "error": "OpenAlex 額度不足，未查詢"})
            continue
        try:
            q, n, _ = _count_and_top([g], mailto, 1, cancel_event)
            singles.append({"group": i, "count": n, "query": q})
        except ApiError as e:
            last_errors.append(str(e))
            if getattr(e, "budget", False):
                budget_out[0] = True
            singles.append({"group": i, "count": None, "query": openalex_query([g]), "error": str(e)})

    def run(idx):
        sub = [G[i] for i in idx]
        if budget_out[0]:
            return {"groups": tuple(idx), "count": None, "gap": None, "query": openalex_query(sub), "top": [],
                    "error": "OpenAlex 額度不足，未查詢"}
        try:
            q, n, top = _count_and_top(sub, mailto, top_n, cancel_event)
            return {"groups": tuple(idx), "count": n, "gap": n < threshold, "query": q, "top": top}
        except ApiError as e:
            last_errors.append(str(e))
            if getattr(e, "budget", False):
                budget_out[0] = True
            return {"groups": tuple(idx), "count": None, "gap": None, "query": openalex_query(sub), "top": [],
                    "error": str(e)}

    pairs = []
    for idx in combos2:
        check_cancel(cancel_event)
        k += 1
        report(progress, "兩兩組合 " + "×".join(str(i + 1) for i in idx), k / total)
        pairs.append(run(idx))
    trip = []
    for idx in combos3:
        check_cancel(cancel_event)
        k += 1
        report(progress, "三三組合 " + "×".join(str(i + 1) for i in idx), k / total)
        trip.append(run(idx))
    trend_by_group = {}
    trend_status = {i: "skipped" for i in range(len(G))}
    for i, g in enumerate(G):
        if (trend_groups == "first" and i > 0) or budget_out[0]:
            break
        check_cancel(cancel_event)
        k += 1
        report(progress, f"年度趨勢：第 {i + 1} 組", k / total)
        try:
            trend_by_group[i] = yearly_counts([g], mailto, 10, cancel_event)
            trend_status[i] = "ok"
        except ApiError as e:
            last_errors.append(str(e))
            trend_status[i] = "failed"
            if getattr(e, "budget", False):
                budget_out[0] = True
    if budget_out[0]:
        for i in trend_status:
            if trend_status[i] == "skipped" and (trend_groups == "all" or i == 0):
                trend_status[i] = "failed"
    trend = trend_by_group.get(0, {})
    outliers = {i: find_outliers(tr) for i, tr in trend_by_group.items()}
    counts = [x.get("count") for x in singles] + [x.get("count") for x in pairs + trip]
    if counts and all(c is None for c in counts):
        status = "failed"
    elif any(c is None for c in counts):
        status = "partial"
    else:
        status = "ok"
    elapsed = round(time.time() - t0, 1)
    report(progress, f"缺口矩陣完成（{elapsed} 秒）" + ("；部分請求失敗" if last_errors else ""), 1.0)
    # 注意：status 在下方 return 前計算；此處訊息僅供進度列
    last_full = _dt.date.today().year - 1
    return {"status": status, "trend_status": trend_status, "trend_outliers": outliers,
            "trend_years": (last_full - 9, last_full), "partial_year_excluded": last_full + 1,
            "groups": G, "threshold": threshold, "budget_exhausted": budget_out[0], "singles": singles, "pairs": pairs, "triples": trip,
            "trend": trend, "trend_by_group": trend_by_group, "elapsed_sec": elapsed, "errors": list(last_errors),
            "source": "OpenAlex works search（標題／摘要／全文索引）"}


def matrix_table(gap: dict) -> list:
    """兩兩組合的 N×N 表（對角線＝單組命中），供 xlsx／GUI 顯示。"""
    G = gap.get("groups") or []
    n = len(G)
    M = [[None] * n for _ in range(n)]
    for s in gap.get("singles") or []:
        M[s["group"]][s["group"]] = s.get("count")
    for p in gap.get("pairs") or []:
        i, j = p["groups"]
        M[i][j] = M[j][i] = p.get("count")
    return M


def estimate_requests(n_groups: int, triples: bool = True, trend_groups: str = "first") -> int:
    """缺口矩陣需要的 OpenAlex 檢索次數（無 API key 時每日約 100 次，GUI 可先提示）。"""
    from math import comb
    return n_groups + comb(n_groups, 2) + (comb(n_groups, 3) if triples else 0) + (n_groups if trend_groups == "all" else 1)
