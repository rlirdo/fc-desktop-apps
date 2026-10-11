"""離線自測（供 app.py --selftest 呼叫）：samples.json → 指標 → APA 7 → 全格式匯出，不連網。

    from core.selftest_core import run_offline_selftest
    ok = run_offline_selftest(out_dir, log=print)   # 回傳 True/False
"""
from __future__ import annotations

import os
import traceback
from typing import Callable, Optional

from .models import load_samples


# Vatansever 2012 的真實 API 欄位（2026/10/11 取得）：Crossref 無 page，OpenAlex biblio 有 255–266，S2 journal.pages 有
_CR_VATANSEVER = {
    "DOI": "10.1515/plm-2012-0034", "type": "journal-article",
    "title": ["Far infrared radiation (FIR): Its biological effects and medical applications"],
    "author": [{"given": "Fatma", "family": "Vatansever"}, {"given": "Michael R.", "family": "Hamblin"}],
    "container-title": ["Photonics &amp; Lasers in Medicine"], "volume": "1", "issue": "4",
    "issued": {"date-parts": [[2012]]}, "ISSN": ["2193-0635", "2193-0643"], "publisher": "Walter de Gruyter GmbH",
}
_OA_VATANSEVER = {
    "id": "https://openalex.org/W0", "doi": "https://doi.org/10.1515/plm-2012-0034", "type": "article",
    "title": "Far infrared radiation (FIR): its biological effects and medical applications", "publication_year": 2012,
    "authorships": [{"raw_author_name": "Fatma Vatansever"}, {"raw_author_name": "Michael R. Hamblin"}],
    "primary_location": {"source": {"display_name": "Photonics & Lasers in Medicine", "type": "journal"}},
    "biblio": {"volume": "1", "issue": "4", "first_page": "255", "last_page": "266"},
}
_VATANSEVER_APA = ("Vatansever, F., & Hamblin, M. R. (2012). Far infrared radiation (FIR): Its biological effects and "
                   "medical applications. Photonics & Lasers in Medicine, 1(4), 255–266. https://doi.org/10.1515/plm-2012-0034")


def regression_checks() -> list:
    """2026/10/11 修正的離線回歸測試；每項回傳 "OK ..." 或 "FAIL ..."。"""
    from .apa import format_apa7
    from .enrich import apply_s2
    from .models import Paper
    from .search import parse_crossref, parse_openalex, dedupe
    from . import gap as gapmod
    from .net import ApiError
    out = []
    # 1a DOI 路徑：Crossref 無頁碼＋OpenAlex 有 → 合併後頁碼完整（兩種順序）
    for order in ("CR→OA", "OA→CR"):
        a, b = parse_crossref(dict(_CR_VATANSEVER)), parse_openalex(dict(_OA_VATANSEVER))
        merged = dedupe([a, b] if order == "CR→OA" else [b, a])
        s = format_apa7(merged[0]) if len(merged) == 1 else "（未合併）"
        out.append(("OK  " if s == _VATANSEVER_APA else "FAIL") + f" 回歸 Vatansever DOI 合併 {order}：{s}")
    # 1b OpenAlex 不可用時：Crossref＋S2 journal.pages 補頁碼
    a = parse_crossref(dict(_CR_VATANSEVER))
    apply_s2(a, {"citationCount": 340, "journal": {"name": "Photonics & lasers in medicine", "pages": "255 - 266", "volume": "4"}})
    s = format_apa7(a)
    out.append(("OK  " if s == _VATANSEVER_APA else "FAIL") + f" 回歸 Vatansever Crossref＋S2 頁碼：{s}")
    # 3 語言判定
    cases = [
        (Paper(title="蛇紋岩の研究【Powered by NICT】", language="zh"), False),
        (Paper(title="遠赤外線放射材料", language="ja"), False),
        (Paper(title="蛇紋岩のかんらん石について"), False),
        (Paper(title="深層海水之研究", authors=[{"family": "王小明", "given": ""}]), True),
        (Paper(title="Serpentine study", authors=[{"family": "Cao", "given": "C."}], language="en"), False),
    ]
    bad = [c[0].title for c in cases if c[0].is_chinese != c[1]]
    out.append(("OK  " if not bad else "FAIL") + " 回歸 語言判定（NICT／ja／假名不套中文）" + (f"：{bad}" if bad else ""))
    # 2 趨勢異常值偵測（B 回報之 far infrared 實際回應：2026＝52,796）
    tr = {2017: 3405, 2018: 3453, 2019: 3645, 2020: 3604, 2021: 4042, 2022: 4360, 2023: 4741, 2024: 4464, 2025: 3867, 2026: 52796}
    o = gapmod.find_outliers(tr)
    out.append(("OK  " if o == [2026] else "FAIL") + f" 回歸 趨勢異常值偵測：{o}")
    # 5 全部查詢失敗 → status="failed"
    orig = gapmod.cached_json

    def boom(*a, **k):
        e = ApiError("模擬：OpenAlex 額度用完", 429)
        e.budget = True
        raise e
    gapmod.cached_json = boom
    try:
        g = gapmod.gap_matrix([["a"], ["b"], ["c"]], 5, mailto="x")
    finally:
        gapmod.cached_json = orig
    good = g.get("status") == "failed" and all(r["gap"] is None for r in g["pairs"] + g["triples"])
    out.append(("OK  " if good else "FAIL") + f" 回歸 缺口矩陣全失敗 status={g.get('status')}、趨勢 {g.get('trend_status')}")
    return out


def run_offline_selftest(out_dir: str, log: Optional[Callable[[str], None]] = None, pdf_engine: str = "reportlab") -> bool:
    log = log or print
    ok = True
    try:
        from .apa import format_apa7, in_text, sort_key
        from .export import export_all, FILES
        from .indicators import evaluate
        from . import stub
        ps = load_samples()
        log(f"樣本 {len(ps)} 筆")
        if len(ps) < 15:
            log("FAIL 樣本少於 15 筆")
            ok = False
        evaluate(ps, {"keywords": ["serpentine; 蛇紋石", "surfactant; micelle; TPGS"], "year_from": 2016})
        for p in sorted(ps, key=sort_key):
            s = format_apa7(p)
            if not s or not in_text(p):
                log(f"FAIL APA 空白：{p.title[:40]}")
                ok = False
        big = [p for p in ps if len(p.authors) >= 21]
        if big and ". . . " not in format_apa7(big[0]):
            log("FAIL ≥21 作者未使用省略號")
            ok = False
        if not any(p.is_chinese for p in ps):
            log("FAIL 無中文樣本")
            ok = False
        for msg in regression_checks():
            log(msg)
            ok = ok and msg.startswith("OK")
        gap = stub.gap_matrix([["serpentine", "蛇紋石"], ["deep sea water"], ["TPGS"]])
        res = export_all(ps, gap, out_dir, "selftest", {"pdf_engine": pdf_engine, "keywords": ["serpentine; 蛇紋石"],
                                                         "settings": {"year_from": 2016}})
        for path, good in res.items():
            size = os.path.getsize(path) if os.path.isfile(path) else 0
            log(f"{'OK  ' if good and size > 0 else 'FAIL'} {os.path.basename(path)} {size} bytes")
            ok = ok and good and size > 0
        if len(res) < len(FILES):
            log(f"FAIL 匯出檔數 {len(res)} < {len(FILES)}")
            ok = False
    except Exception:
        log("FAIL 例外：\n" + traceback.format_exc())
        ok = False
    log("SELFTEST " + ("PASS" if ok else "FAIL"))
    return ok
