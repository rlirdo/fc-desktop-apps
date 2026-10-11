"""離線自測（供 app.py --selftest 呼叫）：samples.json → 指標 → APA 7 → 全格式匯出，不連網。

    from core.selftest_core import run_offline_selftest
    ok = run_offline_selftest(out_dir, log=print)   # 回傳 True/False
"""
from __future__ import annotations

import os
import traceback
from typing import Callable, Optional

from .models import load_samples


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
