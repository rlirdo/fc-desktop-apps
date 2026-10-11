"""合法 OA 全文下載：只下載 Unpaywall best_oa_location.url_for_pdf（或 OpenAlex pdf_url）。

存 out_dir\\OA全文\\<年>_<第一作者>_<短題>.pdf，並寫 out_dir\\OA全文\\OA下載紀錄.csv
（DOI、來源網址、授權 license、下載時間、狀態、檔名）。非 OA 不嘗試任何其他來源。
"""
from __future__ import annotations

import csv
import datetime as _dt
import os
import re
from typing import Optional

from .apa import clean_text
from .models import Paper, Progress, check_cancel, report
from .net import request, ApiError

MAX_BYTES = 80 * 1024 * 1024
LOG_NAME = "OA下載紀錄.csv"


def pdf_filename(p: Paper) -> str:
    fam = re.sub(r'[\\/:*?"<>|\s]+', "", p.first_author_family) or "佚名"
    words = re.findall(r"[\w一-鿿]+", clean_text(p.title))
    short = "_".join(words[:6])[:50] or "untitled"
    return f"{p.year or 'nd'}_{fam}_{short}.pdf"


def _fetch_pdf(url: str, path: str, mailto: str, cancel_event) -> str:
    """下載單一網址；成功回 "ok"，否則回失敗原因。"""
    tmp = path + ".part"
    try:
        r = request("GET", url, mailto=mailto, stream=True, cancel_event=cancel_event, retries=2,
                    headers={"Accept": "application/pdf,*/*;q=0.8"})
        size = 0
        first = b""
        with open(tmp, "wb") as f:
            for chunk in r.iter_content(65536):
                check_cancel(cancel_event)
                if not first:
                    first = chunk[:8]
                size += len(chunk)
                if size > MAX_BYTES:
                    raise ApiError("檔案超過 80 MB")
                f.write(chunk)
        if not first.startswith(b"%PDF"):
            ctype = r.headers.get("Content-Type", "")
            return f"非 PDF（{ctype.split(';')[0] or '未知'}）"
        os.replace(tmp, path)
        return "ok"
    except ApiError as e:
        return str(e)
    except OSError as e:
        return f"存檔失敗 {e}"
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def download_oa(papers: list, out_dir: str, progress: Progress = None, cancel_event=None,
                mailto: str = "") -> list:
    folder = os.path.join(out_dir, "OA全文")
    os.makedirs(folder, exist_ok=True)
    rows = []
    cands = [p for p in papers if p.include]
    for i, p in enumerate(cands):
        check_cancel(cancel_event)
        report(progress, f"OA 全文 {i + 1}/{len(cands)}：{p.first_author_family} {p.year or ''}", i / max(1, len(cands)))
        row = {"DOI": p.doi, "來源網址": p.oa_pdf_url, "授權 license": p.oa_license or "",
               "下載時間": "", "狀態": "", "檔名": ""}
        if not p.oa_pdf_url:
            row["狀態"] = "非 OA 或無 PDF 連結（僅提供 DOI）"
            rows.append(row)
            continue
        fn = pdf_filename(p)
        path = os.path.join(folder, fn)
        if os.path.isfile(path) and os.path.getsize(path) > 1000:
            row.update({"狀態": "已存在", "檔名": fn, "下載時間": _dt.datetime.now().isoformat(timespec="seconds")})
            rows.append(row)
            continue
        msgs = []
        for url in [p.oa_pdf_url] + list(getattr(p, "oa_pdf_alts", []) or [])[:3]:
            check_cancel(cancel_event)
            st = _fetch_pdf(url, path, mailto, cancel_event)
            if st == "ok":
                row.update({"狀態": "已下載" + ("（備用 OA 位置）" if url != p.oa_pdf_url else ""), "檔名": fn,
                            "來源網址": url})
                break
            msgs.append(st)
        else:
            row["狀態"] = "下載失敗：" + "；".join(msgs) + "（請以 DOI 連結手動開啟）"
        row["下載時間"] = _dt.datetime.now().isoformat(timespec="seconds")
        rows.append(row)
    log = os.path.join(folder, LOG_NAME)
    new = not os.path.isfile(log)
    with open(log, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else
                           ["DOI", "來源網址", "授權 license", "下載時間", "狀態", "檔名"])
        if new:
            w.writeheader()
        w.writerows(rows)
    ok = sum(1 for r in rows if r["狀態"] in ("已下載", "已存在"))
    report(progress, f"OA 全文完成：{ok}/{len(rows)} 筆；紀錄 {LOG_NAME}", 1.0)
    return rows
