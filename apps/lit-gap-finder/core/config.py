"""設定檔 %APPDATA%\\FCLitGap\\config.json（不存任何密碼）。

環境變數 FCLITGAP_HOME 可改設定／快取資料夾（自測與 CI 用）。
"""
from __future__ import annotations

import copy
import json
import os

from .models import DEFAULT_CHAPTERS

DEFAULTS = {
    "email": "811554006@gms.ndhu.edu.tw",      # Unpaywall email（可改）
    "mailto": "811554006@gms.ndhu.edu.tw",     # Crossref／OpenAlex polite pool
    "last_dir": "",
    "chapters": list(DEFAULT_CHAPTERS),
    "weights": {"title": 3.0, "abstract": 1.0, "cited": 1.0, "recency": 1.0},
    "year_from": 2016,
    "year_to": None,
    "classic": False,
    "lang": "any",                              # en | zh | any
    "types": ["journal-article"],
    "checks": {"doi": True, "oa": True, "type": True, "year": True, "issn": True, "lang": True},
    "keep_mode": "mark",                        # mark＝全部保留但標示｜strict＝只保留全部通過
    "limit": 50,
    "gap_threshold": 5,
    "s2_enabled": True,
    "s2_max": 60,                               # Semantic Scholar 每次最多查幾筆（≤1 req/s）
    "cache_days": 7,
    # OpenAlex 自 2026 起：無 key 時同一 IP 每日免費額度 1000 credits，每次 search 耗 10 credits（約 100 次）。
    # 使用者可自行至 openalex.org 申請免費 key 填入（選填；不寫入查詢紀錄與匯出檔）。
    "openalex_api_key": "",
    "gap_trend_groups": "first",                # first＝只畫第 1 組趨勢｜all＝每組一條（每組多 1 次檢索）
}


def home_dir() -> str:
    base = os.environ.get("FCLITGAP_HOME")
    if not base:
        base = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "FCLitGap")
    os.makedirs(base, exist_ok=True)
    return base


def config_path() -> str:
    return os.path.join(home_dir(), "config.json")


def _merge(dst: dict, src: dict) -> dict:
    for k, v in src.items():
        if isinstance(v, dict) and isinstance(dst.get(k), dict):
            _merge(dst[k], v)
        else:
            dst[k] = v
    return dst


def load() -> dict:
    cfg = copy.deepcopy(DEFAULTS)
    try:
        with open(config_path(), encoding="utf-8") as f:
            _merge(cfg, json.load(f))
    except (OSError, ValueError):
        pass
    if not cfg.get("mailto"):
        cfg["mailto"] = cfg.get("email") or DEFAULTS["email"]
    return cfg


def save(cfg: dict) -> None:
    data = {k: v for k, v in cfg.items() if "password" not in k.lower() and "token" not in k.lower()}
    tmp = config_path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, config_path())
