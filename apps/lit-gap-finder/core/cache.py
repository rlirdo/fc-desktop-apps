"""SQLite 快取 %APPDATA%\\FCLitGap\\cache.sqlite；同一查詢 7 天內不重查。

鍵慣例：
  q:<api>:<url+params>      查詢回應（整包 JSON）
  raw:crossref:<doi小寫>    單筆原始 JSON（不杜撰書目的證據）
  raw:openalex:<id>         同上
"""
from __future__ import annotations

import json
import os
import sqlite3
import threading
import time
from typing import Any, Optional

from .config import home_dir

TTL_DAYS = 7
_lock = threading.Lock()


def db_path() -> str:
    return os.path.join(home_dir(), "cache.sqlite")


def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(db_path(), timeout=30)
    c.execute("CREATE TABLE IF NOT EXISTS kv (k TEXT PRIMARY KEY, v TEXT NOT NULL, t REAL NOT NULL)")
    return c


def get(key: str, max_age_days: Optional[float] = TTL_DAYS) -> Any:
    with _lock:
        try:
            c = _conn()
            try:
                row = c.execute("SELECT v, t FROM kv WHERE k=?", (key,)).fetchone()
            finally:
                c.close()
        except sqlite3.Error:
            return None
    if not row:
        return None
    if max_age_days is not None and time.time() - row[1] > max_age_days * 86400:
        return None
    try:
        return json.loads(row[0])
    except ValueError:
        return None


def set(key: str, value: Any) -> None:  # noqa: A001 — 契約名稱
    s = json.dumps(value, ensure_ascii=False)
    with _lock:
        try:
            c = _conn()
            try:
                c.execute("INSERT OR REPLACE INTO kv (k, v, t) VALUES (?, ?, ?)", (key, s, time.time()))
                c.commit()
            finally:
                c.close()
        except sqlite3.Error:
            pass


def set_many(items: list) -> None:
    now = time.time()
    rows = [(k, json.dumps(v, ensure_ascii=False), now) for k, v in items]
    if not rows:
        return
    with _lock:
        try:
            c = _conn()
            try:
                c.executemany("INSERT OR REPLACE INTO kv (k, v, t) VALUES (?, ?, ?)", rows)
                c.commit()
            finally:
                c.close()
        except sqlite3.Error:
            pass


def clear() -> int:
    with _lock:
        try:
            c = _conn()
            try:
                n = c.execute("SELECT COUNT(*) FROM kv").fetchone()[0]
                c.execute("DELETE FROM kv")
                c.commit()
                c.execute("VACUUM")
            finally:
                c.close()
            return n
        except sqlite3.Error:
            return 0


def stats() -> dict:
    try:
        c = _conn()
        try:
            n = c.execute("SELECT COUNT(*) FROM kv").fetchone()[0]
        finally:
            c.close()
        return {"entries": n, "path": db_path(), "bytes": os.path.getsize(db_path())}
    except (sqlite3.Error, OSError):
        return {"entries": 0, "path": db_path(), "bytes": 0}
