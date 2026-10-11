"""共用 HTTP：requests＋User-Agent/mailto、timeout 20 s、重試 3 次、依主機節流。

一般 API ≤ 5 req/s；Semantic Scholar ≤ 1 req/s。
"""
from __future__ import annotations

import threading
import time
from typing import Optional
from urllib.parse import urlparse

import requests

from .models import VERSION, check_cancel

TIMEOUT = 20
RETRIES = 3
_RATE = {"api.semanticscholar.org": 1.05}   # 秒／次
_DEFAULT_INTERVAL = 0.2                      # 5 req/s

_local = threading.local()
_rate_lock = threading.Lock()
_next_slot: dict = {}


class ApiError(Exception):
    def __init__(self, msg: str, status: Optional[int] = None):
        super().__init__(msg)
        self.status = status
        self.budget = False      # True＝每日額度用完（重試無用）


def budget_message(host: str, retry_after: float = 0.0) -> str:
    hrs = f"約 {retry_after / 3600:.1f} 小時後重置" if retry_after else "每日 UTC 0 時重置"
    if "openalex" in host:
        return (f"OpenAlex 免費額度已用完（無 API key 時同一網路每日約 100 次檢索，{hrs}）。"
                "已取得的結果仍可匯出；可於設定填入免費 OpenAlex API key 或明日再試。")
    return f"{host} 額度已用完（{hrs}）"


def openalex_params(params: Optional[dict]) -> dict:
    """OpenAlex：若使用者有設定 API key（環境變數 FCLITGAP_OPENALEX_KEY 或設定 openalex_api_key）則附上。"""
    import os
    p = dict(params or {})
    key = os.environ.get("FCLITGAP_OPENALEX_KEY", "")
    if not key:
        try:
            from .config import load
            key = load().get("openalex_api_key", "") or ""
        except Exception:
            key = ""
    if key:
        p["api_key"] = key
    return p


def user_agent(mailto: str = "") -> str:
    ua = f"FCLitGapFinder/{VERSION} (NDHU academic literature tool"
    if mailto:
        ua += f"; mailto:{mailto}"
    return ua + ")"


def session() -> requests.Session:
    s = getattr(_local, "s", None)
    if s is None:
        s = requests.Session()
        _local.s = s
    return s


def _throttle(host: str) -> None:
    interval = _RATE.get(host, _DEFAULT_INTERVAL)
    with _rate_lock:
        now = time.monotonic()
        slot = max(now, _next_slot.get(host, 0.0))
        _next_slot[host] = slot + interval
    wait = slot - now
    if wait > 0:
        time.sleep(wait)


def request(method: str, url: str, params: Optional[dict] = None, mailto: str = "", headers: Optional[dict] = None,
            retries: int = RETRIES, cancel_event=None, allow_redirects: bool = True, stream: bool = False,
            ok_status=(200,)) -> requests.Response:
    host = urlparse(url).netloc
    if host == "api.openalex.org":
        params = openalex_params(params)
    h = {"User-Agent": user_agent(mailto)}
    if headers:
        h.update(headers)
    last = None
    for attempt in range(retries):
        check_cancel(cancel_event)
        _throttle(host)
        try:
            r = session().request(method, url, params=params, headers=h, timeout=TIMEOUT,
                                  allow_redirects=allow_redirects, stream=stream)
        except requests.RequestException as e:
            last = ApiError(f"{host} 連線失敗：{type(e).__name__}")
            time.sleep(1.0 * (attempt + 1))
            continue
        if r.status_code in ok_status:
            return r
        if r.status_code == 429:
            ra = r.headers.get("Retry-After")
            try:
                ra_s = float(ra) if ra else 0.0
            except ValueError:
                ra_s = 0.0
            if ra_s > 60 or "budget" in (r.text or "").lower():
                e = ApiError(budget_message(host, ra_s), 429)
                e.budget = True
                raise e
        if r.status_code in (429, 500, 502, 503, 504):
            last = ApiError(f"{host} 回應 {r.status_code}", r.status_code)
            ra = r.headers.get("Retry-After")
            try:
                delay = min(float(ra), 10.0) if ra else 1.5 * (attempt + 1)
            except ValueError:
                delay = 1.5 * (attempt + 1)
            time.sleep(delay)
            continue
        raise ApiError(f"{host} 回應 {r.status_code}", r.status_code)
    raise last or ApiError(f"{host} 失敗")


def get_json(url: str, params: Optional[dict] = None, mailto: str = "", cancel_event=None, retries: int = RETRIES,
             headers: Optional[dict] = None) -> dict:
    r = request("GET", url, params=params, mailto=mailto, cancel_event=cancel_event, retries=retries, headers=headers)
    try:
        return r.json()
    except ValueError:
        raise ApiError(f"{urlparse(url).netloc} 回傳非 JSON")
