# -*- coding: utf-8 -*-
"""GUI 與 core 之間的轉接層。

模式：
  real  → core.search / enrich / indicators / gap / apa / export / oa_download / cache / config
  stub  → core.stub（離線假資料，同名同簽章）
  auto  → 真模組可匯入就用真模組，否則退回 stub（並在 status 記錄原因）
環境變數 LITGAP_BACKEND=stub|real|auto（預設 auto）。
"""
import inspect
import os
import sys

_HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from core import models  # noqa: E402
from core.models import Paper, Cancelled, INDICATOR_NAMES, TYPE_LABELS, DEFAULT_CHAPTERS  # noqa: E402,F401

MODE = "stub"
STATUS = ""


class _Backend:
    pass


B = _Backend()


def _load_stub():
    from core import stub
    B.search = stub.search
    B.enrich = stub.enrich
    B.evaluate = stub.evaluate
    B.gap_matrix = stub.gap_matrix
    B.format_apa7 = stub.format_apa7
    B.in_text = stub.in_text
    B.sort_key = stub.sort_key
    B.export_all = stub.export_all
    B.download_oa = stub.download_oa
    B.cache_clear = stub.cache_clear
    B.config_load = stub.load
    B.config_save = stub.save
    B.DEFAULTS = stub.DEFAULTS


def _load_real():
    from core.search import search
    from core.enrich import enrich
    from core.indicators import evaluate
    from core.gap import gap_matrix
    from core.apa import format_apa7, in_text, sort_key
    from core.export import export_all
    from core.oa_download import download_oa
    from core import cache, config
    B.search = search
    B.enrich = enrich
    B.evaluate = evaluate
    B.gap_matrix = gap_matrix
    B.format_apa7 = format_apa7
    B.in_text = in_text
    B.sort_key = sort_key
    B.export_all = export_all
    B.download_oa = download_oa
    B.cache_clear = getattr(cache, "clear", None) or getattr(cache, "cache_clear")
    B.config_load = config.load
    B.config_save = config.save
    B.DEFAULTS = getattr(config, "DEFAULTS", {})


def init(mode=None):
    """選擇後端；回傳實際模式字串。"""
    global MODE, STATUS
    mode = (mode or os.environ.get("LITGAP_BACKEND") or "auto").lower()
    if mode in ("real", "auto"):
        try:
            _load_real()
            MODE, STATUS = "real", "核心模組：正式版"
            return MODE
        except Exception as e:  # noqa: BLE001
            if mode == "real":
                raise
            STATUS = f"核心模組未就緒，改用離線假資料（{type(e).__name__}: {e}）"
    _load_stub()
    MODE = "stub"
    STATUS = STATUS or "核心模組：離線假資料（stub）"
    return MODE


def call(fn, *args, **kw):
    """只把函式簽章接受的關鍵字參數傳進去（例如 cancel_event 選用）。"""
    try:
        sig = inspect.signature(fn)
        params = sig.parameters
        if not any(p.kind == p.VAR_KEYWORD for p in params.values()):
            kw = {k: v for k, v in kw.items() if k in params}
    except (TypeError, ValueError):
        pass
    return fn(*args, **kw)


def load_samples():
    return models.load_samples()


def __getattr__(name):  # 讓 backend.search(...) 這種寫法可用
    if hasattr(B, name):
        return getattr(B, name)
    raise AttributeError(name)
