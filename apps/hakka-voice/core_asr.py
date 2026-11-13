# -*- coding: utf-8 -*-
"""客語語音 → 文字（試用版）：本機離線 faster-whisper（CTranslate2）。

模型：國立臺南大學 KWS 中心 NUTN-KWS/Whisper-Taiwanese-Hakka-model-v0.2.6
（whisper-large-v3-turbo 微調，CC BY-NC 4.0）轉成的 CTranslate2 int8 版。
音檔不離開電腦；不使用 VAD（避免帶入 onnxruntime）；不做自動下載。
實測輸出是偏華語的繁體中文、準確度有限，只能抓大意，一定要校對。
"""
import json
import os
import sys

MODEL_DIRNAME = "hakka-whisper-ct2-int8"
APP_DIRNAME = "HakkaVoiceHelper"
REQUIRED = ("model.bin", "config.json")

# 2026/11/13 實測後的辨識參數（SPEC 1-2b）。
# 試驗：repetition_penalty=1.1 讓 probe.wav 變差（「早安早安細兔會去學校上課」→「早餐 細兔仔細小」），
# no_repeat_ngram_size=3 在五個測試檔結果完全不變，兩者皆不採用。
TRANSCRIBE_KW = dict(language="zh", task="transcribe", beam_size=5, vad_filter=False,
                     condition_on_previous_text=False)

MISSING_MSG = ("語音辨識需要約 0.8 GB 的模型資料夾，請向講師索取 models 資料夾，放在本程式旁邊"
               "（models\\%s\\）。也可以按「選擇模型資料夾…」指定位置。" % MODEL_DIRNAME)


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def settings_path():
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, APP_DIRNAME, "settings.json")


def load_settings():
    try:
        with open(settings_path(), encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:  # noqa: BLE001
        return {}


def save_settings(d):
    p = settings_path()
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        return True
    except Exception:  # noqa: BLE001
        return False


def is_model_dir(path):
    return bool(path) and os.path.isdir(path) and all(
        os.path.exists(os.path.join(path, f)) for f in REQUIRED)


def candidate_dirs():
    """搜尋順序：① exe 同層 models\\… ② %LOCALAPPDATA%\\HakkaVoiceHelper\\models\\… ③ 設定檔指定路徑。"""
    c = [os.path.join(app_dir(), "models", MODEL_DIRNAME)]
    la = os.environ.get("LOCALAPPDATA")
    if la:
        c.append(os.path.join(la, APP_DIRNAME, "models", MODEL_DIRNAME))
    s = load_settings().get("model_dir")
    if s:
        c.append(s)
    return c


def find_model(explicit=None):
    if explicit:
        return explicit if is_model_dir(explicit) else None
    for p in candidate_dirs():
        if is_model_dir(p):
            return p
    return None


_MODEL_CACHE = {}


def load_model(path):
    if path in _MODEL_CACHE:
        return _MODEL_CACHE[path]
    from faster_whisper import WhisperModel
    m = WhisperModel(path, device="cpu", compute_type="int8")
    _MODEL_CACHE[path] = m
    return m


def transcribe_iter(model, audio_path, **extra):
    """產生器：逐段回傳 (start, end, text)，第一個回傳值是 ('info', 總長秒數)。"""
    kw = dict(TRANSCRIBE_KW)
    kw.update(extra)
    segs, info = model.transcribe(audio_path, **kw)
    yield ("info", float(getattr(info, "duration", 0.0) or 0.0))
    for s in segs:
        yield (float(s.start), float(s.end), s.text.strip())


def transcribe_text(model, audio_path, **extra):
    parts = []
    for item in transcribe_iter(model, audio_path, **extra):
        if item[0] == "info":
            continue
        parts.append(item[2])
    return join_segments(parts)


def join_segments(parts):
    return "\n".join(p for p in parts if p)
