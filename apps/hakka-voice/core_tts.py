# -*- coding: utf-8 -*-
"""客語文字 → 語音：呼叫 Hugging Face Space「臺灣客語語音合成系統」的公開 Gradio API。

服務：https://huggingface.co/spaces/united-link/taiwanese-hakka-tts
模型：yourtts-htia-240704（CC BY-NC 4.0）；合作：國立聯合大學智慧客家實驗室。
這是共用的公開示範服務：一次一句依序送、句與句之間至少間隔 1 秒、單次工作上限 1,200 字。

語者清單一律在執行時從 /gradio_api/info 與 /config 取得，原始碼不寫死任何語者的值或姓名；
介面只顯示「語者 N（性別／地區）」這種友善標籤。

注意：本模組只連 united-link-taiwanese-hakka-tts.hf.space。
客委會、長問翻譯器、Gohakka 三個網站一律不以程式連線（見 SPEC 第 0 節）。
"""
import datetime
import json
import re
import time

SPACE = "https://united-link-taiwanese-hakka-tts.hf.space"
API = SPACE + "/gradio_api"
MODEL_ID = "yourtts-htia-240704"
SPEAKER_TYPE = "預設語者"
TIMEOUT = 120
MIN_INTERVAL = 1.0          # 句與句之間至少間隔（秒）
MAX_CHARS_PER_JOB = 1200
MAX_SENT = 60

ACCENTS = [("sixian", "四縣"), ("hailu", "海陸"), ("dapu", "大埔"),
           ("raoping", "饒平"), ("zhaoan", "詔安"), ("nansixian", "南四縣")]
ACCENT_ZH = dict(ACCENTS)

SOURCE_LICENSE = ("臺灣客語語音合成系統（Hugging Face Space united-link/taiwanese-hakka-tts；"
                  "合作：國立聯合大學智慧客家實驗室）；模型 yourtts-htia-240704；"
                  "授權 CC BY-NC 4.0；非商業教學使用")
REVIEW_NOTE = ("AI 生成，教師審核；給孩子聽的請由老師親口唸，"
               "本音檔供老師備課練習與家長參考")
UA = "HakkaVoiceHelper/1.0 (non-commercial teaching tool)"


class TTSError(Exception):
    pass


class Cancelled(Exception):
    pass


# ---------------------------------------------------------------- 文字處理
_END = "。！？!?\n"


def split_sentences(text, max_len=MAX_SENT):
    """依句號、問號、驚嘆號、換行切句；每句 ≤ max_len 字，過長再依逗號（，、；,;）切，仍過長就硬切。"""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    raw = []
    buf = ""
    for ch in text:
        if ch == "\n":
            if buf.strip():
                raw.append(buf.strip())
            buf = ""
            continue
        buf += ch
        if ch in _END:
            if buf.strip():
                raw.append(buf.strip())
            buf = ""
    if buf.strip():
        raw.append(buf.strip())

    out = []
    for s in raw:
        if not re.sub(r"[\s。！？!?，、；,;：:]", "", s):
            continue                       # 只有標點的句子略過
        if len(s) <= max_len:
            out.append(s)
            continue
        pieces = re.split(r"(?<=[，、；,;])", s)
        cur = ""
        for p in pieces:
            if len(cur) + len(p) <= max_len:
                cur += p
            else:
                if cur:
                    out.append(cur)
                while len(p) > max_len:
                    out.append(p[:max_len])
                    p = p[max_len:]
                cur = p
        if cur:
            out.append(cur)
    return out


_BAD = re.compile(r'[\\/:*?"<>|\r\n\t]')


def safe_filename(text, n=12):
    """取前 n 字當檔名，去掉 Windows 不合法字元與標點；空白時回傳「客語語音」。"""
    s = _BAD.sub("", text or "")
    s = re.sub(r"[\s。！？!?，、；,;：:「」『』（）()…\.]", "", s)
    s = s[:n].strip(" .")
    return s or "客語語音"


def unique_path(folder, stem, ext=".mp3"):
    import os
    """自動加序號避免覆蓋（mp3 與來源說明檔都不覆蓋）。"""
    name, i = stem, 2
    while (os.path.exists(os.path.join(folder, name + ext))
           or os.path.exists(os.path.join(folder, name + "_來源說明.txt"))):
        name = "%s_%d" % (stem, i)
        i += 1
    return os.path.join(folder, name + ext)


VERSION = "1.0.1"


def source_note(*, mandarin, hakka, segs, pinyins, accent, speaker_label, speed,
                translate_source, when=None, official_pinyin="", compare_lines=None):
    """來源說明檔內容（UTF-8 文字）。"""
    when = when or datetime.datetime.now()
    lines = ["客語語音小幫手 HakkaVoiceHelper %s — 來源說明" % VERSION, ""]
    if (mandarin or "").strip():
        lines += ["【華語原句】", mandarin.strip(), ""]
    lines += ["【客語漢字】", (hakka or "").strip(), ""]
    lines += ["【斷詞】"] + list(segs or []) + [""]
    lines += ["【客語拼音】"] + list(pinyins or []) + [""]
    lines += ["【官方拼音】", (official_pinyin or "").strip() or "未提供", ""]
    lines += ["【與官方拼音比對】"] + (list(compare_lines) if compare_lines else
                                     ["未提供官方拼音，未比對；請老師聽一遍確認"]) + [""]
    lines += ["腔調：%s" % ACCENT_ZH.get(accent, accent),
              "語者：%s" % speaker_label,
              "語速：%s（越大越慢）" % speed,
              "翻譯來源：%s" % (translate_source or "自行輸入")]
    if translate_source and "gohakka" in translate_source.lower():
        lines.append("翻譯來源：GoHakka.org（CC BY-SA 4.0）")
    lines += ["語音來源與授權：" + SOURCE_LICENSE,
              "產生時間：%s" % when.strftime("%Y-%m-%d %H:%M:%S"),
              "",
              REVIEW_NOTE, ""]
    return "\n".join(lines)


# ---------------------------------------------------------------- 網路
def _session():
    import requests
    s = requests.Session()
    s.headers["User-Agent"] = UA
    return s


def friendly_label(i, label):
    """「X / 男 / 苗栗」→「語者 3（男／苗栗）」：不顯示姓名。"""
    parts = [p.strip() for p in str(label).split("/")]
    extra = "／".join(p for p in parts[1:] if p)
    return "語者 %d（%s）" % (i, extra) if extra else "語者 %d" % i


def fetch_speakers(session=None, timeout=30):
    """回傳 (speakers, default_value)。speakers：[{"value", "label", "raw_label"}]；取不到就丟 TTSError。"""
    s = session or _session()
    try:
        info = s.get(API + "/info", timeout=timeout).json()
        ps = info["named_endpoints"]["/text_to_speech"]["parameters"]
        values = list(ps[3]["type"].get("enum") or [])
        default = ps[3].get("parameter_default")
    except Exception as e:  # noqa: BLE001
        raise TTSError("取不到語者清單（語音合成服務可能在休眠或連不上）：%s" % e)
    raw = {}
    try:
        cfg = s.get(SPACE + "/config", timeout=timeout).json()
        for c in cfg.get("components", []):
            pr = c.get("props", {})
            if pr.get("label") == "語者" and pr.get("choices"):
                for ch in pr["choices"]:
                    if isinstance(ch, (list, tuple)) and len(ch) == 2:
                        raw[ch[1]] = ch[0]
    except Exception:  # noqa: BLE001
        pass
    if not values:
        values = list(raw.keys())
    if not values:
        raise TTSError("語音合成服務沒有回傳語者清單")
    if default not in values:
        default = values[0]
    sp = [{"value": v, "raw_label": raw.get(v, ""), "label": friendly_label(i + 1, raw.get(v, ""))}
          for i, v in enumerate(values)]
    return sp, default


def _check(cancel):
    if cancel is not None and cancel.is_set():
        raise Cancelled()


def synth_once(text, speaker, accent="sixian", speed=1.0, session=None, cancel=None,
               timeout=TIMEOUT):
    """合成一句 → (斷詞, 拼音, WAV 位元組)。"""
    s = session or _session()
    deadline = time.time() + timeout
    payload = {"data": [MODEL_ID, SPEAKER_TYPE, None, speaker, accent, float(speed), text]}
    r = s.post(API + "/call/text_to_speech", json=payload, timeout=min(60, timeout))
    if r.status_code != 200:
        raise TTSError("語音合成服務回應 HTTP %d" % r.status_code)
    eid = r.json().get("event_id")
    if not eid:
        raise TTSError("語音合成服務沒有回傳工作編號")
    _check(cancel)
    result = None
    with s.get(API + "/call/text_to_speech/" + eid, stream=True, timeout=timeout) as st:
        ev = None
        for line in st.iter_lines(decode_unicode=True):
            _check(cancel)
            if time.time() > deadline:
                raise TTSError("語音合成逾時（%d 秒）" % timeout)
            if not line:
                continue
            if line.startswith("event:"):
                ev = line[6:].strip()
            elif line.startswith("data:"):
                d = line[5:].strip()
                if ev == "complete":
                    result = json.loads(d)
                    break
                if ev == "error":
                    raise TTSError("語音合成服務回報錯誤：%s" % (d if d and d != "null" else "（未說明，可能在休眠或忙碌）"))
    if not result or len(result) < 3:
        raise TTSError("語音合成服務沒有回傳結果")
    seg, pin, audio = result[0], result[1], result[2]
    url = audio.get("url") if isinstance(audio, dict) else None
    if not url:
        raise TTSError("語音合成結果沒有音檔網址")
    _check(cancel)
    wav = s.get(url, timeout=max(10, deadline - time.time())).content
    from core_audio import is_wav
    if not is_wav(wav):
        raise TTSError("下載到的音檔不是 WAV")
    return str(seg or ""), str(pin or ""), wav


def synth_text(text, speaker, accent="sixian", speed=1.0, progress=None, cancel=None,
               session=None):
    """整段文字：切句 → 逐句（間隔 ≥1 秒、失敗重試 1 次）→ [(句, 斷詞, 拼音, wav)]。"""
    sents = split_sentences(text)
    if not sents:
        raise TTSError("沒有可以合成的文字")
    total = sum(len(x) for x in sents)
    if total > MAX_CHARS_PER_JOB:
        raise TTSError("一次最多 %d 字（目前 %d 字），請分批產生" % (MAX_CHARS_PER_JOB, total))
    s = session or _session()
    out = []
    last = 0.0
    for i, sent in enumerate(sents):
        for attempt in (1, 2):
            _check(cancel)
            wait = MIN_INTERVAL - (time.time() - last)
            if wait > 0:
                time.sleep(wait)
            if progress:
                progress("合成第 %d／%d 句%s：%s" % (i + 1, len(sents),
                                                 "（重試）" if attempt == 2 else "", sent[:20]))
            last = time.time()
            try:
                seg, pin, wav = synth_once(sent, speaker, accent, speed, session=s, cancel=cancel)
                out.append((sent, seg, pin, wav))
                break
            except Cancelled:
                raise
            except Exception as e:  # noqa: BLE001
                last = time.time()
                if attempt == 2:
                    raise TTSError("第 %d 句合成失敗：%s" % (i + 1, e))
                time.sleep(2.0)
    return out
