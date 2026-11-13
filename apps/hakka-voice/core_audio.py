# -*- coding: utf-8 -*-
"""音訊工具：WAV 讀取、PCM 串接、用 PyAV（libmp3lame）編碼成 mp3、讀取音檔長度。

不使用 ffmpeg 執行檔、不使用 pydub；PyAV 的 wheel 內含 libav 與 libmp3lame。
"""
import io
import math
import wave

import numpy as np

GAP_SECONDS = 0.35          # 句與句之間插入的靜音
MP3_BITRATE = 128000
FALLBACK_RATES = (24000, 44100)


def read_wav_bytes(data):
    """WAV 位元組 → (float32 單聲道 numpy 陣列 -1..1, 取樣率)。"""
    with wave.open(io.BytesIO(data), "rb") as w:
        ch = w.getnchannels()
        sw = w.getsampwidth()
        sr = w.getframerate()
        raw = w.readframes(w.getnframes())
    if sw == 2:
        a = np.frombuffer(raw, dtype="<i2").astype(np.float32) / 32768.0
    elif sw == 4:
        a = np.frombuffer(raw, dtype="<i4").astype(np.float32) / 2147483648.0
    elif sw == 1:
        a = (np.frombuffer(raw, dtype=np.uint8).astype(np.float32) - 128.0) / 128.0
    else:
        raise ValueError("不支援的 WAV 取樣寬度：%d bytes" % sw)
    if ch > 1:
        a = a.reshape(-1, ch).mean(axis=1)
    return a, sr


def make_wav_bytes(samples, sr):
    """float32 陣列 → 16-bit 單聲道 WAV 位元組（測試用）。"""
    pcm = np.clip(samples, -1.0, 1.0)
    pcm = (pcm * 32767.0).astype("<i2").tobytes()
    bio = io.BytesIO()
    with wave.open(bio, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(sr)
        w.writeframes(pcm)
    return bio.getvalue()


def sine_wave(seconds, sr=16000, freq=440.0):
    t = np.arange(int(seconds * sr), dtype=np.float32) / sr
    return (0.3 * np.sin(2 * math.pi * freq * t)).astype(np.float32)


def resample(a, src, dst):
    if src == dst or len(a) == 0:
        return a
    n = int(round(len(a) * dst / float(src)))
    x_old = np.linspace(0.0, 1.0, num=len(a), endpoint=False)
    x_new = np.linspace(0.0, 1.0, num=n, endpoint=False)
    return np.interp(x_new, x_old, a).astype(np.float32)


def concat_wavs(wav_list, gap=GAP_SECONDS):
    """多段 WAV 位元組 → (float32 陣列, 取樣率)；取樣率以第一段為準，其餘重取樣。"""
    parts = []
    sr0 = None
    for data in wav_list:
        a, sr = read_wav_bytes(data)
        if sr0 is None:
            sr0 = sr
        elif sr != sr0:
            a = resample(a, sr, sr0)
        if parts:
            parts.append(np.zeros(int(gap * sr0), dtype=np.float32))
        parts.append(a)
    if not parts:
        raise ValueError("沒有任何音訊可以串接")
    return np.concatenate(parts), sr0


def _encode_mp3(samples, sr, path):
    import av
    pcm = (np.clip(samples, -1.0, 1.0) * 32767.0).astype(np.int16)
    with av.open(path, mode="w", format="mp3") as cont:
        st = cont.add_stream("libmp3lame", rate=sr)
        st.bit_rate = MP3_BITRATE
        try:
            st.layout = "mono"
        except Exception:
            pass
        cc = st.codec_context
        frame_size = cc.frame_size or 1152
        for i in range(0, len(pcm), frame_size):
            chunk = pcm[i:i + frame_size]
            fr = av.AudioFrame.from_ndarray(chunk.reshape(1, -1), format="s16", layout="mono")
            fr.sample_rate = sr
            for pkt in st.encode(fr):
                cont.mux(pkt)
        for pkt in st.encode(None):
            cont.mux(pkt)


def write_mp3(samples, sr, path):
    """存成單聲道 128 kbps mp3；編碼器不接受原取樣率時改用 24000 / 44100。回傳實際取樣率。"""
    tried = []
    for rate in (sr,) + tuple(r for r in FALLBACK_RATES if r != sr):
        try:
            _encode_mp3(resample(samples, sr, rate), rate, path)
            return rate
        except Exception as e:  # noqa: BLE001
            tried.append("%d Hz：%s" % (rate, e))
    raise RuntimeError("mp3 編碼失敗：" + "；".join(tried))


def audio_duration(path):
    """用 PyAV 取音檔長度（秒）；容器沒記錄時改用解碼計算。"""
    import av
    with av.open(path) as c:
        if c.duration:
            return c.duration / 1_000_000.0
        st = c.streams.audio[0]
        if st.duration and st.time_base:
            return float(st.duration * st.time_base)
        total = 0
        rate = st.rate or 1
        for fr in c.decode(st):
            total += fr.samples
        return total / float(rate)


def decoded_duration(path):
    """實際解碼計算長度（自測用，比 metadata 可靠）。"""
    import av
    total = 0
    rate = None
    with av.open(path) as c:
        st = c.streams.audio[0]
        for fr in c.decode(st):
            total += fr.samples
            rate = fr.sample_rate
    return total / float(rate or 1)


def wav_duration_bytes(data):
    with wave.open(io.BytesIO(data), "rb") as w:
        return w.getnframes() / float(w.getframerate())


def is_wav(data):
    return len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WAVE"


__all__ = ["read_wav_bytes", "make_wav_bytes", "sine_wave", "concat_wavs", "write_mp3",
           "audio_duration", "decoded_duration", "wav_duration_bytes", "is_wav", "resample",
           "GAP_SECONDS"]
