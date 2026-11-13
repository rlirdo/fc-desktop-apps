# -*- coding: utf-8 -*-
"""客語語音小幫手 HakkaVoiceHelper 1.0.0

給幼兒園客語老師的三步驟小工具（畫面與草稿交給工具，客語定稿交給老師）：
  ① 華語變客語文字：把華語放進剪貼簿、開官方翻譯網頁，由老師本人翻譯後貼回（輔助交接）。
     本工具不會用程式連客委會、長問翻譯器、Gohakka 三個網站（只開瀏覽器＋剪貼簿）。
  ② 客語文字變語音 mp3：臺灣客語語音合成系統（Hugging Face Space 公開 API，CC BY-NC 4.0）。
  ③ 客語語音變文字（試用）：本機離線 faster-whisper，音檔不上傳；結果偏華語、需校對。

命令列：
  HakkaVoiceHelper.exe --version
  HakkaVoiceHelper.exe --selftest            離線自我測試（結果寫 selftest\\selftest.log）
  HakkaVoiceHelper.exe --selftest-online     實際合成一句「恁早。」並驗證長度
  HakkaVoiceHelper.exe --tts "客語漢字" [--accent sixian] [--speaker 標籤或值] [--speed 1] --out x.mp3
  HakkaVoiceHelper.exe --asr 音檔 [--model 模型資料夾] [--out x.txt]
  HakkaVoiceHelper.exe --shots 資料夾        以示範資料產生教學截圖（不連網、不跑模型）
"""
import datetime
import os
import queue
import sys
import tempfile
import threading
import time
import traceback
import webbrowser

import core_asr
import core_audio
import core_tts

APP_NAME = "客語語音小幫手"
EXE_NAME = "HakkaVoiceHelper"
VERSION = "1.0.0"
CREDIT = "NDHU 自資系 游豐兆 製作"          # 右下角製作者字樣
MOTTO = "畫面與草稿交給工具，客語定稿交給老師｜AI 生成，教師審核"

URL_MOE_TRANSLATE = "https://speech.hakka.gov.tw/Translation/Online"
URL_BRONCI = "https://translate.bronci.com.tw/"
URL_GOHAKKA = "https://gohakka.org/translate"
URL_MOE_TTS = "https://speech.hakka.gov.tw/TTS/Synthesis"
URL_MOE_ASR = "https://speech.hakka.gov.tw/AsR/Recognition"

SRC_MOE = "客委會即時翻譯（speech.hakka.gov.tw）"
SRC_BRONCI = "長問翻譯器（translate.bronci.com.tw）"
SRC_GOHAKKA = "Gohakka（gohakka.org）"
SRC_SELF = "自行輸入"

RULES_TEXT = (
    "本工具不會自動連這三個網站：只幫您把文字放進剪貼簿並開啟官方網頁，"
    "翻譯由老師本人在官方網頁操作。\n\n"
    "●客委會（臺灣客語語音資料庫）：翻譯、語音合成、語音辨識 API 須先申請核准"
    "（限政府機關、學校、公益法人，非營利，每次三個月，不得轉授權）；"
    "網站非會員每次 100 字、線上即時翻譯每日 1,000 字。\n\n"
    "●Gohakka：免費會員翻譯每日 20 字（需登入）；API 僅 Enterprise 方案；"
    "內容 CC BY-SA 4.0；使用其語音須註明「語音來源：GoHakka.org」；不得擷取資源訓練 AI。\n\n"
    "●長問翻譯器：Beta 版，翻譯內容由 AI 產生；沒有公開 API。\n\n"
    "客語用字、文法與句法以官方工具結果為準，最後由老師定稿。")

ASR_NOTICE = ("在這台電腦上離線辨識，音檔不會上傳。結果是華語為主的繁體中文，"
              "只能抓大意、常會聽錯，請老師對照原音校對。短句、發音清楚、教材式的句子比較準。")
TTS_NOTICE = ("文字會送到公開的語音合成示範服務，請勿輸入幼兒姓名或任何個資；"
              "這是共用服務，請勿大量連續使用。")

AUDIO_TYPES = [("音檔", "*.mp3 *.wav *.m4a *.aac *.ogg *.flac"), ("所有檔案", "*.*")]

# --shots 示範內容（示意路徑，不出現真實使用者路徑）
DEMO_DIR = r"C:\Users\老師\Desktop\客語語音"
DEMO_MANDARIN = "早安，小兔子要去學校上課。"
DEMO_HAKKA = "恁早啊，細兔仔愛去學校上課。"
DEMO_SEG = "恁早 啊 ， 細兔仔 愛去 學校 上課 。"
DEMO_PIN = "an31 zo31 a55，se55 tu55 e31 oi55 hi55 hog5 gau31 song24 ko55。"
DEMO_ASR_FILE = "客語語音_示範.mp3"
DEMO_ASR_TEXT = "早安，小兔子愛去學校上課"


# --------------------------------------------------------------------------
# 記錄與路徑
# --------------------------------------------------------------------------
_LOG = []


def out(msg=""):
    _LOG.append(str(msg))
    try:
        if sys.stdout is not None:
            print(msg)
            sys.stdout.flush()
    except Exception:  # noqa: BLE001
        pass


def write_log(name):
    """寫到程式旁 selftest\\<name>；不能寫就改寫暫存資料夾。回傳路徑。"""
    for d in (os.path.join(core_asr.app_dir(), "selftest"),
              os.path.join(tempfile.gettempdir(), EXE_NAME + "_selftest")):
        try:
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, name)
            with open(p, "w", encoding="utf-8") as f:
                f.write("\n".join(_LOG) + "\n")
            return p
        except Exception:  # noqa: BLE001
            continue
    return None


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def attach_console():
    """console=False 封裝的 exe 從 cmd 執行時，把輸出接回父主控台（失敗就只寫 log）。"""
    if sys.platform != "win32" or not getattr(sys, "frozen", False) or sys.stdout is not None:
        return
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        if k32.AttachConsole(-1):
            cp = k32.GetConsoleOutputCP() or 65001
            enc = "utf-8" if cp == 65001 else "cp%d" % cp
            sys.stdout = open("CONOUT$", "w", encoding=enc, errors="replace")
            sys.stderr = sys.stdout
    except Exception:  # noqa: BLE001
        pass


def default_out_dir():
    home = os.path.expanduser("~")
    for d in (os.path.join(home, "Desktop"), os.path.join(home, "OneDrive", "Desktop"),
              os.path.join(home, "OneDrive", "桌面")):
        if os.path.isdir(d):
            return os.path.join(d, "客語語音")
    return os.path.join(home, "客語語音")


def open_path(p):
    try:
        if sys.platform == "win32":
            os.startfile(p)  # noqa: S606
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", p])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", p])
    except Exception:  # noqa: BLE001
        pass


def find_speaker(speakers, want, default):
    """--speaker 可給：API 值、網頁標籤、友善標籤、或編號（3 / 語者 3）。"""
    if not want:
        return default
    w = str(want).strip()
    for i, s in enumerate(speakers):
        if w in (s["value"], s["raw_label"], s["label"], str(i + 1), "語者 %d" % (i + 1)):
            return s["value"]
    raise core_tts.TTSError("找不到語者：%s（可用編號 1–%d）" % (want, len(speakers)))


def speaker_label_of(speakers, value):
    for s in speakers:
        if s["value"] == value:
            return s["label"]
    return "預設語者"


def save_tts_result(results, out_path, *, mandarin, hakka, accent, speaker_label, speed,
                    translate_source):
    """把逐句結果串接存 mp3，並寫 <檔名>_來源說明.txt。回傳 (mp3, note, 秒數)。"""
    samples, sr = core_audio.concat_wavs([r[3] for r in results])
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    core_audio.write_mp3(samples, sr, out_path)
    note = os.path.splitext(out_path)[0] + "_來源說明.txt"
    with open(note, "w", encoding="utf-8") as f:
        f.write(core_tts.source_note(
            mandarin=mandarin, hakka=hakka, segs=[r[1] for r in results],
            pinyins=[r[2] for r in results], accent=accent, speaker_label=speaker_label,
            speed=speed, translate_source=translate_source))
    return out_path, note, len(samples) / float(sr)


# --------------------------------------------------------------------------
# 自我測試
# --------------------------------------------------------------------------
def run_selftest():
    ok = True

    def check(name, cond, detail=""):
        nonlocal ok
        out(("  [通過] " if cond else "  [失敗] ") + name + (("：" + detail) if detail else ""))
        if not cond:
            ok = False

    out("%s %s 離線自我測試（%s）" % (APP_NAME, VERSION, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    out("Python %s / %s / frozen=%s" % (sys.version.split()[0], sys.platform, getattr(sys, "frozen", False)))

    # 1. 切句
    s = core_tts.split_sentences("恁早啊，細兔仔愛去學校上課。你好無？\n好！")
    check("切句（句號、問號、換行、驚嘆號）", s == ["恁早啊，細兔仔愛去學校上課。", "你好無？", "好！"], str(s))
    long = "，".join(["細兔仔愛去學校"] * 12) + "。"
    s2 = core_tts.split_sentences(long)
    check("長句依逗號再切（每句 ≤60 字）", len(s2) >= 2 and all(len(x) <= 60 for x in s2)
          and "".join(s2) == long, "%d 段" % len(s2))
    s3 = core_tts.split_sentences("無標點" * 30)
    check("無標點長句硬切", all(len(x) <= 60 for x in s3) and "".join(s3) == "無標點" * 30)
    check("空白與純標點略過", core_tts.split_sentences(" \n。，\n") == [])

    # 2. 檔名清理
    fn = core_tts.safe_filename('恁早啊，細兔仔/愛去:學校*上課?"<>|。')
    check("檔名清理", fn == "恁早啊細兔仔愛去學校上課"[:12] and not any(c in fn for c in '\\/:*?"<>|'), fn)
    check("檔名空白時預設", core_tts.safe_filename("。，") == "客語語音")
    tmp = tempfile.mkdtemp(prefix="HakkaVoiceHelper_selftest_")
    open(os.path.join(tmp, "測試.mp3"), "wb").close()
    p2 = core_tts.unique_path(tmp, "測試")
    check("自動加序號避免覆蓋", os.path.basename(p2) == "測試_2.mp3", os.path.basename(p2))

    # 3. 來源說明檔
    note = core_tts.source_note(mandarin="早安", hakka="恁早", segs=["恁早"], pinyins=["an31 zo31"],
                                accent="hailu", speaker_label="語者 1（女／苗栗）", speed=1.0,
                                translate_source=SRC_GOHAKKA)
    for kw in ("【華語原句】", "恁早", "an31 zo31", "腔調：海陸", "語者 1", "語速：1.0", "CC BY-NC 4.0",
               "yourtts-htia-240704", "GoHakka.org（CC BY-SA 4.0）", "AI 生成，教師審核", "產生時間："):
        check("來源說明含「%s」" % kw, kw in note)
    note2 = core_tts.source_note(mandarin="", hakka="恁早", segs=[], pinyins=[], accent="sixian",
                                 speaker_label="x", speed=1, translate_source=SRC_MOE)
    check("非 Gohakka 不加 CC BY-SA 行、無華語時不列華語", "CC BY-SA" not in note2 and "【華語原句】" not in note2)

    # 4. WAV → MP3 → 解碼長度
    try:
        w1 = core_audio.make_wav_bytes(core_audio.sine_wave(1.2, 22050), 22050)
        w2 = core_audio.make_wav_bytes(core_audio.sine_wave(0.8, 16000, 660), 16000)
        res = [("a", "a", "a", w1), ("b", "b", "b", w2)]
        mp3 = os.path.join(tmp, "合成正弦波.mp3")
        _, notep, secs = save_tts_result(res, mp3, mandarin="", hakka="測試", accent="sixian",
                                         speaker_label="測試", speed=1, translate_source=SRC_SELF)
        expect = 1.2 + 0.8 + core_audio.GAP_SECONDS
        got = core_audio.decoded_duration(mp3)
        err = abs(got - expect) / expect
        check("WAV→MP3 編碼再解碼，長度誤差 <5%", err < 0.05,
              "預期 %.2f 秒、實得 %.2f 秒（誤差 %.1f%%）" % (expect, got, err * 100))
        check("來源說明檔已寫出", os.path.exists(notep))
    except Exception as e:  # noqa: BLE001
        check("WAV→MP3 編碼", False, "%s\n%s" % (e, traceback.format_exc()))

    # 5. 模型搜尋（不載入）
    try:
        c = core_asr.candidate_dirs()
        check("模型搜尋路徑可產生", len(c) >= 1 and c[0].endswith(core_asr.MODEL_DIRNAME))
        check("不存在的模型資料夾判定為無效", core_asr.find_model(os.path.join(tmp, "nope")) is None)
    except Exception as e:  # noqa: BLE001
        check("模型搜尋", False, str(e))

    # 6. 介面建立並關閉
    try:
        import tkinter as tk
        try:
            root = tk.Tk()
        except tk.TclError as e:
            out("  [略過] 無顯示環境，介面測試略過：%s" % e)
            root = None
        if root is not None:
            app = App(root, offline=True)
            root.update()
            ok_gui = len(app.nb.tabs()) == 3
            root.destroy()
            check("介面可建立（三個分頁）並關閉", ok_gui)
    except Exception as e:  # noqa: BLE001
        check("介面建立", False, "%s\n%s" % (e, traceback.format_exc()))

    out("SELFTEST OK" if ok else "SELFTEST FAILED")
    return 0 if ok else 1


def run_selftest_online():
    out("%s %s 線上測試：合成「恁早。」" % (APP_NAME, VERSION))
    try:
        sp, default = core_tts.fetch_speakers()
        out("  取得語者 %d 位、預設：%s" % (len(sp), speaker_label_of(sp, default)))
        t = time.time()
        res = core_tts.synth_text("恁早。", default, "sixian", 1.0, progress=out)
        mp3 = os.path.join(tempfile.mkdtemp(prefix="HakkaVoiceHelper_online_"), "恁早.mp3")
        _, _, secs = save_tts_result(res, mp3, mandarin="", hakka="恁早。", accent="sixian",
                                     speaker_label=speaker_label_of(sp, default), speed=1.0,
                                     translate_source=SRC_SELF)
        dur = core_audio.decoded_duration(mp3)
        out("  斷詞：%s" % res[0][1])
        out("  拼音：%s" % res[0][2])
        out("  mp3 長度 %.2f 秒、%d bytes、耗時 %.1f 秒" % (dur, os.path.getsize(mp3), time.time() - t))
        if dur > 0.5:
            out("SELFTEST-ONLINE OK")
            return 0
        out("SELFTEST-ONLINE FAILED：長度不足 0.5 秒")
    except Exception as e:  # noqa: BLE001
        out("SELFTEST-ONLINE FAILED：%s" % e)
    return 1


def run_tts_cli(text, accent, speaker, speed, out_path):
    try:
        if accent not in core_tts.ACCENT_ZH:
            raise core_tts.TTSError("腔調要是：" + "、".join(core_tts.ACCENT_ZH))
        sp, default = core_tts.fetch_speakers()
        val = find_speaker(sp, speaker, default)
        res = core_tts.synth_text(text, val, accent, float(speed), progress=out)
        if not out_path:
            out_path = os.path.join(os.getcwd(), core_tts.safe_filename(text) + ".mp3")
        mp3, note, secs = save_tts_result(res, out_path, mandarin="", hakka=text, accent=accent,
                                          speaker_label=speaker_label_of(sp, val), speed=float(speed),
                                          translate_source=SRC_SELF)
        for r in res:
            out("斷詞：" + r[1])
            out("拼音：" + r[2])
        out("已存：%s（%.1f 秒）" % (mp3, secs))
        out("來源說明：%s" % note)
        return 0
    except Exception as e:  # noqa: BLE001
        out("失敗：%s" % e)
        return 1


def run_asr_cli(audio, model_dir, out_path):
    try:
        path = core_asr.find_model(model_dir)
        if not path:
            out(core_asr.MISSING_MSG)
            return 2
        t = time.time()
        m = core_asr.load_model(path)
        out("模型載入 %.1f 秒" % (time.time() - t))
        t = time.time()
        text = core_asr.transcribe_text(m, audio)
        out("辨識 %.1f 秒" % (time.time() - t))
        out("辨識結果：")
        out(text)
        if out_path:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(text + "\n")
        return 0 if text.strip() else 1
    except Exception as e:  # noqa: BLE001
        out("失敗：%s" % e)
        return 1


# --------------------------------------------------------------------------
# 介面
# --------------------------------------------------------------------------
def pick_font():
    import tkinter.font as tkfont
    try:
        fams = set(tkfont.families())
    except Exception:  # noqa: BLE001
        fams = set()
    for f in ("Microsoft JhengHei UI", "Microsoft JhengHei", "微軟正黑體", "PingFang TC",
              "Heiti TC", "Noto Sans CJK TC"):
        if f in fams:
            return f
    return "TkDefaultFont"


class App:
    def __init__(self, root, offline=False, demo=False):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk = tk, ttk
        self.root = root
        self.offline = offline or demo
        self.demo = demo
        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.speakers, self.speaker_default = [], None
        self.translate_source = SRC_SELF
        self.last_mp3 = None
        self.audio_path = None
        self.audio_secs = 0.0
        self.model_dir = None
        self.fn_user_edited = False

        root.title("%s %s" % (APP_NAME, VERSION))
        scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        root.geometry("%dx%d" % (int(980 * scale), int(720 * scale)))
        root.minsize(int(820 * scale), int(600 * scale))
        try:
            ico = resource_path("icon.ico")
            if sys.platform == "win32" and os.path.exists(ico):
                root.iconbitmap(ico)
        except Exception:  # noqa: BLE001
            pass

        fam = pick_font()
        self.F = (fam, 11)
        self.FB = (fam, 11, "bold")
        self.FS = (fam, 9)
        self.FT = (fam, 14, "bold")
        self.FE = (fam, 13)
        st = ttk.Style(root)
        try:
            st.theme_use("vista" if sys.platform == "win32" else st.theme_use())
        except Exception:  # noqa: BLE001
            pass
        st.configure(".", font=self.F)
        st.configure("TNotebook.Tab", font=self.FB, padding=(14, 6))
        st.configure("Primary.TButton", font=self.FB)
        st.configure("TLabelframe.Label", font=self.FB, foreground="#065A82")

        # 底部列先佔位（避免被擠出視窗）
        foot = tk.Frame(root)
        foot.pack(side="bottom", fill="x", padx=12, pady=(0, 6))
        tk.Label(foot, text=MOTTO, font=self.FS, fg="#64748B", anchor="w").pack(side="left")
        tk.Label(foot, text=CREDIT, font=self.FS, fg="#204a87", anchor="e").pack(side="right")

        status = tk.Frame(root)
        status.pack(side="bottom", fill="x", padx=12, pady=(2, 2))
        self.v_status = tk.StringVar(value="準備好了。")
        tk.Label(status, textvariable=self.v_status, font=self.F, fg="#1E293B", anchor="w",
                 justify="left").pack(side="left", fill="x", expand=True)
        self.btn_cancel = ttk.Button(status, text="取消", command=self.on_cancel, state="disabled")
        self.btn_cancel.pack(side="right")

        self.nb = ttk.Notebook(root)
        self.nb.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 2))
        self.tab1 = ttk.Frame(self.nb, padding=10)
        self.tab2 = ttk.Frame(self.nb, padding=10)
        self.tab3 = ttk.Frame(self.nb, padding=10)
        self.nb.add(self.tab1, text="① 華語變客語文字")
        self.nb.add(self.tab2, text="② 客語文字變語音 mp3")
        self.nb.add(self.tab3, text="③ 客語語音變文字（試用）")
        self.build_tab1()
        self.build_tab2()
        self.build_tab3()

        root.after(100, self.poll)
        root.after(400, self.refresh)
        if not self.offline:
            self.load_speakers()

    # ------------------------------------------------------------ 共用
    def text(self, parent, height, **kw):
        tk = self.tk
        fr = tk.Frame(parent)
        t = tk.Text(fr, height=height, font=self.FE, wrap="word", undo=True, relief="solid",
                    borderwidth=1, padx=6, pady=4, **kw)
        sb = self.ttk.Scrollbar(fr, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(side="left", fill="both", expand=True)
        return fr, t

    @staticmethod
    def get(t):
        return t.get("1.0", "end-1c")

    @staticmethod
    def set(t, s):
        st = str(t.cget("state"))
        t.configure(state="normal")
        t.delete("1.0", "end")
        t.insert("1.0", s)
        t.configure(state=st)

    def status(self, s):
        self.v_status.set(s)

    def to_clipboard(self, s):
        self.root.clipboard_clear()
        self.root.clipboard_append(s)
        self.root.update()

    def start_job(self, fn, *a):
        if self.busy:
            return
        self.busy = True
        self.cancel.clear()
        self.btn_cancel.configure(state="normal")

        def run():
            try:
                fn(*a)
            except Exception as e:  # noqa: BLE001
                self.q.put(("error", str(e)))
            finally:
                self.q.put(("done", None))
        threading.Thread(target=run, daemon=True).start()

    def on_cancel(self):
        self.cancel.set()
        self.status("取消中…（目前這一步做完就停）")

    def poll(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "done":
                    self.busy = False
                    self.btn_cancel.configure(state="disabled")
                elif kind == "status":
                    self.status(data)
                elif kind == "error":
                    self.status("發生錯誤：%s" % data)
                elif callable(kind):
                    kind(*data) if isinstance(data, tuple) else kind(data)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def ui(self, fn, *a):
        self.q.put((fn, a))

    def refresh(self):
        """定時更新字數、提醒與按鈕狀態。"""
        try:
            m = self.get(self.t_man).strip()
            h = self.get(self.t_hak).strip()
            n = len(m)
            self.v_count.set("%d 字" % n + ("　（客委會非會員每次 100 字，請分段）" if n > 100 else ""))
            self.lbl_count.configure(fg="#B45309" if n > 100 else "#64748B")
            if h and m and h == m:
                self.v_hint.set("看起來還沒翻譯（客語框和華語框一模一樣）。")
            elif not h:
                self.v_hint.set("客語漢字是空的：請先到官方網頁翻譯後貼上，或直接打客語漢字。")
            else:
                self.v_hint.set("")
            self.btn_next.configure(state="normal" if h else "disabled")
            h2 = self.get(self.t_tts).strip()
            if not self.fn_user_edited:
                self.v_fname.set(core_tts.safe_filename(h2) if h2 else "")
            can = bool(h2) and bool(self.speakers) and not self.busy
            self.btn_gen.configure(state="normal" if can else "disabled")
            self.btn_asr.configure(state="normal" if (self.audio_path and not self.busy) else "disabled")
        except Exception:  # noqa: BLE001
            pass
        self.root.after(400, self.refresh)

    # ------------------------------------------------------------ 分頁 1
    def build_tab1(self):
        tk, ttk = self.tk, self.ttk
        f = self.tab1
        f.columnconfigure(0, weight=3)
        f.columnconfigure(1, weight=0)
        f.rowconfigure(0, weight=1)
        left = ttk.Frame(f)
        left.grid(row=0, column=0, sticky="nsew")
        self.rules_col = ttk.Frame(f)
        self.rules_col.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        top = ttk.Frame(left)
        top.pack(fill="x")
        tk.Label(top, text="華語句子", font=self.FB, fg="#065A82").pack(side="left")
        self.v_count = tk.StringVar(value="0 字")
        self.lbl_count = tk.Label(top, textvariable=self.v_count, font=self.FS, fg="#64748B")
        self.lbl_count.pack(side="left", padx=8)
        fr, self.t_man = self.text(left, 4)
        fr.pack(fill="both", expand=True, pady=(2, 6))

        bx = ttk.LabelFrame(left, text=" 複製華語並開啟官方翻譯網頁（翻譯由老師本人操作） ", padding=(8, 4))
        bx.pack(fill="x")
        rows = [("複製並開啟：客委會即時翻譯", URL_MOE_TRANSLATE, SRC_MOE, "建議", "#2E7D5B"),
                ("複製並開啟：長問翻譯器", URL_BRONCI, SRC_BRONCI, "Beta，AI 翻譯", "#64748B"),
                ("複製並開啟：Gohakka", URL_GOHAKKA, SRC_GOHAKKA, "需登入，免費每日 20 字", "#64748B")]
        for i, (label, url, src, tip, color) in enumerate(rows):
            ttk.Button(bx, text=label, style="Primary.TButton" if i == 0 else "TButton",
                       command=lambda u=url, s=src: self.copy_and_open(u, s)).grid(
                row=0, column=i, sticky="ew", padx=4, pady=(2, 0))
            tk.Label(bx, text=tip, font=self.FS, fg=color).grid(row=1, column=i, pady=(0, 2))
            bx.columnconfigure(i, weight=1)

        mid = ttk.Frame(left)
        mid.pack(fill="x", pady=(8, 0))
        tk.Label(mid, text="客語漢字", font=self.FB, fg="#065A82").pack(side="left")
        self.v_src = tk.StringVar(value="翻譯來源：%s" % SRC_SELF)
        tk.Label(mid, textvariable=self.v_src, font=self.FS, fg="#64748B").pack(side="left", padx=8)
        fr, self.t_hak = self.text(left, 4)
        fr.pack(fill="both", expand=True, pady=(2, 4))
        self.t_hak.bind("<Key>", lambda e: self.mark_self_typed(e))

        bar = ttk.Frame(left)
        bar.pack(fill="x")
        ttk.Button(bar, text="貼上客語漢字（從剪貼簿）", command=self.paste_hakka).pack(side="left")
        ttk.Button(bar, text="清除", command=self.clear_tab1).pack(side="left", padx=6)
        self.btn_next = ttk.Button(bar, text="下一步：變成語音 →", style="Primary.TButton",
                                   command=self.go_tab2, state="disabled")
        self.btn_next.pack(side="right")
        self.v_hint = tk.StringVar(value="")
        tk.Label(left, textvariable=self.v_hint, font=self.FS, fg="#B45309", anchor="w").pack(fill="x", pady=(4, 0))

        # 可收合的規定說明
        self.v_rules = tk.BooleanVar(value=True)
        self.btn_rules = ttk.Button(self.rules_col, text="▼ 三個網站的使用規定（按此收合）", command=self.toggle_rules)
        self.btn_rules.pack(fill="x")
        self.rules_box = tk.Text(self.rules_col, font=self.FS, fg="#1E293B", wrap="char", width=34,
                                 bg="#F5F9FA", relief="solid", borderwidth=1, padx=8, pady=8,
                                 spacing2=2, cursor="arrow")
        self.rules_box.insert("1.0", RULES_TEXT)
        self.rules_box.configure(state="disabled")
        self.rules_box.pack(fill="both", expand=True, pady=(4, 0))

    def toggle_rules(self):
        if self.v_rules.get():
            self.rules_box.pack_forget()
            self.btn_rules.configure(text="▶ 三個網站的使用規定（按此展開）")
            self.v_rules.set(False)
        else:
            self.rules_box.pack(fill="both", expand=True, pady=(4, 0))
            self.btn_rules.configure(text="▼ 三個網站的使用規定（按此收合）")
            self.v_rules.set(True)

    def set_source(self, src):
        self.translate_source = src
        self.v_src.set("翻譯來源：%s" % src)

    def copy_and_open(self, url, src):
        m = self.get(self.t_man).strip()
        if not m:
            self.status("請先在「華語句子」輸入要翻譯的句子。")
            return
        self.to_clipboard(m)
        self.set_source(src)
        webbrowser.open(url)
        self.status("已複製，請到瀏覽器貼上並按翻譯，複製客語漢字後回來按「貼上客語漢字」。")

    def mark_self_typed(self, e):
        if e.char and e.char.isprintable() and not self.get(self.t_hak).strip():
            self.set_source(SRC_SELF)

    def paste_hakka(self):
        try:
            s = self.root.clipboard_get()
        except Exception:  # noqa: BLE001
            self.status("剪貼簿是空的：請先在官方網頁複製客語漢字。")
            return
        self.set(self.t_hak, s.strip())
        self.status("已貼上客語漢字，請檢查用字後按「下一步：變成語音 →」。")

    def clear_tab1(self):
        self.set(self.t_man, "")
        self.set(self.t_hak, "")
        self.set_source(SRC_SELF)

    def go_tab2(self):
        h = self.get(self.t_hak).strip()
        if not h:
            return
        self.set(self.t_tts, h)
        self.fn_user_edited = False
        self.nb.select(self.tab2)

    # ------------------------------------------------------------ 分頁 2
    def build_tab2(self):
        tk, ttk = self.tk, self.ttk
        f = self.tab2
        tk.Label(f, text="客語漢字（可直接修改）", font=self.FB, fg="#065A82", anchor="w").pack(fill="x")
        fr, self.t_tts = self.text(f, 3)
        fr.pack(fill="x", pady=(2, 6))

        opt = ttk.Frame(f)
        opt.pack(fill="x")
        tk.Label(opt, text="腔調", font=self.FB).grid(row=0, column=0, sticky="w")
        self.v_acc = tk.StringVar(value="sixian")
        af = ttk.Frame(opt)
        af.grid(row=0, column=1, columnspan=3, sticky="w", padx=6)
        for code, zh in core_tts.ACCENTS:
            ttk.Radiobutton(af, text=zh, value=code, variable=self.v_acc).pack(side="left", padx=(0, 10))

        tk.Label(opt, text="語者", font=self.FB).grid(row=1, column=0, sticky="w", pady=4)
        self.v_spk = tk.StringVar(value="")
        self.cb_spk = ttk.Combobox(opt, textvariable=self.v_spk, state="readonly", width=24, font=self.F)
        self.cb_spk.grid(row=1, column=1, sticky="w", padx=6)
        ttk.Button(opt, text="重新整理", command=self.load_speakers).grid(row=1, column=2, sticky="w")
        self.v_spk_msg = tk.StringVar(value="")
        self.lbl_spk = tk.Label(opt, textvariable=self.v_spk_msg, font=self.FS, fg="#64748B")
        self.lbl_spk.grid(row=1, column=3, sticky="w", padx=6)

        tk.Label(opt, text="語速", font=self.FB).grid(row=2, column=0, sticky="w")
        sf = ttk.Frame(opt)
        sf.grid(row=2, column=1, columnspan=3, sticky="w", padx=6)
        self.v_speed = tk.DoubleVar(value=1.0)
        self.v_speed_txt = tk.StringVar(value="1.00")
        ttk.Scale(sf, from_=0.5, to=1.5, variable=self.v_speed, length=220,
                  command=lambda v: self.v_speed_txt.set("%.2f" % (round(float(v) * 20) / 20))).pack(side="left")
        tk.Label(sf, textvariable=self.v_speed_txt, font=self.F, width=5).pack(side="left")
        tk.Label(sf, text="（0.5–1.5，越大越慢）", font=self.FS, fg="#64748B").pack(side="left")

        tk.Label(opt, text="輸出資料夾", font=self.FB).grid(row=3, column=0, sticky="w", pady=4)
        self.v_dir = tk.StringVar(value=DEMO_DIR if self.demo else default_out_dir())
        ttk.Entry(opt, textvariable=self.v_dir, width=48, font=self.F).grid(row=3, column=1, columnspan=2, sticky="we", padx=6)
        ttk.Button(opt, text="選擇…", command=self.pick_dir).grid(row=3, column=3, sticky="w")
        tk.Label(opt, text="檔名", font=self.FB).grid(row=4, column=0, sticky="w")
        self.v_fname = tk.StringVar(value="")
        e = ttk.Entry(opt, textvariable=self.v_fname, width=30, font=self.F)
        e.grid(row=4, column=1, sticky="w", padx=6)
        e.bind("<Key>", lambda ev: setattr(self, "fn_user_edited", True))
        tk.Label(opt, text=".mp3（同名會自動加序號）", font=self.FS, fg="#64748B").grid(row=4, column=2, columnspan=2, sticky="w")

        tk.Label(f, text=TTS_NOTICE, font=self.FS, fg="#B45309", anchor="w", justify="left",
                 wraplength=900).pack(fill="x", pady=(6, 4))

        bar = ttk.Frame(f)
        bar.pack(fill="x")
        self.btn_gen = ttk.Button(bar, text="產生客語語音 mp3", style="Primary.TButton",
                                  command=self.on_generate, state="disabled")
        self.btn_gen.pack(side="left")
        self.btn_play = ttk.Button(bar, text="播放", command=lambda: self.last_mp3 and open_path(self.last_mp3),
                                   state="disabled")
        self.btn_play.pack(side="left", padx=6)
        self.btn_folder = ttk.Button(bar, text="開啟資料夾",
                                     command=lambda: self.last_mp3 and open_path(os.path.dirname(self.last_mp3)),
                                     state="disabled")
        self.btn_folder.pack(side="left")
        self.btn_moe_tts = ttk.Button(bar, text="改用客委會語音合成網頁", command=self.open_moe_tts)

        fr, self.t_res = self.text(f, 6, state="disabled", bg="#F5F9FA")
        self.t_res.configure(font=self.F)
        fr.pack(fill="both", expand=True, pady=(6, 0))

    def pick_dir(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(initialdir=self.v_dir.get() if os.path.isdir(self.v_dir.get()) else None)
        if d:
            self.v_dir.set(os.path.normpath(d))

    def open_moe_tts(self):
        h = self.get(self.t_tts).strip()
        if h:
            self.to_clipboard(h)
        webbrowser.open(URL_MOE_TTS)
        self.status("已複製客語漢字並開啟客委會語音合成網頁，請在網頁貼上後操作。")

    def load_speakers(self):
        if self.offline:
            return
        self.v_spk_msg.set("取得語者清單中…")
        self.lbl_spk.configure(fg="#64748B")

        def work():
            try:
                sp, d = core_tts.fetch_speakers()
                self.ui(self.set_speakers, sp, d)
            except Exception as e:  # noqa: BLE001
                self.ui(self.speakers_failed, str(e))
        threading.Thread(target=work, daemon=True).start()

    def set_speakers(self, sp, default):
        self.speakers, self.speaker_default = sp, default
        self.cb_spk.configure(values=[s["label"] for s in sp])
        self.v_spk.set(speaker_label_of(sp, default))
        self.v_spk_msg.set("共 %d 位語者（執行時取得）" % len(sp))
        self.lbl_spk.configure(fg="#2E7D5B")
        self.btn_moe_tts.pack_forget()

    def speakers_failed(self, msg):
        self.speakers = []
        self.cb_spk.configure(values=[])
        self.v_spk.set("")
        self.v_spk_msg.set("連不上語音合成服務，暫時無法產生")
        self.lbl_spk.configure(fg="#B91C1C")
        self.set(self.t_res, "連線錯誤：%s\n\n可能是網路沒通，或公開示範服務正在休眠（第一次喚醒要等一兩分鐘）。"
                 "請按「重新整理」再試，或按「改用客委會語音合成網頁」。" % msg)
        self.btn_moe_tts.pack(side="right")

    def on_generate(self):
        h = self.get(self.t_tts).strip()
        if not h or not self.speakers:
            return
        lab = self.v_spk.get()
        val = next((s["value"] for s in self.speakers if s["label"] == lab), self.speaker_default)
        folder = os.path.expanduser(self.v_dir.get().strip() or default_out_dir())
        stem = core_tts.safe_filename(self.v_fname.get().strip() or h, n=40)
        acc = self.v_acc.get()
        spd = round(float(self.v_speed.get()) * 20) / 20
        man = self.get(self.t_man).strip()
        src = self.translate_source
        self.set(self.t_res, "")
        self.btn_moe_tts.pack_forget()

        def work():
            res = core_tts.synth_text(h, val, acc, spd, cancel=self.cancel,
                                      progress=lambda s: self.q.put(("status", s)))
            os.makedirs(folder, exist_ok=True)
            path = core_tts.unique_path(folder, stem)
            mp3, note, secs = save_tts_result(res, path, mandarin=man, hakka=h, accent=acc,
                                              speaker_label=lab, speed=spd, translate_source=src)
            self.ui(self.tts_done, res, mp3, note, secs)

        def guarded():
            try:
                work()
            except core_tts.Cancelled:
                self.q.put(("status", "已取消。"))
            except Exception as e:  # noqa: BLE001
                self.ui(self.tts_failed, str(e))
        self.status("開始產生…（一句一句送，每句約 5–10 秒）")
        self.start_job(guarded)

    def show_tts_result(self, segs, pins, mp3, note, secs):
        txt = ("【斷詞】\n" + "\n".join(segs) + "\n\n【客語拼音】\n" + "\n".join(pins)
               + "\n\n【檔案】\n%s（%.1f 秒）\n%s" % (mp3, secs, note))
        self.set(self.t_res, txt)
        self.last_mp3 = mp3
        self.btn_play.configure(state="normal")
        self.btn_folder.configure(state="normal")
        self.status("完成！可以按「播放」聽聽看。請老師確認發音；給孩子聽的請由老師親口唸。")

    def tts_done(self, res, mp3, note, secs):
        self.show_tts_result([r[1] for r in res], [r[2] for r in res], mp3, note, secs)

    def tts_failed(self, msg):
        self.set(self.t_res, "產生失敗：%s\n\n可能原因：網路沒通、公開示範服務正在休眠或忙碌。"
                 "請稍等一兩分鐘再按一次，或按右邊「改用客委會語音合成網頁」。" % msg)
        self.btn_moe_tts.pack(side="right")
        self.status("產生失敗，請看下方說明。")

    # ------------------------------------------------------------ 分頁 3
    def build_tab3(self):
        tk, ttk = self.tk, self.ttk
        f = self.tab3
        tk.Label(f, text=ASR_NOTICE, font=self.FB, fg="#9A3412", bg="#FFF7ED", anchor="w",
                 justify="left", wraplength=900, relief="solid", borderwidth=1, padx=10, pady=8
                 ).pack(fill="x")

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(8, 4))
        ttk.Button(bar, text="開啟客委會語音辨識網頁（官方，建議）", style="Primary.TButton",
                   command=lambda: (webbrowser.open(URL_MOE_ASR),
                                    self.status("已開啟客委會語音辨識網頁，請在網頁上傳音檔。"))).pack(side="left")
        ttk.Button(bar, text="選擇音檔…", style="Primary.TButton", command=self.pick_audio).pack(side="left", padx=8)
        self.btn_asr = ttk.Button(bar, text="開始辨識", style="Primary.TButton", command=self.on_asr, state="disabled")
        self.btn_asr.pack(side="left")

        self.v_audio = tk.StringVar(value="尚未選擇音檔（mp3／wav／m4a／aac／ogg／flac）")
        tk.Label(f, textvariable=self.v_audio, font=self.F, fg="#1E293B", anchor="w").pack(fill="x")

        mf = ttk.Frame(f)
        mf.pack(fill="x", pady=(2, 4))
        self.v_model = tk.StringVar(value="")
        self.lbl_model = tk.Label(mf, textvariable=self.v_model, font=self.FS, fg="#64748B", anchor="w",
                                  justify="left", wraplength=700)
        self.lbl_model.pack(side="left", fill="x", expand=True)
        ttk.Button(mf, text="選擇模型資料夾…", command=self.pick_model).pack(side="right")
        self.update_model_status()

        tk.Label(f, text="辨識結果（可直接修改校對）", font=self.FB, fg="#065A82", anchor="w").pack(fill="x", pady=(4, 0))
        fr, self.t_asr = self.text(f, 8)
        fr.pack(fill="both", expand=True, pady=(2, 4))
        b2 = ttk.Frame(f)
        b2.pack(fill="x")
        ttk.Button(b2, text="複製文字", command=self.copy_asr).pack(side="left")
        ttk.Button(b2, text="另存 .txt", command=self.save_asr).pack(side="left", padx=6)

    def update_model_status(self, path=None):
        if self.demo:
            self.model_dir = "demo"
            self.v_model.set("模型：已找到（程式旁的 models\\%s）" % core_asr.MODEL_DIRNAME)
            self.lbl_model.configure(fg="#2E7D5B")
            return
        p = path or core_asr.find_model()
        self.model_dir = p
        if p:
            where = "程式旁的 models 資料夾" if os.path.normcase(p).startswith(os.path.normcase(core_asr.app_dir())) else p
            self.v_model.set("模型：已找到（%s）" % where)
            self.lbl_model.configure(fg="#2E7D5B")
        else:
            self.v_model.set(core_asr.MISSING_MSG)
            self.lbl_model.configure(fg="#B91C1C")

    def pick_model(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(title="選擇模型資料夾（內含 model.bin）")
        if not d:
            return
        d = os.path.normpath(d)
        if not core_asr.is_model_dir(d):
            self.status("這個資料夾裡沒有 model.bin 與 config.json，不是模型資料夾。")
            return
        s = core_asr.load_settings()
        s["model_dir"] = d
        core_asr.save_settings(s)
        self.update_model_status(d)
        self.status("已設定模型資料夾。")

    def pick_audio(self):
        from tkinter import filedialog
        p = filedialog.askopenfilename(title="選擇音檔", filetypes=AUDIO_TYPES)
        if not p:
            return
        self.audio_path = os.path.normpath(p)
        try:
            secs = core_audio.audio_duration(self.audio_path)
        except Exception as e:  # noqa: BLE001
            self.audio_path = None
            self.v_audio.set("讀不到這個音檔：%s" % e)
            return
        self.audio_secs = secs
        msg = "音檔：%s（長度 %s）" % (os.path.basename(p), fmt_secs(secs))
        if secs > 600:
            msg += "　超過 10 分鐘，辨識會比較久"
        self.v_audio.set(msg)

    def on_asr(self):
        if not self.audio_path:
            return
        if not self.model_dir or not core_asr.is_model_dir(self.model_dir):
            self.update_model_status()
            if not self.model_dir:
                self.status("找不到模型資料夾，請看上方說明。")
                return
        self.set(self.t_asr, "")
        path, mdir = self.audio_path, self.model_dir

        def work():
            try:
                if mdir not in core_asr._MODEL_CACHE:
                    self.q.put(("status", "載入模型中（約十幾秒）…"))
                m = core_asr.load_model(mdir)
                total = 0.0
                for item in core_asr.transcribe_iter(m, path):
                    if self.cancel.is_set():
                        self.q.put(("status", "已取消。"))
                        return
                    if item[0] == "info":
                        total = item[1]
                        self.q.put(("status", "辨識中… 0／%s" % fmt_secs(total)))
                        continue
                    self.ui(self.asr_append, item[2])
                    self.q.put(("status", "辨識中… %s／%s" % (fmt_secs(item[1]), fmt_secs(total))))
                self.q.put(("status", "辨識完成。請對照原音校對（結果偏華語、常會聽錯）。"))
            except Exception as e:  # noqa: BLE001
                self.q.put(("status", "辨識失敗：%s" % e))
        self.start_job(work)

    def asr_append(self, s):
        if not s:
            return
        cur = self.get(self.t_asr)
        self.t_asr.insert("end", ("\n" if cur else "") + s)
        self.t_asr.see("end")

    def copy_asr(self):
        s = self.get(self.t_asr).strip()
        if s:
            self.to_clipboard(s)
            self.status("已複製辨識文字。")

    def save_asr(self):
        from tkinter import filedialog
        s = self.get(self.t_asr).strip()
        if not s:
            return
        init = (os.path.splitext(os.path.basename(self.audio_path))[0] if self.audio_path else "辨識結果") + "_辨識.txt"
        p = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=init,
                                         filetypes=[("文字檔", "*.txt")])
        if p:
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(s + "\n")
            self.status("已存：%s" % p)

    # ------------------------------------------------------------ 截圖示範
    def fill_demo(self):
        self.set(self.t_man, DEMO_MANDARIN)
        self.set(self.t_hak, DEMO_HAKKA)
        self.set_source(SRC_MOE)
        self.set(self.t_tts, DEMO_HAKKA)
        demo_sp = [{"value": "demo", "raw_label": "", "label": "語者 1（女／苗栗）"}]
        self.speakers, self.speaker_default = demo_sp, "demo"
        self.cb_spk.configure(values=[demo_sp[0]["label"]])
        self.v_spk.set(demo_sp[0]["label"])
        self.v_spk_msg.set("共 18 位語者（執行時取得）")
        self.lbl_spk.configure(fg="#2E7D5B")
        self.audio_path = os.path.join(DEMO_DIR, DEMO_ASR_FILE)
        self.v_audio.set("音檔：%s（長度 0:03）" % DEMO_ASR_FILE)


def fmt_secs(s):
    s = int(round(s or 0))
    return "%d:%02d" % (s // 60, s % 60)


def window_bbox(root):
    """整個視窗（含標題列，不含陰影）的螢幕座標。"""
    root.update()
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            hwnd = int(root.wm_frame(), 16)
            rect = wintypes.RECT()
            ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(rect), ctypes.sizeof(rect))
            if rect.right > rect.left:
                return (rect.left, rect.top, rect.right, rect.bottom)
        except Exception:  # noqa: BLE001
            pass
    x, y = root.winfo_rootx(), root.winfo_rooty()
    return (x, y, x + root.winfo_width(), y + root.winfo_height())


def run_shots(folder):
    import tkinter as tk
    from PIL import ImageGrab
    os.makedirs(folder, exist_ok=True)
    root = tk.Tk()
    app = App(root, demo=True)
    root.geometry("+40+40")
    root.attributes("-topmost", True)
    app.fill_demo()
    files = []

    def grab(name):
        root.lift()
        root.update()
        time.sleep(0.8)
        root.update()
        img = ImageGrab.grab(bbox=window_bbox(root), all_screens=True)
        p = os.path.join(folder, name)
        img.save(p)
        files.append(p)
        out("已存：%s（%dx%d）" % (p, img.width, img.height))

    app.nb.select(app.tab1)
    app.status("已複製，請到瀏覽器貼上並按翻譯，複製客語漢字後回來按「貼上客語漢字」。")
    grab("shot1_華語變客語.png")
    app.nb.select(app.tab2)
    app.status("準備好了：確認客語漢字、腔調與語者後，按「產生客語語音 mp3」。")
    grab("shot2_文字變語音.png")
    mp3 = DEMO_DIR + "\\" + core_tts.safe_filename(DEMO_HAKKA) + ".mp3"
    note = DEMO_DIR + "\\" + core_tts.safe_filename(DEMO_HAKKA) + "_來源說明.txt"
    app.show_tts_result([DEMO_SEG], [DEMO_PIN], mp3, note, 3.3)
    grab("shot2b_產生完成.png")
    app.nb.select(app.tab3)
    app.set(app.t_asr, DEMO_ASR_TEXT)
    app.status("辨識完成。請對照原音校對（結果偏華語、常會聽錯）。")
    grab("shot3_語音變文字.png")
    root.destroy()
    return 0 if len(files) == 4 else 1


# --------------------------------------------------------------------------
# 進入點
# --------------------------------------------------------------------------
def _arg(args, name, default=None):
    if name in args:
        i = args.index(name)
        if i + 1 < len(args):
            return args[i + 1]
    return default


def set_dpi_aware():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:  # noqa: BLE001
                pass


def main():
    args = sys.argv[1:]
    if args:
        attach_console()
    if "--version" in args:
        out("%s %s %s" % (EXE_NAME, APP_NAME, VERSION))
        write_log("version.log")
        return 0
    if "--selftest" in args:
        set_dpi_aware()
        code = run_selftest()
        p = write_log("selftest.log")
        out("（記錄：%s）" % p)
        return code
    if "--selftest-online" in args:
        code = run_selftest_online()
        write_log("selftest_online.log")
        return code
    if "--tts" in args:
        code = run_tts_cli(_arg(args, "--tts", ""), _arg(args, "--accent", "sixian"),
                           _arg(args, "--speaker"), _arg(args, "--speed", "1"), _arg(args, "--out"))
        write_log("tts.log")
        return code
    if "--asr" in args:
        code = run_asr_cli(_arg(args, "--asr"), _arg(args, "--model"), _arg(args, "--out"))
        write_log("asr.log")
        return code
    if "--shots" in args:
        set_dpi_aware()
        code = run_shots(_arg(args, "--shots", os.path.join(core_asr.app_dir(), "shots")))
        write_log("shots.log")
        return code

    set_dpi_aware()
    import tkinter as tk
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
