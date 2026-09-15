# -*- coding: utf-8 -*-
"""
app.py — 學生作業文字雲（HomeworkWordCloud）桌面版入口

給課程教學助理（TA）用：把「姓名遮罩與學生編號」處理過的 xlsx
（或原始 Zuvio 匯出、通用 CSV）變成「同學作答分析」簡報＋文字雲＋概念矩陣＋提問分類。
輸出一律使用學生編號（例 115-1_EC_3），姓名已遮罩、學號與 email 全部改成 O。

雙擊 exe → 出現視窗；也可以用命令列：
    HomeworkWordCloud.exe --selftest                          無視窗自我測試
                                                              （成功印 SELFTEST OK、回傳碼 0）
    HomeworkWordCloud.exe --check-privacy <輸出資料夾>         隱私檢查
                                                              [--input <原始輸入資料夾>]
    HomeworkWordCloud.exe --version                           印版本
"""
import os
import sys
import queue
import threading
import traceback
import datetime as dt

# 打包成 --windowed 時 sys.stdout / sys.stderr 會是 None，
# 任何第三方套件的 print / logging 都可能炸掉，先補上安全的空輸出。
class _NullIO:
    def write(self, *a, **k):
        return 0

    def flush(self):
        pass

    def isatty(self):
        return False


if sys.stdout is None:
    sys.stdout = _NullIO()
if sys.stderr is None:
    sys.stderr = _NullIO()

sys.path.insert(0, getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))))

import pipeline as P                                              # noqa: E402


# ------------------------------------------------------------------ CLI
def cli_log_factory(fh):
    def log(*a):
        line = " ".join(str(x) for x in a)
        try:
            fh.write(line + "\n")
            fh.flush()
        except Exception:
            pass
        try:
            print(line, flush=True)
        except Exception:
            pass
    return log


def run_selftest():
    """--selftest：無 GUI，成功印 SELFTEST OK 並回傳 0；同時寫一份 log 檔。"""
    import tempfile
    log_path = os.path.join(tempfile.gettempdir(), f"{P.APP_NAME}_selftest.log")
    code = 1
    with open(log_path, "w", encoding="utf-8") as fh:
        log = cli_log_factory(fh)
        log(f"{P.APP_TITLE} {P.VERSION} selftest　{dt.datetime.now():%Y-%m-%d %H:%M:%S}")
        log(f"python={sys.version.split()[0]}　frozen={P.is_frozen()}　platform={sys.platform}")
        try:
            code = P.selftest(log=log)
        except Exception:
            log(traceback.format_exc())
            log("SELFTEST FAILED")
            code = 1
        log(f"exit code = {code}")
    try:
        print(f"（詳細紀錄：{log_path}）", flush=True)
    except Exception:
        pass
    return code


def _opt(argv, name):
    """取命令列選項的值：`--name 值` 或 `--name=值`。"""
    for i, a in enumerate(argv):
        if a == name and i + 1 < len(argv):
            return argv[i + 1]
        if a.startswith(name + "="):
            return a.split("=", 1)[1]
    return ""


def run_check_privacy(argv):
    """--check-privacy <輸出資料夾> [--input <原始輸入資料夾>]

    掃描輸出資料夾，比對名冊姓名、9 碼學號與電子郵件樣式；0 命中才回傳 0。
    """
    target = _opt(argv, "--check-privacy")
    if not target:
        rest = [a for a in argv if not a.startswith("-")]
        target = rest[0] if rest else ""
    src = _opt(argv, "--input") or None
    log = lambda *a: print(" ".join(str(x) for x in a), flush=True)   # noqa: E731
    if not target:
        log("用法：HomeworkWordCloud.exe --check-privacy <輸出資料夾> [--input <輸入資料夾>]")
        return 2
    try:
        hits = P.check_privacy(target, input_dir=src, log=log)
    except P.PipelineError as e:
        log(str(e))
        return 2
    except Exception:
        log(traceback.format_exc())
        return 2
    return 1 if hits else 0


# ------------------------------------------------------------------ GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    cfg = P.load_config()
    msgq = queue.Queue()

    root = tk.Tk()
    root.title(f"{P.APP_TITLE}　{P.APP_NAME} {P.VERSION}")
    root.geometry("1040x720")
    root.minsize(900, 600)
    try:
        ico = P.resource_path("icon.ico")
        if os.path.exists(ico) and sys.platform.startswith("win"):
            root.iconbitmap(ico)
    except Exception:
        pass

    style = ttk.Style()
    try:
        style.theme_use("vista" if sys.platform.startswith("win") else "clam")
    except Exception:
        pass

    pad = {"padx": 6, "pady": 4}
    body = ttk.Frame(root, padding=10)
    body.pack(fill="both", expand=True)

    # ---------------- 課程資訊
    info = ttk.LabelFrame(body, text="課程資訊（會印在簡報封面）", padding=8)
    info.pack(fill="x")
    for c in (1, 3):
        info.columnconfigure(c, weight=1)

    v_course = tk.StringVar(value=cfg.get("course_name", ""))
    v_teacher = tk.StringVar(value=cfg.get("teacher", ""))
    v_ta = tk.StringVar(value=cfg.get("ta_name", ""))
    v_start = tk.StringVar(value=cfg.get("semester_start", ""))
    v_code = tk.StringVar(value=cfg.get("course_code", "EC"))
    v_sem = tk.StringVar(value=cfg.get("semester", "115-1"))

    ttk.Label(info, text="課程名稱").grid(row=0, column=0, sticky="e", **pad)
    ttk.Entry(info, textvariable=v_course).grid(row=0, column=1, sticky="ew", **pad)
    ttk.Label(info, text="教師").grid(row=0, column=2, sticky="e", **pad)
    ttk.Entry(info, textvariable=v_teacher).grid(row=0, column=3, sticky="ew", **pad)
    ttk.Label(info, text="教學助理（TA）").grid(row=1, column=0, sticky="e", **pad)
    ttk.Entry(info, textvariable=v_ta).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Label(info, text="學期起日（第 1 週的星期日）").grid(row=1, column=2, sticky="e", **pad)
    ttk.Entry(info, textvariable=v_start).grid(row=1, column=3, sticky="ew", **pad)

    idrow = ttk.Frame(info)
    idrow.grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(2, 0))
    ttk.Label(idrow, text="課程縮寫（2 個英文字母）").pack(side="left")
    ttk.Entry(idrow, textvariable=v_code, width=6).pack(side="left", padx=(6, 14))
    ttk.Label(idrow, text="學期（例 115-1）").pack(side="left")
    ttk.Entry(idrow, textvariable=v_sem, width=8).pack(side="left", padx=(6, 14))
    v_codehint = tk.StringVar(value="")
    ttk.Label(idrow, textvariable=v_codehint, foreground="#065A82").pack(side="left")

    def refresh_hint(*_a):
        code = P.norm_course_code(v_code.get())
        sem = P.norm_semester(v_sem.get())
        bad = []
        if code != v_code.get().strip().upper():
            bad.append("課程縮寫")
        if sem != v_sem.get().strip():
            bad.append("學期")
        tip = f"學生編號會長成 {sem}_{code}_1、{sem}_{code}_2…"
        if bad:
            tip += "（" + "／".join(bad) + "格式不對，先用預設值）"
        v_codehint.set(tip)

    v_code.trace_add("write", refresh_hint)
    v_sem.trace_add("write", refresh_hint)
    refresh_hint()

    ttk.Label(info, text="學期起日格式 2026-09-06；「跑本週」要靠它自動算第幾週。",
              foreground="#64748B").grid(row=3, column=0, columnspan=4, sticky="w", padx=6)

    # ---------------- 資料夾
    fold = ttk.LabelFrame(body, text="資料夾", padding=8)
    fold.pack(fill="x", pady=(8, 0))
    fold.columnconfigure(1, weight=1)

    v_in = tk.StringVar(value=cfg.get("input_dir", ""))
    v_out = tk.StringVar(value=cfg.get("output_dir", "") or P.default_output_root())

    def browse(var, title):
        d = filedialog.askdirectory(title=title, initialdir=var.get() or os.path.expanduser("~"))
        if d:
            var.set(os.path.normpath(d))

    ttk.Label(fold, text="輸入資料夾（放遮罩後的 xlsx／CSV）").grid(row=0, column=0, sticky="e", **pad)
    ttk.Entry(fold, textvariable=v_in).grid(row=0, column=1, sticky="ew", **pad)
    ttk.Button(fold, text="瀏覽…", command=lambda: browse(v_in, "選擇輸入資料夾")
               ).grid(row=0, column=2, **pad)
    ttk.Label(fold, text="輸出資料夾").grid(row=1, column=0, sticky="e", **pad)
    ttk.Entry(fold, textvariable=v_out).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Button(fold, text="瀏覽…", command=lambda: browse(v_out, "選擇輸出資料夾")
               ).grid(row=1, column=2, **pad)

    # ---------------- 按鈕
    bar = ttk.Frame(body, padding=(0, 8))
    bar.pack(fill="x")
    row1 = ttk.Frame(bar)
    row1.pack(fill="x")
    row2 = ttk.Frame(bar)
    row2.pack(fill="x", pady=(6, 0))
    v_week = tk.StringVar(value="")
    v_thumb = tk.BooleanVar(value=False)

    buttons = []

    def add_btn(parent, text, cmd, width=14):
        b = ttk.Button(parent, text=text, command=cmd, width=width)
        b.pack(side="left", padx=(0, 6))
        buttons.append(b)
        return b

    # ---------------- 狀態列（先占住視窗底部，視窗變小也不會被擠掉）
    v_status = tk.StringVar(value="準備好了。第一次用請先按「跑示範」看看輸出長什麼樣子。")
    ttk.Label(body, textvariable=v_status, foreground="#065A82", anchor="w"
              ).pack(side="bottom", fill="x", pady=(6, 0))

    # ---------------- log
    logf = ttk.LabelFrame(body, text="執行紀錄", padding=6)
    logf.pack(fill="both", expand=True)
    log_font = ("Microsoft JhengHei UI", 10) if sys.platform.startswith("win") else ("Menlo", 11)
    txt = tk.Text(logf, wrap="none", height=16, font=log_font,
                  background="#F5F9FA", foreground="#1E293B")
    ysb = ttk.Scrollbar(logf, orient="vertical", command=txt.yview)
    xsb = ttk.Scrollbar(logf, orient="horizontal", command=txt.xview)
    txt.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set, state="disabled")
    xsb.pack(side="bottom", fill="x")
    ysb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)

    def put(line=""):
        msgq.put(("log", str(line)))

    def append(line):
        txt.configure(state="normal")
        txt.insert("end", line + "\n")
        txt.see("end")
        txt.configure(state="disabled")

    def set_busy(busy):
        for b in buttons:
            b.configure(state="disabled" if busy else "normal")

    def current_cfg():
        c = dict(cfg)
        c.update({
            "course_name": v_course.get().strip() or "我的課程",
            "teacher": v_teacher.get().strip(),
            "ta_name": v_ta.get().strip(),
            "semester_start": v_start.get().strip(),
            "course_code": P.norm_course_code(v_code.get()),
            "semester": P.norm_semester(v_sem.get()),
            "input_dir": v_in.get().strip(),
            "output_dir": v_out.get().strip(),
        })
        return c

    def remember():
        try:
            P.save_config(current_cfg())
        except Exception:
            pass

    running = {"on": False}

    def start(label, fn):
        if running["on"]:
            return
        remember()
        running["on"] = True
        set_busy(True)
        v_status.set(f"{label}… 進行中，請等它跑完（大約 10–60 秒）。")
        put("")
        put("─" * 60)
        put(f"▶ {label}　{dt.datetime.now():%Y-%m-%d %H:%M:%S}")

        def worker():
            try:
                fn(put)
                msgq.put(("done", f"{label} 完成。"))
            except P.PipelineError as e:
                msgq.put(("error", str(e)))
            except Exception:
                msgq.put(("error", "程式出現非預期錯誤：\n" + traceback.format_exc()))

        threading.Thread(target=worker, daemon=True).start()

    def pump():
        try:
            while True:
                kind, payload = msgq.get_nowait()
                if kind == "log":
                    append(payload)
                elif kind == "done":
                    running["on"] = False
                    set_busy(False)
                    v_status.set(payload)
                    append(payload)
                elif kind == "error":
                    running["on"] = False
                    set_busy(False)
                    v_status.set("沒有跑完，請看下面的訊息。")
                    append("")
                    append("[沒有跑完] " + payload)
                    messagebox.showerror(P.APP_TITLE, payload)
        except queue.Empty:
            pass
        root.after(120, pump)

    # ---------------- 動作
    def out_root():
        d = v_out.get().strip() or P.default_output_root()
        os.makedirs(d, exist_ok=True)
        return d

    def do_demo():
        start("跑示範（內建合成資料）",
              lambda log: P.run_demo(current_cfg(), out_root(),
                                     thumbs=v_thumb.get(), log=log))

    def do_week(n=None):
        if not v_in.get().strip():
            messagebox.showwarning(P.APP_TITLE, "請先選「輸入資料夾」（放 Zuvio 下載的 xlsx）。")
            return
        label = "跑本週" if n is None else f"跑第 {n} 週"
        start(label, lambda log: P.run_week(current_cfg(), v_in.get().strip(), out_root(),
                                            week_no=n, thumbs=v_thumb.get(), log=log))

    def do_week_n():
        s = v_week.get().strip()
        if not s.isdigit() or int(s) < 1:
            messagebox.showwarning(P.APP_TITLE, "請在右邊的欄位填週次數字，例如 3。")
            return
        do_week(int(s))

    def do_privacy():
        d = filedialog.askdirectory(title="選擇要檢查的輸出資料夾",
                                    initialdir=v_out.get() or os.path.expanduser("~"))
        if not d:
            return
        src = v_in.get().strip() or None

        def job(log):
            hits = P.check_privacy(d, input_dir=src, log=log)
            if hits:
                raise P.PipelineError(
                    f"隱私檢查沒有過：發現 {hits} 處未遮罩的真實姓名。\n"
                    "請先不要把這份輸出交出去，並把執行紀錄回報給開發者。")
            return hits

        start("隱私檢查", job)

    def do_open():
        try:
            P.open_folder(out_root(), log=lambda *a: put(" ".join(str(x) for x in a)))
        except P.PipelineError as e:
            messagebox.showwarning(P.APP_TITLE, str(e))

    add_btn(row1, "跑示範", do_demo, width=10)
    add_btn(row1, "跑本週", do_week, width=10)
    add_btn(row1, "跑指定週次", do_week_n, width=12)
    ttk.Entry(row1, textvariable=v_week, width=4).pack(side="left")
    ttk.Label(row1, text="週").pack(side="left", padx=(3, 12))
    add_btn(row1, "隱私檢查", do_privacy, width=11)
    add_btn(row1, "開啟輸出資料夾", do_open, width=16)
    ttk.Checkbutton(row2, text="順便用 PowerPoint 輸出簡報縮圖（沒裝 PowerPoint 會自動略過）",
                    variable=v_thumb).pack(side="left")

    append(f"{P.APP_TITLE}　版本 {P.VERSION}")
    append("兩步流程：① 先用「姓名遮罩與學生編號」處理 Zuvio 匯出的 xlsx，"
           "② 把它的輸出放進「輸入資料夾」再按「跑本週」。")
    append("步驟：1) 填課程資訊與課程縮寫／學期　2) 選輸入／輸出資料夾　3) 按「跑本週」"
           "　4) 按「隱私檢查」確認 0 命中。")
    append("隱私鐵律：輸出一律用學生編號；姓名第 2 字改 O、學號與 email 全部改 O；"
           "「學生編號連結姓名」對照表只能留在自己電腦，不要上傳。")
    append("")

    def on_close():
        remember()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.after(120, pump)
    root.mainloop()
    return 0


# ------------------------------------------------------------------ main
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--version" in argv or "-V" in argv:
        try:
            print(f"{P.APP_NAME} {P.VERSION}", flush=True)
        except Exception:
            pass
        return 0
    if "--selftest" in argv:
        return run_selftest()
    if "--check-privacy" in argv:
        return run_check_privacy(argv)
    if "--help" in argv or "-h" in argv:
        try:
            print(__doc__, flush=True)
        except Exception:
            pass
        return 0
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())
