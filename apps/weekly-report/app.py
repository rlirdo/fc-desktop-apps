# -*- coding: utf-8 -*-
"""
週報產生器（WeeklyReportMaker）
================================
東華大學自然資源與環境學系　仿生與環境工作坊　個人週報（碩士生標準版）

用法：
  雙擊執行            → 開啟圖形介面（填表 → 產生週報 .pptx）
  WeeklyReportMaker --selftest   → 不開視窗，用內建合成資料跑完整流程並自我驗收
  WeeklyReportMaker --version    → 印出版本

規格來源：YCYLabWKReport.md；驗收錨點來源：weekly-report-workshop/practice.html
本程式不含任何真實個資，內建示範資料一律虛構。
"""
import os
import sys
import json
import shutil
import tempfile
import datetime
import traceback
import subprocess

APP_NAME = "週報產生器"
EXE_NAME = "WeeklyReportMaker"
VERSION = "1.0.0"


# ------------------------------------------------------------------ console
def ensure_console():
    """PyInstaller --windowed 之下 stdout 是 None；若是從命令列啟動，接上父行程的
    主控台，讓 --selftest 仍然印得出字。失敗就算了，回傳碼才是主要訊號。"""
    if sys.platform.startswith("win"):
        # --windowed 打包之下沒有 stdout：接上呼叫端的主控台，並把主控台切成
        # UTF-8，中文才不會變成亂碼（只影響這個主控台視窗）。
        try:
            import ctypes
            ctypes.windll.kernel32.AttachConsole(-1)
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
    for name in ("stdout", "stderr"):
        s = getattr(sys, name, None)
        if s is None:
            try:
                f = open("CONOUT$" if sys.platform.startswith("win") else os.devnull,
                         "w", encoding="utf-8", errors="replace")
                setattr(sys, name, f)
            except Exception:
                pass
        else:
            try:                      # 導向檔案／管線時也統一輸出 UTF-8
                s.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


_LOG_LINES = []


def log(msg=""):
    _LOG_LINES.append(str(msg))
    try:
        print(msg, flush=True)
    except Exception:
        # 主控台編碼吃不下某些字元時，退成 ASCII 也要把訊息印出來
        try:
            sys.stdout.write(str(msg).encode("ascii", "replace").decode("ascii") + "\n")
            sys.stdout.flush()
        except Exception:
            pass


def write_log(path):
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(_LOG_LINES) + "\n")
    except Exception:
        pass
    return path


# ------------------------------------------------------------------ 路徑
def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def default_outdir():
    home = os.path.expanduser("~")
    for cand in (os.path.join(home, "Desktop"), os.path.join(home, "桌面"), home):
        if os.path.isdir(cand):
            return os.path.join(cand, f"{APP_NAME}_輸出")
    return os.path.join(home, f"{APP_NAME}_輸出")


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)          # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


# ------------------------------------------------------------------ 合成照片
def make_placeholder(path, caption, size=(1600, 1000)):
    """用 PIL 畫一張灰底格線佔位圖（合成，非真實照片）。"""
    from PIL import Image, ImageDraw, ImageFont
    from core import pick_pil_font_path
    im = Image.new("RGB", size, (0xDD, 0xE4, 0xE8))
    d = ImageDraw.Draw(im)
    for gx in range(0, size[0], 80):
        d.line([(gx, 0), (gx, size[1])], fill=(0xD2, 0xDB, 0xE0), width=1)
    for gy in range(0, size[1], 80):
        d.line([(0, gy), (size[0], gy)], fill=(0xD2, 0xDB, 0xE0), width=1)
    d.rectangle([40, 40, size[0] - 40, size[1] - 40], outline=(0x9F, 0xB2, 0xBC), width=6)
    fp = pick_pil_font_path()
    font = None
    if fp:
        try:
            font = ImageFont.truetype(fp, 46)
        except Exception:
            font = None
    try:
        tw = d.textlength(caption, font=font)
    except Exception:
        tw = len(caption) * 20
    d.text(((size[0] - tw) / 2, size[1] / 2 - 28), caption, fill=(0x64, 0x74, 0x8B), font=font)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    im.save(path, "PNG")
    return path


def load_sample_data(workdir):
    """讀內建合成資料，並把 $PLACEHOLDER:xxx 的照片實際畫出來。"""
    with open(resource_path(os.path.join("selftest", "sample_data.json")),
              "r", encoding="utf-8") as f:
        data = json.load(f)
    for s in data.get("sections", []):
        for p in s.get("photos", []):
            v = str(p.get("path", ""))
            if v.startswith("$PLACEHOLDER:"):
                name = v.split(":", 1)[1]
                out = os.path.join(workdir, "assets", name)
                make_placeholder(out, "SYNTHETIC PLACEHOLDER  /  " + os.path.splitext(name)[0].upper())
                p["path"] = out
    return data


# ------------------------------------------------------------------ selftest
def run_selftest(workdir=None):
    ensure_console()
    workdir = workdir or os.path.join(tempfile.gettempdir(), f"{EXE_NAME}_selftest")
    logfile = os.path.join(tempfile.gettempdir(), f"{EXE_NAME}_selftest.log")
    ok = False
    try:
        if os.path.isdir(workdir):
            shutil.rmtree(workdir, ignore_errors=True)
        os.makedirs(workdir, exist_ok=True)
        log(f"{EXE_NAME} {VERSION} selftest")
        log(f"工作資料夾：{workdir}")

        import report_builder as rb
        import validator as va

        data = load_sample_data(workdir)
        out = os.path.join(workdir, rb.default_filename(data))
        path, pages = rb.build_report(data, out)
        log(f"產出：{path}（{pages} 頁）")
        if not os.path.exists(path) or os.path.getsize(path) < 20000:
            raise RuntimeError("產出的 pptx 不存在或太小")

        checks, summary = va.validate(path)
        log(va.format_report(checks, summary))

        fails = [c["name"] for c in checks if not c["ok"]]
        if fails:
            raise RuntimeError("驗收未過：" + "、".join(fails))
        if summary["passed"] != 6:
            raise RuntimeError(f"六項必檢只過 {summary['passed']} 項")
        if not (11 <= summary["slides"] <= 15):
            raise RuntimeError(f"頁數 {summary['slides']} 不在 11–15 之間")

        # 額外：確認每一頁備忘稿都是「一句一行」而且有內容
        from pptx import Presentation
        prs = Presentation(path)
        for i, sl in enumerate(prs.slides, start=1):
            txt = sl.notes_slide.notes_text_frame.text.strip()
            if len(txt) < 10:
                raise RuntimeError(f"第 {i} 頁逐字稿太短")
        log("逐字稿抽查：每一頁都有內容。")
        ok = True
    except Exception as e:
        log("SELFTEST FAILED: " + repr(e))
        log(traceback.format_exc())
    log("SELFTEST OK" if ok else "SELFTEST FAILED")
    write_log(logfile)
    try:
        print(f"（記錄檔：{logfile}）", flush=True)
    except Exception:
        pass
    return 0 if ok else 1


# ------------------------------------------------------------------ GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    import report_builder as rb
    import validator as va

    CN = rb.CN_NUM
    PAD = 6

    class SectionTab(ttk.Frame):
        def __init__(self, master, idx):
            super().__init__(master, padding=PAD)
            self.idx = idx
            self.name = tk.StringVar(value=(rb.DEFAULT_SECTIONS[idx]
                                            if idx < len(rb.DEFAULT_SECTIONS) else ""))
            top = ttk.Frame(self)
            top.pack(fill="x")
            ttk.Label(top, text=f"{CN[idx]}、區塊名稱：").pack(side="left")
            ttk.Entry(top, textvariable=self.name, width=24).pack(side="left")
            ttk.Label(top, text="（序號是規格，不可改；名稱可以改）",
                      foreground="#64748B").pack(side="left", padx=6)

            ttk.Label(self, text="本週進度（一行一件事；留空白會自動填「本週無進度」）").pack(
                anchor="w", pady=(8, 2))
            self.progress = tk.Text(self, height=8, wrap="word", undo=True)
            self.progress.pack(fill="both", expand=True)

            ttk.Label(self, text="照片（最多 2 張，請填拍攝日期與地點）").pack(
                anchor="w", pady=(8, 2))
            self.photos = []
            for i in range(2):
                self.photos.append(self._photo_row(i))

        def _photo_row(self, i):
            fr = ttk.LabelFrame(self, text=f"照片 {i+1}", padding=4)
            fr.pack(fill="x", pady=2)
            v = {"path": tk.StringVar(), "caption": tk.StringVar(),
                 "date": tk.StringVar(), "place": tk.StringVar()}
            r1 = ttk.Frame(fr)
            r1.pack(fill="x")
            ttk.Entry(r1, textvariable=v["path"]).pack(side="left", fill="x", expand=True)
            ttk.Button(r1, text="選擇…", width=8,
                       command=lambda: self._pick(v["path"])).pack(side="left", padx=3)
            ttk.Button(r1, text="清除", width=6,
                       command=lambda: v["path"].set("")).pack(side="left")
            r2 = ttk.Frame(fr)
            r2.pack(fill="x", pady=2)
            ttk.Label(r2, text="圖說").pack(side="left")
            ttk.Entry(r2, textvariable=v["caption"]).pack(side="left", fill="x",
                                                          expand=True, padx=3)
            ttk.Label(r2, text="日期").pack(side="left")
            ttk.Entry(r2, textvariable=v["date"], width=12).pack(side="left", padx=3)
            ttk.Label(r2, text="地點").pack(side="left")
            ttk.Entry(r2, textvariable=v["place"], width=18).pack(side="left", padx=3)
            return v

        def _pick(self, var):
            p = filedialog.askopenfilename(
                title="選擇照片",
                filetypes=[("圖片", "*.jpg *.jpeg *.png *.bmp *.gif"), ("全部檔案", "*.*")])
            if p:
                var.set(p)

        def get(self):
            return {"name": self.name.get().strip(),
                    "progress": self.progress.get("1.0", "end").strip(),
                    "photos": [{"path": v["path"].get().strip(),
                                "caption": v["caption"].get().strip(),
                                "date": v["date"].get().strip(),
                                "place": v["place"].get().strip()}
                               for v in self.photos if v["path"].get().strip()]}

        def set(self, s):
            self.name.set(s.get("name", ""))
            self.progress.delete("1.0", "end")
            self.progress.insert("1.0", s.get("progress", ""))
            ph = s.get("photos", []) or []
            for i, v in enumerate(self.photos):
                d = ph[i] if i < len(ph) else {}
                for k in ("path", "caption", "date", "place"):
                    v[k].set(d.get(k, ""))

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(f"{APP_NAME}　{EXE_NAME} {VERSION}")
            self.geometry("1020x760")
            self.minsize(880, 640)
            try:
                ico = resource_path("icon.ico")
                if sys.platform.startswith("win") and os.path.exists(ico):
                    self.iconbitmap(ico)
            except Exception:
                pass
            self._build()

        # ---- 版面 ----
        def _build(self):
            head = ttk.Frame(self, padding=(10, 8))
            head.pack(fill="x")
            self.reporter = tk.StringVar()
            self.grade = tk.StringVar(value="碩士班一年級")
            self.advisor = tk.StringVar()
            self.rdate = tk.StringVar(value=datetime.date.today().strftime("%Y/%m/%d"))
            self.wstart = tk.StringVar()
            self.wend = tk.StringVar()
            self.outdir = tk.StringVar(value=default_outdir())
            self.nsec = tk.IntVar(value=5)

            r = ttk.Frame(head); r.pack(fill="x")
            for lab, var, w in (("報告人", self.reporter, 14), ("年級", self.grade, 14),
                                ("指導教授", self.advisor, 14),
                                ("報告日期 (YYYY/MM/DD)", self.rdate, 14)):
                ttk.Label(r, text=lab + "：").pack(side="left")
                ttk.Entry(r, textvariable=var, width=w).pack(side="left", padx=(0, 10))
            ttk.Button(r, text="依報告日期帶入本週期間", command=self._fill_week).pack(side="left")

            r2 = ttk.Frame(head); r2.pack(fill="x", pady=4)
            ttk.Label(r2, text="本週期間：").pack(side="left")
            ttk.Entry(r2, textvariable=self.wstart, width=14).pack(side="left")
            ttk.Label(r2, text="（日） －").pack(side="left")
            ttk.Entry(r2, textvariable=self.wend, width=14).pack(side="left")
            ttk.Label(r2, text="（六）　區塊數：").pack(side="left")
            ttk.Spinbox(r2, from_=3, to=6, width=4, textvariable=self.nsec,
                        command=self._sync_tabs).pack(side="left")
            ttk.Label(r2, text="　輸出資料夾：").pack(side="left")
            ttk.Entry(r2, textvariable=self.outdir, width=42).pack(side="left", fill="x",
                                                                   expand=True)
            ttk.Button(r2, text="…", width=3, command=self._pick_out).pack(side="left", padx=3)
            self._fill_week()

            self.nb = ttk.Notebook(self)
            self.nb.pack(fill="both", expand=True, padx=10, pady=4)
            self.tabs = []
            for i in range(6):
                t = SectionTab(self.nb, i)
                self.tabs.append(t)

            self.qa = ttk.Frame(self.nb, padding=PAD)
            ttk.Label(self.qa, text="問題與請益（一行一題；留空白會自動填「"
                      + rb.NO_QUESTION + "」）").pack(anchor="w")
            self.questions = tk.Text(self.qa, height=8, wrap="word", undo=True)
            self.questions.pack(fill="both", expand=True, pady=(2, 8))
            ttk.Label(self.qa, text="下週計畫（一行一件事；句子裡有「硬性／截止／期限／小考」"
                      "會自動標成硬性時程）").pack(anchor="w")
            self.nextweek = tk.Text(self.qa, height=8, wrap="word", undo=True)
            self.nextweek.pack(fill="both", expand=True, pady=2)

            self._sync_tabs()

            bar = ttk.Frame(self, padding=(10, 4))
            bar.pack(fill="x")
            for lab, cmd in (("載入內容…", self.on_load), ("儲存內容…", self.on_save),
                             ("產生週報", self.on_build), ("驗收 .pptx…", self.on_validate),
                             ("開啟輸出資料夾", self.on_open), ("填入示範內容", self.on_demo)):
                ttk.Button(bar, text=lab, command=cmd).pack(side="left", padx=3)
            ttk.Button(bar, text="離開", command=self.destroy).pack(side="right")

            self.logbox = scrolledtext.ScrolledText(self, height=7, wrap="word")
            self.logbox.pack(fill="x", padx=10, pady=(0, 8))
            self.say(f"{APP_NAME} {VERSION}　準備就緒。填完表單按「產生週報」。")
            self.say("提醒：沒有進度的區塊留空白即可，程式會自動在總覽頁寫「本週無進度」。")

        # ---- 小工具 ----
        def say(self, msg):
            self.logbox.insert("end", str(msg) + "\n")
            self.logbox.see("end")
            self.update_idletasks()

        def _fill_week(self):
            a, b = rb.week_range(self.rdate.get())
            self.wstart.set(a.strftime("%Y/%m/%d"))
            self.wend.set(b.strftime("%Y/%m/%d"))

        def _pick_out(self):
            p = filedialog.askdirectory(title="選擇輸出資料夾")
            if p:
                self.outdir.set(p)

        def _sync_tabs(self):
            n = max(3, min(6, int(self.nsec.get() or 5)))
            for t in self.tabs:
                try:
                    self.nb.forget(t)
                except Exception:
                    pass
            try:
                self.nb.forget(self.qa)
            except Exception:
                pass
            for i in range(n):
                self.nb.add(self.tabs[i], text=f"{CN[i]}、{self.tabs[i].name.get() or '區塊'}")
            self.nb.add(self.qa, text="問題與請益／下週計畫")

        def collect(self):
            n = max(3, min(6, int(self.nsec.get() or 5)))
            return {"schema": "weekly-report/1",
                    "reporter": self.reporter.get().strip(),
                    "grade": self.grade.get().strip(),
                    "advisor": self.advisor.get().strip(),
                    "unit": "國立東華大學自然資源與環境學系　仿生與環境工作坊",
                    "report_date": self.rdate.get().strip(),
                    "week_start": self.wstart.get().strip(),
                    "week_end": self.wend.get().strip(),
                    "sections": [self.tabs[i].get() for i in range(n)],
                    "questions": self.questions.get("1.0", "end").strip(),
                    "next_week": self.nextweek.get("1.0", "end").strip()}

        def apply(self, d):
            self.reporter.set(d.get("reporter", ""))
            self.grade.set(d.get("grade", "碩士班一年級"))
            self.advisor.set(d.get("advisor", ""))
            self.rdate.set(d.get("report_date", self.rdate.get()))
            self.wstart.set(d.get("week_start", ""))
            self.wend.set(d.get("week_end", ""))
            if not self.wstart.get() or not self.wend.get():
                self._fill_week()
            secs = d.get("sections", []) or []
            self.nsec.set(max(3, min(6, len(secs) or 5)))
            for i, t in enumerate(self.tabs):
                t.set(secs[i] if i < len(secs) else {"name": "", "progress": "", "photos": []})
            self.questions.delete("1.0", "end")
            self.questions.insert("1.0", d.get("questions", ""))
            self.nextweek.delete("1.0", "end")
            self.nextweek.insert("1.0", d.get("next_week", ""))
            self._sync_tabs()

        # ---- 按鈕 ----
        def on_demo(self):
            if not messagebox.askyesno(APP_NAME, "要用內建的示範內容（虛構人物「林小仿」）"
                                                 "覆蓋目前表單嗎？"):
                return
            work = os.path.join(tempfile.gettempdir(), f"{EXE_NAME}_demo")
            try:
                d = load_sample_data(work)
            except Exception as e:
                messagebox.showerror(APP_NAME, f"讀取示範內容失敗：{e}")
                return
            self.apply(d)
            self.say("已填入示範內容（全部虛構）。可直接按「產生週報」看看成品。")

        def on_save(self):
            p = filedialog.asksaveasfilename(title="儲存內容", defaultextension=".json",
                                             initialfile="週報內容.json",
                                             filetypes=[("JSON", "*.json")])
            if not p:
                return
            with open(p, "w", encoding="utf-8") as f:
                json.dump(self.collect(), f, ensure_ascii=False, indent=2)
            self.say(f"已儲存內容：{p}")

        def on_load(self):
            p = filedialog.askopenfilename(title="載入內容",
                                           filetypes=[("JSON", "*.json"), ("全部檔案", "*.*")])
            if not p:
                return
            try:
                with open(p, "r", encoding="utf-8") as f:
                    self.apply(json.load(f))
                self.say(f"已載入內容：{p}")
            except Exception as e:
                messagebox.showerror(APP_NAME, f"載入失敗：{e}")

        def on_build(self):
            d = self.collect()
            if not d["reporter"]:
                messagebox.showwarning(APP_NAME, "請先填「報告人」。")
                return
            outdir = self.outdir.get().strip() or default_outdir()
            try:
                os.makedirs(outdir, exist_ok=True)
                out = os.path.join(outdir, rb.default_filename(d))
                self.say("開始產生週報…")
                path, pages = rb.build_report(d, out)
                self.say(f"完成：{path}（{pages} 頁）")
                checks, summary = va.validate(path)
                self.say(f"自動驗收：六項必檢 {summary['passed']} / {summary['total']} 通過，"
                         f"共 {summary['slides']} 頁。")
                for c in checks:
                    if not c["ok"]:
                        self.say(f"　未過：{c['name']} — {c['detail']}")
                self._show_result(checks, summary)
            except Exception as e:
                self.say("產生失敗：" + repr(e))
                messagebox.showerror(APP_NAME, f"產生失敗：\n{e}")

        def on_validate(self):
            p = filedialog.askopenfilename(title="選擇要驗收的週報",
                                           filetypes=[("PowerPoint", "*.pptx")])
            if not p:
                return
            try:
                checks, summary = va.validate(p)
                self._show_result(checks, summary)
            except Exception as e:
                messagebox.showerror(APP_NAME, f"驗收失敗：{e}")

        def on_open(self):
            d = self.outdir.get().strip() or default_outdir()
            os.makedirs(d, exist_ok=True)
            if not open_folder(d):
                messagebox.showinfo(APP_NAME, f"請自行開啟：{d}")

        def _show_result(self, checks, summary):
            win = tk.Toplevel(self)
            win.title("驗收結果")
            win.geometry("820x420")
            ttk.Label(win, padding=8,
                      text=f"{os.path.basename(summary['path'])}　｜　{summary['slides']} 頁"
                           f"　｜　六項必檢 {summary['passed']} / {summary['total']} 通過"
                      ).pack(anchor="w")
            cols = ("項目", "結果", "說明")
            tv = ttk.Treeview(win, columns=cols, show="headings", height=10)
            for c, w in zip(cols, (220, 70, 500)):
                tv.heading(c, text=c)
                tv.column(c, width=w, anchor="w")
            for c in checks:
                tv.insert("", "end", values=(c["name"], "通過" if c["ok"] else "未過",
                                             c["detail"]))
            tv.pack(fill="both", expand=True, padx=8, pady=4)
            ttk.Button(win, text="關閉", command=win.destroy).pack(pady=6)

    App().mainloop()
    return 0


# ------------------------------------------------------------------ main
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--version" in argv or "-V" in argv:
        ensure_console()
        try:
            print(f"{EXE_NAME} {VERSION}（{APP_NAME}）", flush=True)
        except Exception:
            pass
        return 0
    if "--selftest" in argv:
        return run_selftest()
    if "--help" in argv or "-h" in argv:
        ensure_console()
        print(__doc__)
        return 0
    try:
        return run_gui()
    except Exception:
        ensure_console()
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
