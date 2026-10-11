# -*- coding: utf-8 -*-
"""進度視窗（可取消）＋背景執行緒工作器。

run_task(root, title, job, on_done) 會：
  1. 開 ProgressWindow；
  2. 在 daemon 執行緒跑 job(progress, cancel_event)；
  3. progress(msg, frac) 透過 queue 回到主執行緒更新畫面（GUI 不凍結）；
  4. 完成 → on_done(result)；取消 → on_cancel()；例外 → FcDialog 錯誤視窗。
"""
import gc
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk

from . import dialogs
from . import theme as T
from .credit import CreditToplevel


class ProgressWindow(CreditToplevel):
    def __init__(self, parent, title, on_cancel=None):
        super().__init__(parent, title, size=(540, 280), modal=False)
        self._on_cancel = on_cancel
        self.var_msg = tk.StringVar(value="準備中…")
        tk.Label(self.body, text=title, font=T.FT, fg=T.NAVY, bg=T.BG, anchor="w").pack(fill="x")
        tk.Label(self.body, textvariable=self.var_msg, font=T.F, fg=T.INK, bg=T.BG, anchor="w",
                 wraplength=500, justify="left").pack(fill="x", pady=(6, 4))
        self.bar = ttk.Progressbar(self.body, mode="indeterminate", maximum=1000)
        self.bar.pack(fill="x", pady=4)
        self.bar.start(12)
        self._indet = True
        self.log = tk.Text(self.body, height=4, font=T.FS, bg="white", fg=T.GREY, wrap="none")
        self.log.pack(fill="both", expand=True, pady=(4, 4))
        bar = tk.Frame(self.body, bg=T.BG)
        bar.pack(fill="x")
        self.btn_cancel = ttk.Button(bar, text="取消", width=10, command=self.cancel)
        self.btn_cancel.pack(side="right")
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.show()

    def update_progress(self, msg, frac):
        if msg:
            self.var_msg.set(msg)
            self.log.insert("end", msg + "\n")
            self.log.see("end")
        if frac is None:
            if not self._indet:
                self.bar.config(mode="indeterminate")
                self.bar.start(12)
                self._indet = True
        else:
            if self._indet:
                self.bar.stop()
                self.bar.config(mode="determinate")
                self._indet = False
            self.bar["value"] = max(0.0, min(1.0, float(frac))) * 1000

    def cancel(self):
        self.var_msg.set("正在取消，等待目前的網路請求結束…")
        self.btn_cancel.config(state="disabled")
        if self._on_cancel:
            self._on_cancel()


class Task:
    """一個背景工作；cancel_event 可由「停止」按鈕或進度視窗的「取消」觸發。"""

    def __init__(self, root, title, job, on_done=None, on_cancel=None, on_error=None,
                 show_window=True):
        self.root = root
        self.title = title
        self.job = job
        self.on_done = on_done
        self.on_cancel = on_cancel
        self.on_error = on_error
        self.cancel_event = threading.Event()
        self.q = queue.Queue()
        self.win = ProgressWindow(root, title, on_cancel=self.cancel_event.set) if show_window else None
        self.finished = False
        gc.collect()   # 先在主執行緒回收含 Tk 變數的循環垃圾，避免在背景執行緒觸發 Variable.__del__
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        self.root.after(80, self._poll)

    def progress(self, msg, frac=None):
        self.q.put(("p", msg, frac))

    def cancel(self):
        self.cancel_event.set()
        if self.win is not None:
            try:
                self.win.cancel()
            except tk.TclError:
                pass

    def _run(self):
        try:
            res = self.job(self.progress, self.cancel_event)
            self.q.put(("done", res, None))
        except Exception as e:  # noqa: BLE001
            if type(e).__name__ == "Cancelled" or self.cancel_event.is_set():
                self.q.put(("cancel", None, None))
            else:
                self.q.put(("err", e, traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                kind, a, b = self.q.get_nowait()
                if kind == "p":
                    if self.win is not None:
                        try:
                            self.win.update_progress(a, b)
                        except tk.TclError:
                            pass
                    continue
                self._finish(kind, a, b)
                return
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _finish(self, kind, a, b):
        self.finished = True
        gc.collect()
        if self.win is not None:
            try:
                self.win.destroy()
            except tk.TclError:
                pass
        if kind == "done":
            if self.on_done:
                self.on_done(a)
        elif kind == "cancel":
            if self.on_cancel:
                self.on_cancel()
            else:
                dialogs.info(self.root, "已取消", f"「{self.title}」已取消；已取得的結果保留。")
        else:
            if self.on_error:
                self.on_error(a, b)
            else:
                dialogs.error(self.root, "發生錯誤", f"「{self.title}」失敗：{a}", detail=b)


def run_task(root, title, job, on_done=None, **kw):
    return Task(root, title, job, on_done=on_done, **kw)
