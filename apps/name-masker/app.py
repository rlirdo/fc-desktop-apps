#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
姓名遮罩 NameMasker — GUI 入口
=================================
把 Excel 名單「姓名」欄的中文姓名第 2 個字改成 O（王小明 → 王O明），
另存成「原檔名02.xlsx」，原檔完全不動。全程離線，資料不離開你的電腦。

用法：
  雙擊 NameMasker.exe                → 開啟視窗，選檔案後按「開始遮罩」
  把 Excel 檔拖到 NameMasker.exe 上  → 直接處理，不開視窗（結果用訊息框顯示）
  NameMasker.exe 檔案.xlsx           → 同上（命令列模式）
  NameMasker.exe --selftest          → 自我測試（成功印 SELFTEST OK 並 exit 0）
  NameMasker.exe --version           → 顯示版本

環境變數 NAMEMASKER_NO_DIALOG=1：命令列模式不跳訊息框（供自動化／CI 使用）。
"""
import os
import sys
import tempfile
import traceback

APP_NAME = "姓名遮罩"
EXE_NAME = "NameMasker"
VERSION = "1.0.0"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mask_names import MaskError, mask_file, mask_name, SUPPORTED_EXT  # noqa: E402

_LOG_LINES = []


def resource_path(rel):
    """取得打包後（sys._MEIPASS）或原始碼目錄下的資料檔路徑。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def out(msg=""):
    """console=False 封裝時 stdout 可能不存在；一律記錄到 log，有 console 就順便印。"""
    _LOG_LINES.append(str(msg))
    try:
        if sys.stdout is not None:
            print(msg)
            sys.stdout.flush()
    except Exception:
        pass


def write_log(name):
    path = os.path.join(tempfile.gettempdir(), name)
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(_LOG_LINES) + "\n")
    except Exception:
        return None
    return path


# --------------------------------------------------------------------------
# 自我測試
# --------------------------------------------------------------------------
def run_selftest():
    import shutil
    import hashlib
    import openpyxl
    import sample_data

    out(f"{EXE_NAME} {VERSION} — 自我測試開始（使用合成資料，零真實個資）")
    tmp = tempfile.mkdtemp(prefix="NameMasker_selftest_")
    try:
        # --- 0. 遮罩規則單元檢查 -----------------------------------------
        unit_cases = [
            ("王小明", "王O明", True),
            ("李明", "李O", True),
            ("歐陽小花", "歐O小花", True),
            ("匿名作答者", "匿O作答者", True),
            ("Mary Chen", "Mary Chen", False),
            ("", "", False),
            ("A王小明B", "A王O明B", True),
        ]
        for src_text, want, want_flag in unit_cases:
            got, flag = mask_name(src_text)
            assert got == want and flag == want_flag, \
                f"mask_name('{src_text}') = ({got!r}, {flag}) 期望 ({want!r}, {want_flag})"
        out(f"[1/6] 遮罩規則單元檢查通過（{len(unit_cases)} 條）")

        # --- 1. 產生合成檔 -----------------------------------------------
        src = os.path.join(tmp, "合成名單.xlsx")
        sample_data.write_sample(src)
        before_hash = hashlib.sha256(open(src, "rb").read()).hexdigest()
        out(f"[2/6] 已產生合成測試檔：{src}")

        # --- 2. 執行遮罩 ---------------------------------------------------
        r = mask_file(src)
        out(f"[3/6] 遮罩完成：{r['masked']} 筆 → 工作表「{r['sheet']}」，"
            f"姓名欄 {r['column']}（標題「{r['header']}」）")

        assert os.path.basename(r["dst"]) == "合成名單02.xlsx", \
            f"輸出檔名應為 合成名單02.xlsx，實際為 {os.path.basename(r['dst'])}"
        assert os.path.dirname(r["dst"]) == tmp, "輸出檔應與原檔同資料夾"
        assert r["header"] == "姓名" and r["column"] == "C", \
            f"應精確命中「姓名」欄（C），實際 {r['column']}／{r['header']}"
        assert r["masked"] == sample_data.EXPECTED_MASKED_COUNT, \
            f"遮罩筆數應為 {sample_data.EXPECTED_MASKED_COUNT}，實際 {r['masked']}"

        # --- 3. 原檔不可被更動 --------------------------------------------
        after_hash = hashlib.sha256(open(src, "rb").read()).hexdigest()
        assert before_hash == after_hash, "原檔被更動了（不允許）"
        out("[4/6] 原檔位元組完全未更動")

        # --- 4. 讀回輸出檔逐筆驗證 ----------------------------------------
        wb = openpyxl.load_workbook(r["dst"])
        expect_sheet = sample_data.SHEET_TITLE + "_遮罩"
        assert expect_sheet in wb.sheetnames, \
            f"找不到新工作表「{expect_sheet}」，現有：{wb.sheetnames}"
        assert sample_data.SHEET_TITLE in wb.sheetnames, "原工作表應原樣保留"

        orig = wb[sample_data.SHEET_TITLE]
        masked_ws = wb[expect_sheet]

        # 原工作表在新檔中仍為未遮罩原值
        for i, row in enumerate(sample_data.ROWS):
            got = orig.cell(i + 2, sample_data.NAME_COL).value
            assert got == row[2], f"原工作表第 {i + 2} 列姓名應為 {row[2]!r}，實際 {got!r}"

        # 遮罩工作表標題列
        for c, h in enumerate(sample_data.HEADERS, start=1):
            got = masked_ws.cell(1, c).value
            assert got == h, f"遮罩表標題第 {c} 欄應為 {h!r}，實際 {got!r}"

        # 遮罩工作表逐筆姓名
        for i, want in enumerate(sample_data.EXPECTED_NAMES):
            got = masked_ws.cell(i + 2, sample_data.NAME_COL).value
            assert got == want, f"遮罩表第 {i + 2} 列姓名應為 {want!r}，實際 {got!r}"

        # 非姓名欄不得被遮罩
        for i, row in enumerate(sample_data.ROWS):
            got = masked_ws.cell(i + 2, sample_data.NOTE_COL).value
            assert got == row[1], f"非姓名欄第 {i + 2} 列應維持 {row[1]!r}，實際 {got!r}"
            got_score = masked_ws.cell(i + 2, 4).value
            assert got_score == row[3], f"分數欄第 {i + 2} 列應維持 {row[3]!r}，實際 {got_score!r}"
        wb.close()
        out(f"[5/6] 輸出檔逐筆驗證通過（{len(sample_data.EXPECTED_NAMES)} 列，含中文/英文/空白/數字）")

        # --- 5. 再跑一次應產生 03，且缺姓名欄要給清楚錯誤 -----------------
        r2 = mask_file(src)
        assert os.path.basename(r2["dst"]) == "合成名單03.xlsx", \
            f"第二次輸出應為 合成名單03.xlsx，實際 {os.path.basename(r2['dst'])}"

        bad = os.path.join(tmp, "沒有姓名欄.xlsx")
        wb2 = openpyxl.Workbook()
        wb2.active.append(["座號", "分數"])
        wb2.active.append([1, 100])
        wb2.save(bad)
        wb2.close()
        try:
            mask_file(bad)
            raise AssertionError("缺「姓名」欄時應丟出 MaskError")
        except MaskError as e:
            assert "姓名" in str(e)
        out("[6/6] 重複輸出編號（02→03）與錯誤處理檢查通過")

        out("SELFTEST OK")
        return 0
    except Exception:
        out("SELFTEST FAILED")
        out(traceback.format_exc())
        return 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------------------
# 命令列（拖放到 exe 圖示）模式
# --------------------------------------------------------------------------
def run_cli(paths):
    results, errors = [], []
    for p in paths:
        try:
            r = mask_file(p)
            results.append(r)
            out(f"完成：{os.path.basename(r['src'])} → 遮罩 {r['masked']} 筆，"
                f"已另存 {r['dst']}")
        except MaskError as e:
            errors.append((p, str(e)))
            out(f"失敗：{os.path.basename(p)}\n{e}")
        except Exception as e:  # noqa: BLE001
            errors.append((p, f"未預期的錯誤：{e}"))
            out(f"失敗：{os.path.basename(p)}\n{e}")

    lines = []
    for r in results:
        lines.append(f"✔ {os.path.basename(r['src'])}\n"
                     f"    遮罩 {r['masked']} 筆 → {os.path.basename(r['dst'])}")
    for p, msg in errors:
        lines.append(f"✘ {os.path.basename(p)}\n    {msg}")
    lines.append("")
    lines.append("原檔案完全未更動；全程離線，資料未離開你的電腦。")
    body = "\n".join(lines)

    if os.environ.get("NAMEMASKER_NO_DIALOG") == "1":
        return 1 if errors else 0

    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        if errors:
            messagebox.showwarning(f"{APP_NAME}　處理結果", body)
        else:
            messagebox.showinfo(f"{APP_NAME}　處理結果", body)
        root.destroy()
    except Exception:
        pass
    return 1 if errors else 0


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def pick_font():
    import tkinter.font as tkfont
    try:
        fams = set(tkfont.families())
    except Exception:
        return "TkDefaultFont"
    for f in ("Microsoft JhengHei UI", "微軟正黑體", "Microsoft JhengHei",
              "PingFang TC", "Heiti TC", "Noto Sans CJK TC"):
        if f in fams:
            return f
    return "TkDefaultFont"


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"{APP_NAME}　{EXE_NAME} v{VERSION}")
    root.geometry("760x560")
    root.minsize(660, 480)
    try:
        ico = resource_path("icon.ico")
        if os.path.exists(ico):
            root.iconbitmap(ico)
    except Exception:
        pass

    fam = pick_font()
    F = (fam, 11)
    FB = (fam, 11, "bold")
    FT = (fam, 16, "bold")
    FS = (fam, 9)

    files = []
    last_out_dir = {"path": None}

    pad = {"padx": 14, "pady": 4}

    tk.Label(root, text="姓名遮罩", font=FT, anchor="w").pack(fill="x", padx=14, pady=(14, 0))
    tk.Label(root,
             text="把 Excel 第 1 個工作表「姓名」欄的中文姓名第 2 個字改成 O（王小明 → 王O明），"
                  "結果放在新工作表，另存成「原檔名02.xlsx」。原檔不會被更動。",
             font=F, anchor="w", justify="left", wraplength=720, fg="#333333").pack(fill="x", **pad)
    tk.Label(root, text="🔒 完全離線：只讀寫你電腦上的檔案，不連網、不上傳任何資料。",
             font=FB, anchor="w", fg="#1b6b3a").pack(fill="x", padx=14, pady=(0, 6))

    # ---- 檔案清單 ----
    box = tk.LabelFrame(root, text=" 待處理的 Excel 檔案 ", font=FB, fg="#204a87")
    box.pack(fill="both", expand=False, padx=14, pady=6)
    lb = tk.Listbox(box, font=F, height=6, activestyle="none")
    sb = tk.Scrollbar(box, command=lb.yview)
    lb.config(yscrollcommand=sb.set)
    lb.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
    sb.pack(side="right", fill="y", padx=(0, 8), pady=8)

    def refresh():
        lb.delete(0, "end")
        for p in files:
            lb.insert("end", p)
        btn_run.config(state=("normal" if files else "disabled"))

    def add_files():
        picked = filedialog.askopenfilenames(
            title="選擇 Excel 檔案（可按住 Ctrl 多選）",
            filetypes=[("Excel 活頁簿", "*.xlsx *.xlsm"), ("所有檔案", "*.*")])
        for p in picked:
            if p not in files:
                files.append(p)
        refresh()

    def clear_files():
        files.clear()
        refresh()

    def log(msg=""):
        txt.config(state="normal")
        txt.insert("end", msg + "\n")
        txt.see("end")
        txt.config(state="disabled")
        root.update_idletasks()

    def do_run():
        if not files:
            return
        btn_run.config(state="disabled")
        txt.config(state="normal")
        txt.delete("1.0", "end")
        txt.config(state="disabled")
        ok = fail = 0
        for p in list(files):
            try:
                r = mask_file(p)
                ok += 1
                last_out_dir["path"] = os.path.dirname(r["dst"])
                log(f"✔ {os.path.basename(r['src'])}")
                log(f"    姓名欄：{r['column']} 欄（標題「{r['header']}」）")
                log(f"    遮罩 {r['masked']} 筆 → 新工作表「{r['sheet']}」")
                log(f"    已另存：{r['dst']}")
            except MaskError as e:
                fail += 1
                log(f"✘ {os.path.basename(p)}")
                for line in str(e).splitlines():
                    log("    " + line)
            except Exception as e:  # noqa: BLE001
                fail += 1
                log(f"✘ {os.path.basename(p)}　未預期的錯誤：{e}")
            log("")
        log(f"── 全部完成：成功 {ok} 個、失敗 {fail} 個。原檔案完全未更動。")
        btn_run.config(state="normal")
        btn_open.config(state=("normal" if last_out_dir["path"] else "disabled"))
        if fail == 0:
            messagebox.showinfo(APP_NAME, f"已完成 {ok} 個檔案。\n輸出檔就放在原檔案的同一個資料夾裡。")
        else:
            messagebox.showwarning(APP_NAME, f"完成 {ok} 個，{fail} 個未能處理。\n請看下方訊息區的說明。")

    def open_folder():
        d = last_out_dir["path"]
        if not d or not os.path.isdir(d):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(d)  # noqa: S606
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", d])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", d])
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"無法開啟資料夾：{e}")

    def about():
        messagebox.showinfo(
            f"關於 {APP_NAME}",
            f"{APP_NAME}（{EXE_NAME}）v{VERSION}\n\n"
            "遮罩規則：第 1 個工作表、第 1 列找「姓名」欄，\n"
            "第 2 列起把中文姓名第 2 個字改成 O。\n"
            "結果放新工作表「原表名_遮罩」，另存「原檔名02.xlsx」。\n\n"
            "隱私：完全離線，不連網、不上傳任何資料。")

    bar = tk.Frame(root)
    bar.pack(fill="x", padx=14, pady=(2, 6))
    tk.Button(bar, text="選擇 Excel 檔（可多選）", font=FB, command=add_files,
              width=20).pack(side="left")
    tk.Button(bar, text="清除清單", font=F, command=clear_files, width=10).pack(side="left", padx=6)
    btn_run = tk.Button(bar, text="開始遮罩", font=FB, command=do_run, width=12,
                        state="disabled", bg="#204a87", fg="white",
                        activebackground="#16345f", activeforeground="white")
    btn_run.pack(side="left", padx=6)
    btn_open = tk.Button(bar, text="開啟輸出資料夾", font=F, command=open_folder,
                         width=14, state="disabled")
    btn_open.pack(side="left", padx=6)
    tk.Button(bar, text="關於", font=F, command=about, width=6).pack(side="right")

    outbox = tk.LabelFrame(root, text=" 處理結果 ", font=FB, fg="#204a87")
    outbox.pack(fill="both", expand=True, padx=14, pady=(0, 6))
    txt = tk.Text(outbox, font=(fam, 10), wrap="word", state="disabled",
                  bg="#fbfbfb", relief="flat")
    sb2 = ttk.Scrollbar(outbox, command=txt.yview)
    txt.config(yscrollcommand=sb2.set)
    txt.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
    sb2.pack(side="right", fill="y", padx=(0, 8), pady=8)

    tk.Label(root, text=f"{EXE_NAME} v{VERSION}　｜　也可以把 Excel 檔直接拖到本程式的圖示上執行"
                        "　｜　東華大學自然資源與環境學系　仿生與環境工作坊",
             font=FS, fg="#777777", anchor="w").pack(fill="x", padx=14, pady=(0, 10))

    # 啟動即帶入命令列給的檔案（若有）
    for p in sys.argv[1:]:
        if p.lower().endswith(SUPPORTED_EXT) and os.path.isfile(p) and p not in files:
            files.append(p)
    refresh()

    root.mainloop()
    return 0


# --------------------------------------------------------------------------
def main():
    args = sys.argv[1:]
    if "--version" in args or "-v" in args:
        out(f"{EXE_NAME} {APP_NAME} {VERSION}")
        write_log(f"{EXE_NAME}_version.log")
        return 0
    if "--selftest" in args:
        code = run_selftest()
        p = write_log(f"{EXE_NAME}_selftest.log")
        if p:
            out(f"（測試紀錄：{p}）")
            write_log(f"{EXE_NAME}_selftest.log")
        return code
    xlsx = [a for a in args if a.lower().endswith(SUPPORTED_EXT)]
    if xlsx:
        code = run_cli(xlsx)
        write_log(f"{EXE_NAME}_cli.log")
        return code
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())
