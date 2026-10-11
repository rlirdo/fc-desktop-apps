# -*- coding: utf-8 -*-
"""GUI 自動化截圖（Windows）：實際啟動主視窗、跑一次檢索與缺口矩陣、匯出，並把每個視窗截圖存到 shots/。

用法（在 lit-gap-finder\\ 下）：
    python selftest/gui_shots.py                 → 真核心模組＋真實網路檢索（serpentine deep sea water）
    python selftest/gui_shots.py --stub          → 離線假資料
    python selftest/gui_shots.py --out D:\\某資料夾 → 匯出檔放這裡（預設暫存資料夾）
截圖：selftest/shots/01_檢索頁.png …；結果摘要寫 shots/shots_report.json
"""
import ctypes
import json
import os
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
SHOTS = os.path.join(HERE, "selftest", "shots")


def main():
    stub = "--stub" in sys.argv
    out_dir = None
    if "--out" in sys.argv:
        out_dir = sys.argv[sys.argv.index("--out") + 1]
    os.environ["LITGAP_BACKEND"] = "stub" if stub else "real"
    os.environ["LITGAP_NO_DIALOG"] = "1"          # 匯出完成等訊息不阻塞；對話框另外單獨截圖
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:  # noqa: BLE001
        pass
    import tkinter as tk
    from PIL import ImageGrab
    from ui import backend as BK
    from ui import dialogs, windows, progress
    from ui.main_window import MainWindow
    import app

    BK.init()
    os.makedirs(SHOTS, exist_ok=True)
    out_dir = out_dir or tempfile.mkdtemp(prefix="FCLitGap_gui_")
    report = {"backend": BK.STATUS, "shots": [], "out_dir": out_dir}
    root = tk.Tk()
    mw = MainWindow(root, app.VERSION)
    root.geometry("1100x720+40+30")
    root.update()

    def pump(sec):
        t0 = time.time()
        while time.time() - t0 < sec:
            root.update()
            time.sleep(0.02)

    def shot(win, name):
        try:
            _shot(win, name)
        except tk.TclError as e:
            print("略過截圖（視窗已關閉）：", name, e)

    def _shot(win, name):
        win.update_idletasks()
        win.lift()
        win.attributes("-topmost", True)
        pump(0.6)
        x, y = win.winfo_rootx(), win.winfo_rooty()
        w, h = win.winfo_width(), win.winfo_height()
        # 含標題列：往上抓 ~32px
        img = ImageGrab.grab(bbox=(x - 8, max(0, y - 34), x + w + 8, y + h + 8), all_screens=True)
        p = os.path.join(SHOTS, name + ".png")
        img.save(p)
        win.attributes("-topmost", False)
        report["shots"].append(p)
        print("截圖：", p)

    def wait_task(limit=600):
        t0 = time.time()
        while mw.task is not None and not mw.task.finished and time.time() - t0 < limit:
            root.update()
            time.sleep(0.05)
        return mw.task is not None and mw.task.finished

    # ① 檢索頁
    mw.txt_kw.delete("1.0", "end")
    mw.txt_kw.insert("1.0", "serpentine; 蛇紋石\ndeep sea water; 深層海水\nfar infrared; 遠紅外線")
    mw.var_limit.set(30)
    mw.var_dir.set(out_dir)
    mw.nb.select(mw.tab_search)
    shot(root, "01_主視窗_檢索頁")

    # 進度視窗（真實檢索進行中）
    mw.start_search()
    pump(2.5)
    if mw.task is not None and mw.task.win is not None and not mw.task.finished:
        shot(mw.task.win, "02_進度視窗_檢索中")
    ok = wait_task()
    report["search_done"] = ok
    report["n_papers"] = len(mw.papers)
    report["api_errors"] = list(getattr(mw, "_api_errors", []))
    print("檢索完成：", len(mw.papers), "筆")

    # 缺口矩陣
    mw.start_gap()
    pump(1.5)
    if mw.task is not None and mw.task.win is not None and not mw.task.finished:
        shot(mw.task.win, "03_進度視窗_缺口矩陣")
    wait_task()
    report["gap_pairs"] = len((mw.gap or {}).get("pairs", []))
    report["gap_triples"] = len((mw.gap or {}).get("triples", []))
    report["gap_counts"] = [(r.get("groups"), r.get("count")) for r in (mw.gap or {}).get("pairs", []) + (mw.gap or {}).get("triples", [])]
    mw.nb.select(mw.tab_ind)
    if mw.gap_tree.get_children():
        mw.gap_tree.selection_set(mw.gap_tree.get_children()[0])
    shot(root, "04_主視窗_指標與缺口頁")

    # ③ 結果頁：選第一筆、示範章節下拉
    mw.nb.select(mw.tab_res)
    root.update()
    kids = mw.table.tree.get_children()
    if kids:
        mw.table.tree.selection_set(kids[0])
        mw.show_detail(mw.table.selected_paper())
        mw.table.set_chapter(kids[0], "2.1 文獻回顧")
        if len(kids) > 3:
            mw.table.move_item(kids[3], 0)            # 拖曳排序（程式模擬放開位置）
    shot(root, "05_主視窗_結果頁")
    if kids:
        mw.table.tree.xview_moveto(1.0)              # 章節欄在最右側，先捲過去
        pump(0.4)
        root.lift()
        root.focus_force()
        pump(0.3)
        mw.table.edit_chapter(mw.table.tree.get_children()[0])   # 會自動展開下拉清單
        pump(0.8)
        # 不用 shot()：topmost 會把主視窗蓋在下拉清單之上
        x, y = root.winfo_rootx(), root.winfo_rooty()
        w, h = root.winfo_width(), root.winfo_height()
        img = ImageGrab.grab(bbox=(x - 8, max(0, y - 34), x + w + 8, y + h + 8), all_screens=True)
        p6 = os.path.join(SHOTS, "06_結果頁_章節下拉.png")
        img.save(p6)
        report["shots"].append(p6)
        report["dropdown_editor_alive"] = mw.table._editor is not None
        print("截圖：", p6)
        mw.table._close_editor()

    # 子視窗
    w = mw.open_settings()
    shot(w, "07_設定視窗")
    w.destroy()
    w = mw.open_about()
    shot(w, "08_關於視窗")
    w.destroy()
    w = mw.open_manual()
    for k, v in {"authors": "王小明、陳大華", "year": "2021", "title": "蛇紋石於綠色化學教學之應用",
                 "container": "臺灣化學教育", "volume": "40", "issue": "2", "pages": "13-19"}.items():
        w.vars[k].set(v)
    w.preview()
    shot(w, "09_手動加入中文文獻")
    w.add()
    root.update()
    d = dialogs.FcDialog(root, "匯出完成", "已匯出 8 個檔案到：\n" + out_dir + "\n\n要開啟資料夾嗎？",
                         kind="ask", buttons=["開啟資料夾", "關閉"], modal=False)
    shot(d, "10_對話框_詢問")
    d.destroy()
    d = dialogs.FcDialog(root, "發生錯誤", "「檢索文獻」失敗：連線逾時\n\n已取得的結果保留，可照常匯出。",
                         kind="error", detail="requests.exceptions.ConnectTimeout: …", modal=False)
    shot(d, "11_對話框_錯誤")
    d.destroy()

    # 匯出全部
    mw.start_export()
    wait_task()
    files = []
    for r, _d, fs in os.walk(out_dir):
        files += [os.path.join(r, f) for f in fs]
    report["export_files"] = sorted(files)
    report["dialog_history"] = dialogs.HISTORY[-5:]
    print("匯出：", len(files), "個檔案 →", out_dir)
    with open(os.path.join(SHOTS, "shots_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2, default=str)
    root.destroy()


if __name__ == "__main__":
    main()
