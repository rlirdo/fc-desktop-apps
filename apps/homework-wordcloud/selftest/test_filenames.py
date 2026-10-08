# -*- coding: utf-8 -*-
"""
test_filenames.py — 2.6 輸出檔名規則的單獨測試（全部用合成題名與假 ID，不碰任何真實資料）。

    python selftest\\test_filenames.py

規則：每題輸出檔名＝`Q{nn}_{題目全名}_{題目ID}`
  * 題目：下載清單題名 → 已具名輸入檔的題目 → 匯出檔內題目 → 輸入檔名
  * `\\ / : * ? " < > |` 換半形空格、連續空白合併、去頭尾空白與結尾「.」、最多 80 字
  * 題目 ID：輸入檔名結尾 8 位數字（10 位取前 8 位）；沒有就省略 `_ID`
同一組檢查也包含在 `python app.py --selftest`（第 12 項）裡，打包後的 exe 一樣會跑。
"""
import os
import sys
import shutil
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import pipeline as P                                    # noqa: E402


def main():
    tmp = tempfile.mkdtemp(prefix="hwwc_fn_[課程] ")       # 含中括號與空白：順便測 glob 跳脫
    try:
        cfg = P.load_config()
        cfg["course_name"] = cfg.get("course_name") or "我的課程"
        problems = P._selftest_filenames(tmp, cfg, print)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    for p in problems:
        print("  [失敗]", p)
    print("FILENAME TEST OK" if not problems else "FILENAME TEST FAILED")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
