# -*- coding: utf-8 -*-
"""HomeworkWordCloud 2.1 核心模組。

roster        原始名單（Excel／CSV／選課名單 PDF）→ 固定學生編號 → 對照表 xlsx
mask          去識別化（姓名第 2 字改 O、學號／email 全 O、學生編號）
parse_zuvio   解析匯出檔 → 每題一份去識別化 CSV（學生編號為主鍵）
analyze       斷詞、三個重點與概念矩陣、覆蓋率、提問抽取與兩類分類
wordcloud_gen 每題一張文字雲 PNG
deck          組簡報（封面／總覽／一題一頁／附錄概念矩陣／結尾）＋幾何 QA
"""
__all__ = ["roster", "mask", "parse_zuvio", "analyze", "wordcloud_gen", "deck"]
