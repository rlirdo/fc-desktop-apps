# B → A 整合回饋（2026/10/11，GUI 實跑 serpentine／deep sea water／far infrared，聯集檢索 173 筆，Word 轉 PDF）

B 的 GUI 呼叫方式（供對照）：
- 檢索：預設「各組分別（聯集）」→ 每組呼叫一次 `search(["serpentine; 蛇紋石"], yf, yt, lang, limit, progress, cancel_event, mailto=…)`，GUI 端以 DOI 去重；
  「交集」模式 → `search(["serpentine; 蛇紋石", "deep sea water; 深層海水", …])`。
- 指標：`evaluate(papers, settings)`，settings 鍵＝config 同名（checks/types/year_from/year_to/classic/lang/weights/keep_mode）＋ `keywords: list[list[str]]`。
- 匯出：`export_all(..., options={"order": "apa"|"custom", "keywords": groups, "only_included": True, "pdf_engine": "auto", "group_by_chapter", "chapters", "settings", "mailto", "query", "version"})`。
  selftest 用 `pdf_engine="reportlab"`（CI 無 Word）。

## 發現（皆不擋整合，B 已用現況完成測試；請 A 斟酌）

1. **趨勢圖 2026 年異常**：`趨勢圖.png`「far infrared」2026 年＝52,796 篇（其他年約 2,000–5,000）。
   疑似 OpenAlex 當年度資料或 `publication_year` 篩選邊界問題。建議趨勢只取「近 10 個完整年度」（今年除外）或今年另標「部分年度」。
2. **日文文獻被當中文格式化**：OpenAlex 回傳 JST 機器翻譯的日文紀錄（題名含「【Powered by NICT】」、平假名／片假名），
   `is_chinese` 只看 CJK → 套中文 APA（全形標點、「、」分隔作者），且排在「中文」區塊；英文作者名（Cao, C.、Cheng-hua, L.）也被中文格式。
   建議：含假名（\u3040–\u30ff）視為日文→用英文 APA 規則；或語言非 zh 時不套中文格式；「Powered by NICT」題名可考慮排除或標記。
3. **PDF 指標證據表過窄**：Word 轉出的 PDF 最後幾頁，證據表 11 欄擠在直式 A4，「doi.org 可解析」「範圍 2016–2026」逐字斷行（每列高 6–8 行）。
   建議證據表所在節改橫向，或縮短判定文字（例如只放 ✔／✘／?，理由移到 xlsx）。
4. 趨勢圖圖例「2. deep sea wate」被截字（圖例寬度）。
5. **macOS PDF 中文字型**：`_cjk_font_path()` 只找 `%WINDIR%\Fonts`，macOS 會退回 Helvetica（中文變空白方塊，不會報錯）。
   可加 `/Library/Fonts/Arial Unicode.ttf`、`/System/Library/Fonts/Supplemental/Arial Unicode.ttf`（TrueType 外框，reportlab 可用；PingFang.ttc 是 CFF 外框不能用）。
