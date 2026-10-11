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

## A 回覆（2026/10/11）— 5 項皆已修，selftest（Word 與 reportlab 兩種 PDF）PASS
1. 趨勢改取「近 10 個完整年度」（今年除外），gap.yearly_counts 與趨勢圖標題同步（例：2016–2025）。
2. `Paper.is_chinese`：題名含假名或 language 非 zh → 不算中文，改走英文 APA、排英文區。
3. docx 證據表改只放 ✔／✘／？（理由在 xlsx），表前加圖例；相關性取 1 位小數。
4. 趨勢圖圖例加寬並依實際字寬截字加「…」。
5. `_cjk_font_path()` 加 macOS Arial Unicode.ttf 與 Linux Noto/ukai 路徑。
另：OpenAlex 無 key 每日約 100 次 search（同 IP 共用），今日已用完；B 整合回測請留意（DOI 單筆查詢不耗額度）。詳見 TEST_REPORT_core.md。

## 整合回測（B，經 GUI `start_search` 聯集模式，14 個 DOI＋Yoo 題名，OpenAlex 今日額度已用完）
結果：找回 14/15（Yoo 2002 只有 OpenAlex 有、今日額度用完而未找回，屬預期降級）；與基準 APA 逐字相同 6/11，差異中 4 筆是 A 報告已列的 B/C/D/E（刊名 The、Article 編號、DOI 大小寫），**新增 1 筆**：

6. **Vatansever 2012 以 DOI 檢索時頁碼遺失**：GUI 以 `search(["10.1515/plm-2012-0034"])` 取回後輸出
   `… Photonics & Lasers in Medicine, 1(4). https://doi.org/10.1515/plm-2012-0034`（少「, 255–266」）；A 以題名回測時是完整的。
   疑似 `crossref_by_doi`＋`openalex_by_doi` 的 `merge()` 用到空頁碼（或 OpenAlex 額度用完時的快取紀錄沒有 pages）。請檢查 merge 是否「非空優先」。
7. 額度用完訊息寫「可於設定填入免費 OpenAlex API key」，但 GUI 設定視窗**沒有** API key 欄位（規格「不存任何密碼」，A 報告列為待決）。
   若決定開放，B 可在設定視窗加選填欄；若不開放，請把訊息改為「請明日再試」。

## A 回覆（第 2 輪，2026/10/11）
- 第 6 項（Vatansever DOI 缺頁碼）：已重現。原因是 Crossref 本身沒有 page，OpenAlex 那一步失敗。已改成 enrich 時以 Semantic Scholar `journal.pages` 補頁碼，並加回歸測試；修後在 OpenAlex 失敗的情況下也能輸出 255–266。**GUI 必須在 search 之後呼叫 enrich，頁碼才會補上。**
- 第 7 項（API key）：統籌定案「開放選填」。core 維持 `config["openalex_api_key"]`（預設空字串），B 可以在設定視窗加選填欄位（建議用遮蔽輸入）；訊息維持「可於設定填入」。
- 新增介面：`gap["status"]`（ok／partial／failed）、`gap["trend_status"]`、`gap["trend_outliers"]`、`gap["trend_years"]`。GUI 在 status=="failed" 時請顯示「查詢失敗」，不要顯示「缺口 0 個」；count 為 None 的格請顯示「查詢失敗」。

8. **版本號**：B 已將 app 升為 1.0.1（`app.VERSION`、build.spec、README）；`core\models.py` 的 `VERSION` 仍是 "1.0.0"，`查詢紀錄.json` 的 version 會寫 1.0.0。請 A 改為 "1.0.1"。
