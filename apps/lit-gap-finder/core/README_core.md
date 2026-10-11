# core／FC LitGap Finder 核心模組（施工者 A）

規格：`SPEC_litgap.md` 第 2–5 節。GUI 只需 import 下列函式；所有長作業都接受 `progress(msg, frac)` 與 `cancel_event`（`threading.Event`，被設定時拋 `core.models.Cancelled`）。

| 模組 | 主要介面 | 說明 |
|---|---|---|
| `models.py` | `Paper`、`load_samples()`、`Cancelled`、`CREDIT`、`VERSION`、`DEFAULT_CHAPTERS`、`TYPE_LABELS`、`INDICATOR_NAMES` | 契約欄位外另有擴充欄（url、editors、patent_number、tldr、doi_ok、oa_pdf_alts、note…），皆有預設值 |
| `search.py` | `search(keywords, year_from, year_to, lang, limit, progress, cancel_event, mailto=None)` | 每個元素一組（組內 `;` 分同義詞，組間 AND）；元素是 DOI 時直接取回。Crossref＋OpenAlex 合併，DOI 小寫去重；`search.last_errors` 為部分失敗訊息 |
| | `find_by_title(title, mailto, year)` | 回測用：題名找回單篇 |
| `enrich.py` | `enrich(papers, email, progress, cancel_event, use_s2=None, s2_max=None)` | Unpaywall（OA、授權、DOAJ）＋doi.org HEAD＋Semantic Scholar（≤1 req/s、連續失敗 2 次降級）；`enrich.status` 給狀態列 |
| `indicators.py` | `evaluate(papers, settings)` | 寫入 `indicators{doi,oa,type,year,issn,lang,(manual),relevance}`＝`(True/False/None, 說明)` 與 `relevance`；`keep_mode="strict"` 時不合格者 `include=False` |
| `gap.py` | `gap_matrix(groups, threshold, progress, cancel_event)`、`matrix_table(gap)`、`estimate_requests(n)` | 回傳 `pairs/triples`（count、gap、query、top[Paper×10]）、`singles`、`trend`、`trend_by_group`、`elapsed_sec`、`errors`、`budget_exhausted` |
| `apa.py` | `format_apa7(p)`、`format_apa7_runs(p)`、`in_text(p)`、`sort_key(p)`、`sentence_case(t)` | runs 的 style：`""`、`"i"` 斜體、`"b"` 粗體（中文刊名）；排序＝先中文（Windows zh-TW 筆畫序）後英文 |
| `export.py` | `export_all(papers, gap, out_dir, name, options, progress, cancel_event)`、`default_name(keywords)` | 寫到 `out_dir\name\`：參考文獻_APA7.docx／.pdf、文獻清單與指標.xlsx、缺口矩陣.xlsx、趨勢圖.png、references.ris／.bib、查詢紀錄.json。options：`group_by_chapter`、`order="apa"|"custom"`、`only_included`、`chapters`、`keywords`、`settings`、`mailto`、`pdf_engine="auto"|"word"|"reportlab"` |
| `oa_download.py` | `download_oa(papers, out_dir, progress, cancel_event)` | 只下 Unpaywall／OpenAlex 的 OA PDF（含 Unpaywall 備用 OA 位置）；寫 `OA全文\OA下載紀錄.csv` |
| `cache.py` | `get(key)`、`set(key, v)`、`clear()`、`stats()` | SQLite，7 天；`raw:crossref:<doi>`／`raw:openalex:<id>` 存原始 JSON |
| `config.py` | `load()`、`save(cfg)`、`DEFAULTS`、`home_dir()` | `%APPDATA%\FCLitGap\config.json`；環境變數 `FCLITGAP_HOME` 可改位置（自測／CI） |
| `net.py` | `request()`、`get_json()`、`ApiError(.status, .budget)` | UA 含 mailto、timeout 20 s、重試 3 次、一般 ≤5 req/s、S2 ≤1 req/s；每日額度用完（429＋長 Retry-After）立即失敗 `budget=True` |
| `selftest_core.py` | `run_offline_selftest(out_dir, log)` | 離線樣本跑完整流程，回傳 True/False（給 `--selftest`） |
| `stub.py` | 與上列同名函式 | 離線假資料版（GUI 先行開發用） |

GUI 換用真模組：
```python
from core.search import search; from core.enrich import enrich; from core.indicators import evaluate
from core.gap import gap_matrix, matrix_table, estimate_requests; from core.apa import format_apa7, in_text, sort_key
from core.export import export_all, default_name; from core.oa_download import download_oa
from core import cache, config
```

## 重要限制（2026/10/11 實測）
- **OpenAlex 已改額度制**：沒有 API key 時，同一 IP 每天共用 1000 credits，每次 search 耗 10 credits（約 100 次檢索／日，UTC 0 時重置）；DOI 單筆查詢不耗額度。五組關鍵字的缺口矩陣要 26 次檢索。額度用完時 `search` 仍回傳 Crossref 結果、`gap_matrix` 立即停止並標 `budget_exhausted`。`config["openalex_api_key"]` 或環境變數 `FCLITGAP_OPENALEX_KEY` 可填使用者自行申請的免費 key（選填；不會寫入匯出檔與查詢紀錄）。
- MDPI 的 PDF 網址會擋非瀏覽器程式（403）；工具不偽裝瀏覽器，紀錄會寫「下載失敗，請以 DOI 連結手動開啟」。
- 只有 PubMed 典藏紀錄的舊文獻（如 J Cosmet Sci 2002）OpenAlex 沒有刊名，會在 `note` 標「刊名缺，待人工補正」。
- PyInstaller：python-docx 需打包其 `templates` 資料夾；reportlab 字型取 `C:\Windows\Fonts\kaiu.ttf`（備援 msjh），英文字取 times*.ttf。

## 2026/10/11 修正（第 2 輪，依 ISSUES_from_B.md 與統籌指示）
1. **DOI 取回缺頁碼**：重現 B 的狀況。原因是 Vatansever 2012 的 Crossref 紀錄本來就沒有 `page`，頁碼只在 OpenAlex `biblio`；B 測試時 OpenAlex 那一步失敗（額度用完），合併後就沒有頁碼。`merge()` 本來就是「非空優先」，所以改成兩件事：
   - `enrich` 對缺頁碼／文章號的期刊論文一律查 Semantic Scholar，以 `journal.pages` 補頁碼（S2 的卷號曾見錯誤，所以只補頁碼），`note` 註明「頁碼取自 Semantic Scholar」；新增 `enrich.apply_s2()`。
   - OpenAlex `first_page`／`last_page` 合併照舊，縮寫的末頁會補全（175–84 → 175–184）。
2. **趨勢 2026 年異常值 52,796**：查 B 的快取原始回應，篩選 `publication_year:2017-2026` 有生效（meta.count 88,377＝各年加總），52,796 是 OpenAlex 把該年資料這樣回傳，不是篩選失敗，也不是額度用完的錯誤回應；推測是資料庫批次匯入（今日額度已用完，無法再細查是哪批）。處理方式：
   - 只取近 10 個**完整年度**（今年排除）。
   - `gap.find_outliers()`：某年 > 其他年中位數 3 倍（且多 100 篇以上）標為異常值，圖上紅圈截頂並註記，不計入刻度。
   - 新增 `trend_status`（每組 ok／failed／skipped）、`trend_years`、`partial_year_excluded`。查詢失敗時趨勢圖會寫明「OpenAlex 趨勢查詢失敗」，改用本次檢索結果繪製，不會畫成 0。
3. **語言判定**：`Paper.is_chinese` 依序看——題名含「Powered by NICT」一律非中文 → `language` 欄（zh＝中文，ja／en 等＝非中文）→ 沒有 language 時，含假名＝非中文，否則題名漢字占文字 ≥30% 或第一作者為漢字才算中文。Crossref 解析不再因題名有漢字就自動填 `zh`。
4. **reportlab 證據表**：改用 `Table`，欄寬依頁寬自動分配（符號欄固定窄、文字欄依內容長度比例），儲存格 Paragraph 自動折行，表頭每頁重複。✔✘①② 等符號改用 Segoe UI Symbol（原本 Times 缺字顯示成方塊）。
5. **缺口矩陣狀態**：回傳 `status="ok"|"partial"|"failed"`（全部查詢失敗＝failed）。矩陣格中未取得者寫「查詢失敗」，不寫 0；說明頁加查詢狀態、趨勢年度、趨勢狀態與異常值。
6. **API 金鑰（定案：開放選填）**：`config["openalex_api_key"]` 預設空字串，只存在 `%APPDATA%\FCLitGap\config.json`。只在送出請求時由 `net.openalex_params()` 附上，不進快取鍵，查詢紀錄 json 會濾掉名稱含 key／password 的設定，也不寫入任何匯出檔。額度用完的訊息維持「可於設定填入免費 OpenAlex API key」。
- 新增離線回歸測試 `selftest_core.regression_checks()`（6 項）：Vatansever 兩種合併順序與 Crossref＋S2 補頁碼、語言判定、異常值偵測、全失敗 status。`run_offline_selftest()` 會自動執行。
