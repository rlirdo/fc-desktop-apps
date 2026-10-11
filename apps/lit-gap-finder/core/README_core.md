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
