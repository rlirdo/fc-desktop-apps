# core 驗證報告（施工者 A，2026/10/11）

環境：Windows 11、Python 3.13.5、`PYTHONUTF8=1`、Word（COM）可用。暫存與輸出：`D:\WinTemp\claude\D--02-aitt-claude-agent\a3f42f6a-978b-4f95-80ad-cff155ab54d2\scratchpad\litgap\A\`（腳本在 `scripts\`）。

## (1) 離線：samples.json → indicators → apa → export
- 腳本 `scripts\t1_offline.py`，執行時以 monkeypatch 封鎖 `socket.connect`（確認離線流程完全不連網）。
- 樣本 15 筆：13 筆書目來自 Crossref／查證表（含 26 作者的 Harris 2020 NumPy），2 筆中文（TW201402514A 專利、1 筆手動加入示範），1 筆無卷期（網頁）。
- 兩組匯出（一般排序＋Word COM 轉 PDF／依章節分組＋reportlab PDF），8 個檔案全部成功：docx、pdf、文獻清單與指標.xlsx、缺口矩陣.xlsx、ris、bib、趨勢圖.png、查詢紀錄.json。
- 開檔檢查：
  - docx：懸掛縮排 −1.27 cm、刊名與卷號斜體、12 pt Times New Roman＋標楷體 eastAsia、行距 2.0；先中文（自＜吳，筆畫序正確）後英文字母序；手動加入者後附灰字「〔使用者自填，未經 API 驗證〕」；附錄一指標證據表、附錄二檢索策略都有；分章版依「2.1 文獻回顧／3 方法／未分類」分組。
  - PDF：Word 版 6 頁；reportlab 版 3 頁，中文標楷體、英文 Times（含 č、ć 等字）、斜體正確（第一版 Friščić 缺字，已改用 Windows times*.ttf 修好）。
  - xlsx：34 欄、凍結首列、DOI 超連結、自動篩選；缺口矩陣 5 個工作表。
  - 趨勢圖 PNG：1600×1000，折線＋前 10 高被引橫條，右下 CREDIT。
  - 隱私：查詢紀錄.json 中 email 只出現 1 次（mailto 欄位）；ris／bib／xlsx 都是 0 次；設定中名稱含 key／password 的欄位不寫出。
- `core\selftest_core.run_offline_selftest()`：PASS（8/8 檔），可直接給 `--selftest` 用。

## (2) 連網回測：規格第 5 節 15 篇（`scripts\t2_backtest.py`，結果 `backtest_results.json`）
以**題名**檢索（Crossref query.title＋OpenAlex title.search），有 DOI 者另以 DOI 檢索。

| # | 文獻 | 題名找回 | DOI 找回 | 比對基準 | APA 7 逐字比對 |
|---|---|---|---|---|---|
| 1 | Ajzen 1991 | ✔ | ✔ | 查證表未列 | — |
| 2 | Yoo 2002 J Cosmet Sci | ✔（OpenAlex，PMID 12053208） | 本無 DOI | 查證表未列 | 見下方差異 A |
| 3 | Lipshutz 2011 JOC | ✔ | ✔ | 文獻清單_取代TPGS.xlsx | 差 1 處（B） |
| 4 | Vatansever 2012 | ✔ | ✔ | 查證表 R1 | **完全相同** |
| 5 | Do & Friščić 2017 | ✔ | ✔ | 查證表 R3 | **完全相同** |
| 6 | Sano 2018 | ✔ | ✔ | 查證表 R4 | **完全相同**（修正前少「Article」，已修） |
| 7 | Cortes-Clerget 2019 NatCommun | ✔ | ✔ | 文獻清單_取代TPGS.xlsx | 差 2 處（C） |
| 8 | Cortes-Clerget 2021 ChemSci | ✔ | ✔ | 查證表 R5 | 只差 DOI 大小寫（D） |
| 9 | Garzón 2021 | ✔ | ✔ | 查證表 R6 | **完全相同** |
| 10 | Ivanova 2022 | ✔ | ✔ | 查證表 R7 | **完全相同** |
| 11 | Chen 2023 | ✔ | ✔ | 查證表 R8 | **完全相同** |
| 12 | Hung 2025 | ✔ | ✔ | 查證表 R9 | **完全相同** |
| 13 | Wu 2021 Crystals | ✔ | ✔ | 查證表未列 | — |
| 14 | Long 2019 Minerals | ✔ | ✔ | 文獻清單_取代TPGS.xlsx | 差 1 處（E） |
| 15 | Belhaj 2020 JPEPT | ✔ | ✔ | 查證表未列 | — |

**找回率：題名 15/15、DOI 14/14**（Yoo 2002 本來就沒有 DOI）。對查證表 8 筆：7 筆逐字相同，1 筆只差 DOI 大小寫。

差異：
- A　Yoo 2002：OpenAlex 只有 PubMed 典藏紀錄，沒有刊名，輸出 `Yoo, B. H., Park, C. M., Oh, T. J., Han, S. H., Kang, H. H., & Chang, I. S. (2002). Investigation of jewelry powders radiating far-infrared rays and the biological effects on human skin. https://pubmed.ncbi.nlm.nih.gov/12053208`，`note` 標「刊名缺，待人工補正」（應補 *Journal of Cosmetic Science, 53*(3), 175–184）。作者改用 raw_author_name 後與 PDF 署名一致（原本 OpenAlex 歸併名為 B.-O. Yoo）。
- B　Lipshutz 2011：本工具「*The* Journal of Organic Chemistry」，清單寫「Journal of Organic Chemistry」。Crossref 登錄的正式刊名有 The，以本工具為準。
- C　Cortes-Clerget 2019：本工具「*10*(1), Article 2169」，清單寫「10, 2169」。APA 7 規定有期號要列、文章編號前加 Article，以本工具為準。
- D　Cortes-Clerget 2021：本工具 `10.1039/d0sc06000c`，查證表 `10.1039/D0SC06000C`。Crossref／OpenAlex 都回小寫，DOI 不分大小寫，連結相同。
- E　Long 2019：本工具「Article 629」，清單寫「629」。APA 7 應有「Article」。

其他 4 筆 APA 7 輸出：
- `Ajzen, I. (1991). The theory of planned behavior. Organizational Behavior and Human Decision Processes, 50(2), 179–211. https://doi.org/10.1016/0749-5978(91)90020-t`
- `Wu, S., He, M., Yang, M., Zhang, B., Wang, F., & Li, Q. (2021). Near-infrared spectroscopy study of serpentine minerals and assignment of the OH group. Crystals, 11(9), Article 1130. https://doi.org/10.3390/cryst11091130`
- `Belhaj, A. F., Elraies, K. A., Mahmood, S. M., Zulkifli, N. N., Akbari, S., & Hussien, O. S. (2020). The effect of surfactant concentration, salinity, temperature, and pH on surfactant adsorption for chemical enhanced oil recovery: A review. Journal of Petroleum Exploration and Production Technology, 10(1), 125–137. https://doi.org/10.1007/s13202-019-0685-y`
- Yoo 2002 見 A。

doi.org 驗證：14 筆 HEAD 全部 302（`doi_ok=True`）。隨機 5 筆 DOI（seed 20261011）的 HTTP 結果：

| DOI | doi.org | 跟隨轉址後 |
|---|---|---|
| 10.3390/mti5070037 | 302 → mdpi.com | 403（MDPI 擋程式，瀏覽器可開） |
| 10.1038/s41467-019-09751-4 | 302 → nature.com | 200 |
| 10.1007/s13202-019-0685-y | 302 → link.springer.com | 200 |
| 10.3390/min9100629 | 302 → mdpi.com | 403（同上） |
| 10.1016/j.heliyon.2018.e00923 | 302 → linkinghub.elsevier.com | 200 |

回測耗時：第一次 30.5 s（無快取），重跑 7.0 s（快取）。

## (3) 缺口矩陣：五組關鍵字（`scripts\t3_gap.py`，結果 `gap_results.json`、`out_live2\...\缺口矩陣.xlsx`）
組 1 serpentine; 蛇紋石｜組 2 deep sea water; deep ocean water; 深層海水｜組 3 micelle; TPGS; surfactant｜組 4 far-infrared emissivity; 遠紅外線｜組 5 augmented reality education; WebAR；門檻 < 5 篇＝缺口。

**第 1 次（25.3 s，26＋5 次請求全部成功）**發現：中文詞沒加引號時，OpenAlex 會把中文拆成單字比對，大量誤中日文舊文獻（例如組 2×組 4 有 14,197 篇，前幾名是 1941 年的日文文獻）。所以改成**中文詞與多字詞一律加引號做片語比對**。

**第 2 次（加引號後）兩兩矩陣**（對角線＝單組篇數）：

| | 1 蛇紋石 | 2 深層海水 | 3 微胞/TPGS/界面活性劑 | 4 遠紅外放射率 | 5 AR 教育/WebAR |
|---|---|---|---|---|---|
| 1 | 117,819 | 105 | 3,590 | 21 | **0 缺口** |
| 2 | 105 | 11,801 | 258 | **0 缺口** | **0 缺口** |
| 3 | 3,590 | 258 | 983,253 | 28 | 8 |
| 4 | 21 | **0 缺口** | 28 | 3,680 | **0 缺口** |
| 5 | **0 缺口** | **0 缺口** | 8 | **0 缺口** | 1,957 |

三三組合：1×2×3＝5；1×2×4＝0、1×2×5＝0、1×3×4＝1、1×3×5＝0、1×4×5＝0（都是缺口）；2×3×4、2×3×5、2×4×5、3×4×5＝**未取得**（OpenAlex 額度用完，429）。年度趨勢這一次也沒取得。
**耗時**：第 2 次 278.5 s，大部分是舊版遇 429 時重試等待。已改成遇「每日額度用完」立刻停止（實測 0.2 s 回報）。

**重大發現：OpenAlex 已改額度制。** 沒有 API key 時，同一 IP 每天共用 1000 credits，每次 search 耗 10 credits，約 100 次；回應標頭 `X-RateLimit-Credits-Required: 10`，額度在 UTC 0 時重置。DOI 單筆查詢不耗額度。五組關鍵字一次要 26 次（`gap.estimate_requests(5)`），六組要 42 次。已做的處理：
- 額度用完時，`search` 改回傳 Crossref 結果，並附明確訊息；`gap_matrix` 立刻停止並標 `budget_exhausted=True`；趨勢預設只查第 1 組。
- Crossref 沒有 AND 功能，多組關鍵字時改為每組都要在題名或摘要出現才收（避免 “The Serpentine River” 這類誤中）。
- 新增選填的 `openalex_api_key`（使用者自行到 openalex.org 申請免費 key）。

## (4) OA 下載實測（`scripts\t4_oa.py`；紀錄 `out_oa\OA全文\OA下載紀錄.csv`）
```
DOI,來源網址,授權 license,下載時間,狀態,檔名
10.3390/min15010062,https://www.mdpi.com/2075-163X/15/1/62/pdf,cc-by,2026-10-11T12:30:45,下載失敗：www.mdpi.com 回應 403（請以 DOI 連結手動開啟）,
10.1038/s41467-019-09751-4,https://www.nature.com/articles/s41467-019-09751-4.pdf,cc-by,2026-10-11T12:30:48,已下載,2019_Cortes-Clerget_Bridging_the_gap_between_transition_metal.pdf
10.1007/s13202-019-0685-y,https://link.springer.com/content/pdf/10.1007/s13202-019-0685-y.pdf,cc-by,2026-10-11T12:30:49,已下載,2020_Belhaj_The_effect_of_surfactant_concentration_salinity.pdf
10.1021/jo101974u,,,,非 OA 或無 PDF 連結（僅提供 DOI）,
```
3 筆 OA 中有 2 筆成功（1.34 MB、0.55 MB，檔頭都是 `%PDF-`）。MDPI 擋程式下載（403），工具不偽裝瀏覽器，只照實記錄；Unpaywall 若有其他 OA 位置會依序再試。非 OA 不嘗試其他來源。

## 未完成／待決事項
1. 缺口矩陣有 4 個三三組合與年度趨勢因 OpenAlex 今日額度用完而沒有實測取得，要等 UTC 0 時（台灣 08:00）後重跑 `scripts\t3_gap.py`。B 整合回測也會吃同一個 IP 的額度。
2. OpenAlex 的 `search` 是全文檢索，大組合（如組 3）篇數偏高、前 10 篇相關性較鬆。可考慮改用 `title_and_abstract.search` 過濾，但額度用完無法驗證，所以沒有改。
3. 是否允許在設定存 OpenAlex API key，要請 Fable／使用者決定（規格寫「不存任何密碼」；目前預設空值、選填）。
4. 規格第 5 節的「人工點開核對」是用 HTTP 狀態代替，沒有截圖。

## 附記：B 整合回饋（ISSUES_from_B.md）已處理
趨勢只取完整年度、日文不套中文 APA、證據表改符號、圖例截字、macOS 字型路徑——5 項都已修，修後離線 selftest（Word COM 與 reportlab）PASS。
