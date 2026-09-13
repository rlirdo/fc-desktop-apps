# fc-desktop-apps — FC 自行開發桌面版軟體

三支給非工程師雙擊即用的桌面程式（Python + tkinter，PyInstaller 封裝），
每次推送都會在 GitHub Actions 的**乾淨 Windows 與 macOS 執行器**上封裝並跑自我測試。

| 資料夾 | 程式名 | 用途 |
|---|---|---|
| `apps/weekly-report` | **WeeklyReportMaker** 週報產生器 | 填表→產出碩士生標準版週報 PPTX（11–15 頁、逐字稿、內建 6 項驗收） |
| `apps/name-masker` | **NameMasker** 姓名遮罩 | Excel「姓名」欄第 2 字改 O，另存 `原檔名02.xlsx`，完全離線 |
| `apps/homework-wordcloud` | **HomeworkWordCloud** 學生作業文字雲 | Zuvio 匯出→姓名遮罩→每題文字雲→分析簡報 PPTX |
| `apps/club-manager` | **ClubManager** 仰望社團管理系統 | 本機伺服器＋系統瀏覽器開啟社團管理系統（總版＋各校版），資料存瀏覽器 |

## 下載

到 [Actions](../../actions) 最新一次成功的 `build-desktop-apps`，下載 artifacts：
`<程式名>-windows`（.exe）或 `<程式名>-macos`（.app 的 zip）。

## 第一次執行

- **Windows**：SmartScreen 會擋（未購買程式碼簽章）→ 點「其他資訊」→「仍要執行」。
- **macOS**：Gatekeeper 會擋 → 解壓後對 .app 按右鍵→「打開」→再按「打開」。
- 不支援 iOS／iPadOS。

## 自我測試

每支程式都有 `--selftest`（無視窗、用合成資料跑完整流程並驗證），例如：

```
WeeklyReportMaker.exe --selftest
```

## 隱私

所有程式完全離線；範例與測試資料皆為合成，repo 內不含任何真實學生個資。
