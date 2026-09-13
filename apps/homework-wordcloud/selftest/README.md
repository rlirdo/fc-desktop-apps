# selftest／合成測試資料

這個資料夾放的是**自我測試**用的東西，裡面**沒有任何真實學生資料**。

- `make_sample_input.py`：產生上一層 `sample_input\` 裡三個合成的 Zuvio 匯出 xlsx
  （姓名、學號、作答內容全部虛構，例如「王小明」「賴恩」）。
  需要重新產生時執行：`python selftest\make_sample_input.py`
- 測試資料本體放在 `..\sample_input\`（同時也是 GUI「跑示範」用的資料，
  所以只保留一份，打包時會一起帶進 exe）。

## 自我測試在測什麼

`python app.py --selftest`（或 `HomeworkWordCloud.exe --selftest`）會：

1. 用 `sample_input\` 的合成資料跑完整流程到系統暫存資料夾；
2. 檢查產出的簡報**至少 4 頁**；
3. 檢查文字雲 PNG **至少 3 張**；
4. 跑隱私檢查，必須 **0 命中**（輸出裡不能出現名冊上的原始姓名）；
5. 全部通過就印 `SELFTEST OK` 並回傳 0，任何一項失敗回傳 1。

詳細紀錄會另外寫到系統暫存資料夾的 `HomeworkWordCloud_selftest.log`
（在 Windows 上就是 `%TEMP%\HomeworkWordCloud_selftest.log`）。
