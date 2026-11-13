# SPEC｜客語語音小幫手 HakkaVoiceHelper 1.0.0

對象：花蓮鳳林國小附設幼兒園客語社群老師（非工程背景，Windows 筆電）。講師：游豐兆。
用途：非商業教學。原則：**畫面與草稿交給工具，客語用字與定稿由老師負責**。

## 0. 為什麼這樣設計（必讀，不可違反）

講師要求：輸入華語句子後，轉出的客語在文法、句法、結構上要符合三個網站的用法與規定：
客家委員會「臺灣客語語音資料庫」speech.hakka.gov.tw、長問翻譯器 translate.bronci.com.tw、Gohakka gohakka.org。

三個網站的規定（2026/11/13 查證）：
- **客委會**：依「客家委員會臺灣客語語音資料庫應用程式介面管理要點」（114/7/8），翻譯、語音合成、語音辨識 API 要先申請核准（政府機關、學校、公益法人可申請，每次三個月，僅限非營利，不得轉授權或提供第三方）；未經授權的 API 使用得暫停或終止服務；使用 API 產生的軟體與衍生著作要註明授權範圍。網站本身：非會員每次 100 字、線上即時翻譯每日 1,000 字。
- **Gohakka**：免費會員翻譯每日 20 字（需登入）；API 只在 Enterprise 方案提供；平台內容 CC BY-SA 4.0；使用語音合成結果須註明「語音來源：GoHakka.org」；不得擷取平台資源訓練 AI 模型。
- **長問翻譯器**：Beta，翻譯內容由 AI 產生；沒有公開 API。

**因此本工具絕對不可以用程式去呼叫、爬取或自動操作這三個網站**（不寫任何對這三個網域的 HTTP 請求、不嵌入網頁自動填表、不讀取其頁面內容）。
「華語 → 客語文字」這一步採「**輔助交接**」：工具把華語放進剪貼簿並用預設瀏覽器開啟官方翻譯網頁，由老師本人在官方網頁按翻譯、複製結果，再回工具按「貼上客語漢字」。這樣客語的用字、文法、句法就是官方工具給的，也符合各站規定。

其他兩步使用公開授權的來源：
- **客語文字 → 語音**：Hugging Face Space `united-link/taiwanese-hakka-tts`（臺灣客語語音合成系統；研發：李鴻欣、陳力瑋；合作：國立聯合大學智慧客家實驗室；模型 `yourtts-htia-240704`，CC BY-NC 4.0）。該 Space 頁尾提供「透過 API 使用」，即 Gradio 公開 API。這是共用的公開示範服務：一次一句依序送、每句之間至少間隔 1 秒、單次工作上限 1,200 字，不做任何平行或批量轟炸。
- **客語語音 → 文字**：本機離線執行 faster-whisper（CTranslate2），模型為國立臺南大學 KWS 中心 `NUTN-KWS/Whisper-Taiwanese-Hakka-model-v0.2.6`（whisper-large-v3-turbo 微調，CC BY-NC 4.0）轉成的 int8 版。音檔不離開電腦。

## 1. 已驗證可用的技術事實（2026/11/13，講師端實測）

### 1-1 語音合成 API（純 requests 即可，不需 gradio_client）
```
BASE = "https://united-link-taiwanese-hakka-tts.hf.space/gradio_api"
POST  BASE/call/text_to_speech   json={"data":[模型, 語者類型, 客製化語音, 語者, 腔調, 語速, 文字]}
      例：["yourtts-htia-240704","預設語者",None,<語者值>,"sixian",1,"恁早啊，細兔仔愛去學校上課。"]
      → {"event_id": "..."}
GET   BASE/call/text_to_speech/{event_id}   （SSE 串流；逐行讀，event: complete 的下一行 data: 是 JSON 陣列）
      → [斷詞字串, 客語拼音字串, {"url": "https://…/gradio_api/file=/tmp/gradio/…/audio.wav", …}]
      event: error 時 data 是錯誤訊息。
GET   上面的 url → WAV 位元組（RIFF）。
```
實測：送出到取得 WAV 約 6 秒；斷詞「恁早 啊 ， 細兔仔 愛去 學校 上課 。」；拼音「an31 zo31 a55，se55 tu55 e31 oi55 hi55 hog5 gau31 song24 ko55。」
- `GET BASE/info` 的 `named_endpoints["/text_to_speech"].parameters` 可取得：語者 enum 與預設值（index 3）、腔調 enum（index 4：sixian, hailu, dapu, raoping, zhaoan, nansixian）、語速範圍（0.5–1.5，越大越慢）。
- 網頁上語者顯示成「（姓名） / 男 / 苗栗」這種標籤（姓名部分不可寫入 repo），API 值是另一個字串；`GET https://united-link-taiwanese-hakka-tts.hf.space/config` 的 components 裡 Dropdown「語者」的 choices 是 [標籤, 值] 配對，可用來顯示友善標籤。
- **語者清單一律在執行時從 /info、/config 取得，原始碼與 repo 內不可寫死任何語者的值或姓名**（它們是真人姓名；repo 有 PII 防線）。取不到清單時，停用合成並顯示連線錯誤。
- 腔調在介面上顯示中文：四縣、海陸、大埔、饒平、詔安、南四縣。

### 1-2 語音辨識
- 本機 Python 3.13 已裝 `faster_whisper`、`ctranslate2`、`av`（PyAV）、`requests`。
- 轉好的模型資料夾（講師端背景轉換中，完成後會在）：
  `D:\WinTemp\claude\80f13883a1be0578\667aa898-45a7-4380-a02a-0f500b3129e9\scratchpad\hk\hakka-whisper-ct2-int8\`（含 model.bin、config.json、tokenizer.json、preprocessor_config.json、vocabulary.*）。同資料夾有測試音檔 `probe.wav`（內容：恁早啊，細兔仔愛去學校上課。）。若你做到辨識測試時模型還沒出現，先把其他部分做完並在回報中說明，不要自己重新下載或轉換。
- 用法：`WhisperModel(path, device="cpu", compute_type="int8")`；`transcribe(file, language="zh", task="transcribe", beam_size=5, vad_filter=False)`。
- 不使用 VAD（避免把 onnxruntime 包進 exe；若 PyInstaller 自動帶入就用 excludes 排除並確認仍可執行）。
- 解碼 mp3／m4a／wav 用 PyAV（faster-whisper 的 `decode_audio` 已使用）。
- 模型搜尋順序：① exe 同層 `models\hakka-whisper-ct2-int8\` ② `%LOCALAPPDATA%\HakkaVoiceHelper\models\hakka-whisper-ct2-int8\` ③ 設定檔指定路徑。都沒有就在分頁 3 顯示清楚說明：「語音辨識需要約 0.8 GB 的模型資料夾，請向講師索取 models 資料夾，放在本程式旁邊」，並提供「選擇模型資料夾…」按鈕。**不要做自動下載**（模型還沒有公開下載點）。
- 輸出的是客語漢字（繁體字）。要華語意思時，提供按鈕「複製並開啟客委會翻譯（客語→華語）」（同樣是剪貼簿＋開瀏覽器）。

## 1-2b 語音辨識：2026/11/13 講師端實測結論（以本節為準，覆蓋 1-2 與分頁 3 中不一致之處）

- 模型已就緒：`D:\WinTemp\claude\80f13883a1be0578\667aa898-45a7-4380-a02a-0f500b3129e9\scratchpad\hk\hakka-whisper-ct2-int8\`（788 MB：model.bin、config.json、tokenizer.json、preprocessor_config.json、vocabulary.json）。本機 faster_whisper 1.2.1、ctranslate2 4.8.1、av 18.1.0，載入 0.7 秒，3 秒音檔辨識約 5 秒。
- **輸出是偏華語的繁體中文，不是客語漢字**；而且**準確度有限**。實測：
  - 客委會合成語音「恁早啊，細兔仔愛去學校上課。」→「早安早安，小兔子愛去學校上課」（大意對，開頭重複）。
  - 同一句用 united-link 合成 →「早安早安細兔會去學校上課」。
  - 較長或較文的句子錯很多（例：「食飯愛正經食，正毋會食到半出半入」→「食飯愛認真吃炸魚會十到半出半入」）。
  - 模型說明：以中小學客語教材語料微調（國立臺南大學 KWS 中心）。原始未量化模型結果相同，不是轉檔造成。
- 因此分頁 3 的定位改為「**試用版（Beta）**：抓大意用，一定要校對」，介面文字與 README 都要老實寫：
  - 分頁標題：「③ 客語語音變文字（試用）」。
  - 說明（顯眼、不可省）：「在這台電腦上離線辨識，音檔不會上傳。結果是華語為主的繁體中文，**只能抓大意、常會聽錯**，請老師對照原音校對。短句、發音清楚、教材式的句子比較準。」
  - 另放一顆並列的主要按鈕「開啟客委會語音辨識網頁（官方，建議）」→ `https://speech.hakka.gov.tw/AsR/Recognition`（只用 `webbrowser.open`，不帶任何資料）。
  - 原規格的「複製並開啟客委會翻譯（客語→華語）」按鈕改為「複製文字」「另存 .txt」兩顆即可（輸出已是華語，不需要再翻）。
  - 「選擇模型資料夾…」保留：日後換成更準的 CTranslate2 模型（例如講師申請到其他授權模型）不必改程式。
- 辨識參數：`language="zh", task="transcribe", beam_size=5, vad_filter=False, condition_on_previous_text=False`。可以試 `repetition_penalty=1.1`、`no_repeat_ngram_size=3` 是否消除開頭重複（用 `probe.wav` 與同資料夾 `hakkagov.wav` 比較；只有在兩個檔結果都沒有變差時才採用，並在回報中附前後對照）。**不要寫任何針對特定句子的後處理或字串替換。**
- 自動測試的通過條件改為：`--asr probe.wav` 能跑完、輸出非空且含「學校」。不要把「含恁早」當條件。
- 同資料夾還有測試音檔 `t1.wav`、`t2.wav`、`t4.wav`、`hakkagov.wav` 與 `tests.json`（各檔對應的客語漢字）。
- `--shots` 的分頁 3 示範內容用：檔名 `客語語音_示範.mp3`、結果「早安，小兔子愛去學校上課」。

### 1-3 MP3
- 用 PyAV 編碼（libmp3lame）：把各句 WAV 的 PCM 串接（句間插 0.35 秒靜音），輸出單聲道 mp3（128 kbps；取樣率用 WAV 原值，若編碼器不接受就重取樣到 24000 或 44100）。不要引入 ffmpeg 執行檔或 pydub。

## 2. 介面（tkinter／ttk，微軟正黑體，視窗約 980×720，可縮放，三個分頁）

共同：標題列「客語語音小幫手 1.0.0」；最下方一列小字：左「畫面與草稿交給工具，客語定稿交給老師｜AI 生成，教師審核」，右「NDHU 自資系 游豐兆 製作」（`CREDIT` 常數，底部列用 `pack(side="bottom")` 先佔位，避免被擠出視窗）。長工作在背景執行緒跑，介面不可凍結；有進度文字與「取消」。

### 分頁 1「① 華語變客語文字」
- 上：多行輸入框「華語句子」（旁邊顯示字數；>100 字時提示「客委會非會員每次 100 字，請分段」）。
- 三個按鈕（各自：把華語放進剪貼簿 → `webbrowser.open` 官方頁 → 狀態列提示「已複製，請到瀏覽器貼上並按翻譯，複製客語漢字後回來」）：
  - 「複製並開啟：客委會即時翻譯」→ https://speech.hakka.gov.tw/Translation/Online （標「建議」）
  - 「複製並開啟：長問翻譯器」→ https://translate.bronci.com.tw/
  - 「複製並開啟：Gohakka」→ https://gohakka.org/translate （小字：需登入，免費每日 20 字）
- 下：多行輸入框「客語漢字」＋按鈕「貼上客語漢字（從剪貼簿）」「清除」。老師也可以直接打客語漢字。
- 檢查與提醒（不阻擋，只提示）：客語框內容與華語框完全相同 →「看起來還沒翻譯」；客語框是空的 → 停用下一步。
- 右側或下方一個可收合的「三個網站的使用規定」說明框（文字取自第 0 節，條列，含「本工具不會自動連這三個網站」）。
- 記住「翻譯來源」＝最後按的那個按鈕（或「自行輸入」），之後寫進來源說明檔。
- 按鈕「下一步：變成語音 →」切到分頁 2 並帶入客語漢字。

### 分頁 2「② 客語文字變語音 mp3」
- 輸入框「客語漢字」（可直接編輯）。
- 腔調（六個選項按鈕，預設四縣）、語者（下拉，顯示友善標籤，執行時取得；有「重新整理」）、語速（0.5–1.5 滑桿，預設 1，標「越大越慢」）。
- 輸出資料夾（預設 `桌面\客語語音`，可改；`expanduser`）；檔名（預設取客語文字前 12 字，去掉不合法字元，自動加序號避免覆蓋）。
- 按鈕「產生客語語音 mp3」：依句號、問號、驚嘆號、換行切句（每句 ≤60 字，過長再依逗號切）→ 逐句呼叫 API → 串接 → 存 mp3。完成後顯示：斷詞、客語拼音（每句一行）、檔案路徑；按鈕「播放」（`os.startfile`）、「開啟資料夾」。
- 同時寫 `<檔名>_來源說明.txt`（UTF-8）：華語原句（若有）、客語漢字、斷詞、拼音、腔調、語者標籤、語速、翻譯來源、語音來源與授權（臺灣客語語音合成系統 united-link/taiwanese-hakka-tts；模型 yourtts-htia-240704；CC BY-NC 4.0；非商業教學使用）、產生時間、「AI 生成，教師審核；給孩子聽的請由老師親口唸，本音檔供老師備課練習與家長參考」。若翻譯來源是 Gohakka，另加一行「翻譯來源：GoHakka.org（CC BY-SA 4.0）」。
- 提醒文字：「文字會送到公開的語音合成示範服務，請勿輸入幼兒姓名或任何個資；這是共用服務，請勿大量連續使用。」
- 失敗處理：連不上、Space 休眠或回 error → 中文說明＋按鈕「改用客委會語音合成網頁」（剪貼簿＋開 https://speech.hakka.gov.tw/TTS/Synthesis ）。逾時 120 秒。單句失敗重試 1 次。

### 分頁 3「③ 客語語音變文字」
- 「選擇音檔…」（mp3／wav／m4a／aac／ogg／flac）；顯示檔名與長度；>10 分鐘先提示會比較久。
- 「開始辨識」：第一次載入模型時顯示「載入模型中（約十幾秒）」；逐段把文字附加到結果框（faster-whisper 的 segments 是產生器，可邊跑邊顯示），顯示進度（已處理秒數／總長）。
- 結果框（可編輯）＋「複製文字」「另存 .txt」「複製並開啟客委會翻譯（客語→華語）」。
- 說明：「在這台電腦上離線辨識，音檔不會上傳。結果是客語漢字，AI 可能聽錯，請老師校對。」
- 模型不存在時依 1-2 顯示說明。

## 3. 命令列（給自動測試與講師）
- `--version`
- `--selftest`：**離線**項目——切句函式、檔名清理、來源說明檔內容、用合成正弦波做 WAV→MP3 編碼再解碼驗證長度誤差 <5%、介面可建立並立刻關閉（無顯示環境時略過並註記）。全部通過 exit 0；結果寫 `selftest\selftest.log`。**不可依賴網路或模型**（CI 要跑）。
- `--selftest-online`：實際合成一句「恁早。」存 mp3 並驗證長度 >0.5 秒。
- `--tts "客語漢字" [--accent sixian] [--speaker 標籤或值] [--speed 1] --out x.mp3`
- `--asr 音檔 [--model 模型資料夾] [--out x.txt]`（印出辨識文字）
- `--shots 資料夾`：教學截圖模式——以假資料把三個分頁各截一張 PNG（PIL `ImageGrab` 抓視窗範圍；分頁 2、3 用預先填好的示範內容與結果，不連網、不跑模型；輸出資料夾與路徑一律用示意路徑如 `C:\Users\老師\Desktop\客語語音`，**畫面上不可出現真實使用者路徑**）。檔名 `shot1_華語變客語.png`、`shot2_文字變語音.png`、`shot3_語音變文字.png`，另加 `shot2b_產生完成.png`（顯示斷詞、拼音與檔案路徑的完成狀態）。

## 4. 專案檔案（照 repo 既有慣例，參考 `apps\name-masker\`）
`apps\hakka-voice\`：`app.py`（GUI＋CLI）、`core_tts.py`、`core_asr.py`、`core_audio.py`、`build.spec`（onefile、console=False、name=`HakkaVoiceHelper`；Windows 圖示 `icon.ico` 可用 `make_icon.py` 產生，風格比照其他 app；darwin 加 BUNDLE）、`requirements.txt`（requests、faster-whisper、ctranslate2、av、pillow；不要 torch、不要 onnxruntime）、`README_使用說明.md`（給老師看的中文說明：三步驟、要不要網路、模型資料夾放哪、三個網站規定、授權與來源標示、常見問題）、`SPEC.md`（本檔）。
- `.github\workflows\build.yml` 的 matrix 加一列 `{ dir: hakka-voice, name: HakkaVoiceHelper }`。macOS 那一格若因 ctranslate2／PyAV 封裝失敗，先把這個 app 在 macOS 排除（用 matrix exclude），並在回報中說明，不要為了它卡住其他 app。
- repo 根目錄 `README.md` 的 app 清單加一項。
- 先跑 `python tools_pii_guard.py`；原始碼不得出現任何真人姓名（語者、研習教師）與信箱；講師姓名「游豐兆」可出現。

## 5. 建置與部署（本機）
1. `PYTHONUTF8=1 python app.py --selftest` 通過。
2. `pyinstaller --noconfirm build.spec` → `dist\HakkaVoiceHelper.exe`；用精簡 PATH（`C:\Windows\System32;C:\Windows`）跑 `dist\HakkaVoiceHelper.exe --selftest` 通過；再跑 `--selftest-online`；模型若已就緒，跑 `--asr <probe.wav> --model <模型資料夾>` 並確認輸出含「恁早」。
3. 部署到 `D:\01-FC東華\98-軟體\01-自行開發的桌面板軟體\08-文字轉客語語音\`：`HakkaVoiceHelper.exe`、`README_使用說明.md`、`版本與驗證.txt`（版本、SHA-256、建置時間、三項測試結果）、`models\hakka-whisper-ct2-int8\`（模型已就緒才複製；約 0.8 GB）、`範例\`（放一個由 `--tts` 產生的 mp3 與其來源說明檔）。
4. 用 `--shots` 產生四張截圖到 `D:\01-FC東華\07-研習增能\20261114_鳳林國小幼兒園教師增能研習\screenshots\`，檔名前綴 `S10-0N_客語語音小幫手_…png`（開圖目視確認沒有真實路徑與個資）。
5. git：在 `D:\02-aitt_claude_agent\fc-desktop-apps` commit（訊息結尾 `Co-Authored-By: Claude Fable 5.2 <noreply@anthropic.com>`），`git push origin main`（一般 push；**不可 force push、不可碰任何憑證或 token**）。`dist\`、`build\`、模型、截圖不進版控。

## 6. 回報
exe 路徑與大小、三項測試的實際輸出（selftest、online、asr）、截圖檔名、commit hash 與是否已 push、macOS 是否排除、沒做到與需要講師決定的事。
