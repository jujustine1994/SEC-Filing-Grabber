# 交接：打包散布（給下一個對話）

> 用法：把下面「交接 Prompt」整段貼進新對話。這份檔案本身可以在任務完成後刪掉。

---

## 交接 Prompt

我要做 `docs/TODO.md` 的 **F 段：打包散布給非技術使用者**。專案在
`C:\Users\CTH\Documents\Code\SEC Financial Tools`。

**先讀這三份，不要靠猜：**
1. `README.md`（開頭有 `規則檔: windows-tool.md`，照規則去讀
   `C:\Users\CTH\.claude\project-rules\windows-tool.md`）
2. `docs/TODO.md` 的 **F 段** — F1/F2/F3 三項，裡面已經記了上一輪查證過的事實，
   標「已查明」的**不用重查**
3. `launcher.ps1` 全文（F2 的三個卡點都在裡面，有標行號）

### 要做什麼

**目標**：GitHub 連不上，我要壓一包 zip 直接傳給朋友（他不太會用電腦），
對方解壓、雙擊、一路按 Enter 就能裝好並跑起來。

分兩件事，**順序不能顛倒**：

**第一件：修 `launcher.ps1`，讓「一路按 Enter」真的過得去**（TODO F2）

三個卡點上一輪已經逐行審過，行號在 TODO 裡：
- `launcher.ps1:73` 的 `winget install --id Python.Python.3` 這個 package ID
  **可能已經失效**（現在通常是 `Python.Python.3.12` / `3.13`）。失效的話整段
  自動安裝等於沒用 —— 這是最該先實測的一項，跑
  `winget show --id Python.Python.3` 就知道
- `launcher.ps1:84-87`：winget 裝完 Python 後 PATH 沒刷新，程式叫使用者關掉
  再雙擊一次。不是「一路 Enter」但訊息夠清楚，可接受，說明書要先講
- `launcher.ps1:74-78`：沒有 winget（較舊 Win10）就叫使用者自己去 python.org。
  不會用電腦的人到這裡就停了。要不要處理**先問我朋友的 Windows 版本**

⚠ **Mark-of-the-Web 我已經決定不處理**，別再提。zip 經通訊軟體傳送、解壓後
`.bat` 被 SmartScreen 擋是收件人自己的事，程式端不加 `Unblock-File`。

**第二件：寫 `docs/PACKAGING.md` 並照它執行一次打包**（TODO F1）

那份說明書是**給 AI 照著做的作業指示，不是給人看的教學**。要能讓下一個 AI
不必重新調查就完成打包並自我檢查。至少要有：排除清單、包含清單、產出的 zip
檔名規則、打包後的自我驗證步驟（例如解到暫存目錄確認沒有機敏檔案、確認
`src/` 完整）。

順便寫一份**給收件人看的**簡短說明（放進 zip 裡），內容只要 TODO F3 那兩件事：
填 SEC EDGAR Identity（沒填抓不了任何資料）、首次啟動會跳 Language 視窗。
**不要提 AI API Key** —— `main.NONGAAP_ENABLED = False`，Non-GAAP 功能停用中。

### 關鍵事實（上一輪查證過，直接用）

- ✅ **機敏資訊天生不在專案資料夾內**：`src/config.py:14-18` 把 `config.json`
  放在 `%APPDATA%\SEC Financial Tools\`。API Key、SEC Identity、Watchlist、
  個人輸出路徑全在專案外，壓 zip 碰不到。專案內只有 `config.example.json`
  （假的範例值）
- **要手動排除、git 沒擋的**：`company_cache.json`（414KB，程式會自己重建）、
  `output/_final/*.xlsx`（14 個我的實測輸出檔）、`20260814 sec tool.zip`、
  `.git/`（7.4MB）、`.pytest_cache/`、`.claude/`、`.superpowers/`
- **不必打包**：`venv/`（400MB）、pip 套件 —— `launcher.ps1` 會自己建
- **一定要打包**：`啟動器.bat`、`launcher.ps1`、`README.md`、`requirements.txt`、
  `src/`、`config.example.json`

### 專案現況

- 分支 `master`，commit `8764a6b`，工作區乾淨
- 測試：`./venv/Scripts/python.exe -m pytest -q -m "not slow"` → **841 passed**
- 本機領先 `origin/master` 72 個 commit，**沒 push**（GitHub 連不上，先擱著）
- 上一輪剛做完五件事（視窗置中、掃描鍵改名、同名檔提醒、進階設定改頁籤、
  斷網缺漏回報），細節在 `docs/CHANGELOG.md` 與 `docs/ARCHITECTURE.md`

### 我的偏好

- 繁體中文，直接說重點，不要鋪墊
- 有多種做法時**列選項讓我選**，不要自己決定
- 改現有檔案只做差異編輯，不整份重寫
- **不可刪除 `scripts/` 下的任何腳本**；冗餘的在 `scripts/README.md` 標
  `(停用)` 但保留檔案本體
- 改 `launcher.ps1` 之後不可以只看程式碼就說完成。含互動安裝的流程要我
  雙擊 BAT 實測，我回報看到預期畫面才算過（規則檔有寫）

### 兩件上一輪沒驗完的事（跟打包無關，但你可能會看到）

- GUI 沒實機看過：版面數字都是用探針量的，emoji 按鈕實際長相沒確認
- 斷網沒真的拔線試過：邏輯有 45 條單元測試蓋著，端到端用假 filing 驗過

---

## 給我自己的備忘

上一輪對話做完的 commit（master `8764a6b` 為止）：

| commit | 內容 |
|---|---|
| `d9d8de3` | TODO F 段：打包散布（就是這次要做的） |
| `6780e83` | 掃描鍵／執行鍵分開 + 分段提示 |
| `fbf2c70` | 視窗置中 900×720，不再跳動 |
| `ed03276` → `65f41df` | 同名檔提醒（做了又簡化） |
| `936a05c` | 進階設定改第三個頁籤 |
| `9944e6a` → `4790840` | 斷網缺漏回報（改了兩版） |
| `8764a6b` | 文件同步 |

舊分支還留著沒刪：`feat/8k-scan-optimization`、`fix/nongaap-data-quality`、
`feat/ui-and-resilience` —— 內容都已在 master 裡。
