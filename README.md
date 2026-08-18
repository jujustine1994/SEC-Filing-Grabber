/*  ================================  *\
 *                                    *
 *          C  T  H                   *
 *        created by CTH              *
 *                                    *
\*  ================================  */

規則檔: windows-tool.md
類型: Windows 工具

# SEC Financial Fetcher

股票分析師工具：從 SEC EDGAR 抓取美國上市公司 GAAP 財報並存成 Excel。

## 執行方式

雙擊 `啟動器.bat`

### 指令列（給 skill 用，不經 GUI）

```bash
# GAAP 三表 + 比率 + segment → Excel（與 GUI 產的逐格相同）
./venv/Scripts/python.exe src/cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx

# 8-K 新聞稿的 Non-GAAP 調節表（已解析、已篩過）→ JSON
./venv/Scripts/python.exe src/cli.py press-release ARLO --years 2025-2026 --tables --json
```

兩個子指令都**不呼叫任何 AI API**，只打 SEC EDGAR。共通參數：`--years`
（`2023-2026` 或 `2024`）、`--identity`、`--max-filings`、`--json`（不給路徑
就印到 stdout）、`--lang`（產出 Excel 的顯示語言：`zh_tw` / `zh_cn` / `en` /
`ja`，不給就跟 GUI 用同一個設定；只影響 B 欄與 Index 版面，A 欄機器鍵與 C 欄
公司原文不變）。`gaap` 另有 `--xlsx` / `--quarterly-only` / `--annual-only`，
`press-release` 另有 `--raw`（改吐新聞稿全文，除錯用）。

`press-release` 吐的是**解析後的表格**不是原文：ARLO 一季原文 450K 字元，
篩完 4.4K。

⚠ **季度標籤看 `fiscal_label`，不要看 `label`。** `label` 是用 8-K 的
`period_of_report`（＝發布日）換算的，有系統性 off-by-one（偏 −3 到 +1 季，
見 `docs/8k-period-off-by-one.md`），為了不破壞既有介面而保留原值並帶著
`label_warning`。`fiscal_label` 是從新聞稿表格裡的**期末日**（`period_end`）
加公司財年結束月（`fy_end_month`，payload 頂層）算出來的，與 `Data_Q` 的
財季同一套慣例，兩邊對得起來。抓不到期末日時 `fiscal_label` 留空，不會用
發布日硬算。15 家 120 份實測全部抓得到。

⚠ **`--years` 篩的是發布日不是財期**：篩選發生在下載之前，那時還讀不到
期末日。非 12 月結算的公司在年份邊界可能差到 3 季，要精確就把範圍放寬一年，
再自己用 `fiscal_label` 篩。

### 要傳給別人時

雙擊 `scripts\打包.bat`，產出 `dist\SEC-Financial-Fetcher-YYYYMMDD.zip`（約 190 KB），
跑完 12 項自我驗證才會留下檔案。收件人解壓、雙擊 `啟動器.bat`、一路按 Enter 就能裝好
（沒裝過 Python 也可以，uv 會自己下載）。細節見 `docs/PACKAGING.md`。

## 系統需求

- Windows 10/11
- 需要網路連線（首次安裝 + 每次抓取資料）

## 技術棧

- 語言：Python 3.13（`uv venv` 自動下載，不依賴系統 Python）
- 核心套件：`edgartools`（SEC EDGAR / XBRL 抓取）、`openpyxl`（Excel 讀寫）、tkinter（GUI）
- Non-GAAP 解析可選用 `google-generativeai` / `openai` / `anthropic`（僅該功能需要 API Key）

## 首次設定

1. 雙擊 `啟動器.bat`，按照提示完成套件安裝
2. **第一次啟動會跳一個 `Language` 視窗**，四個按鈕：繁體中文／简体中文／English／日本語。點一下就記住，之後不會再問（直接關掉視窗＝用繁體中文，一樣不再問）
3. 程式啟動後點「進階設定」，填入 SEC EDGAR Identity（姓名 + 信箱）
4. 若要使用 Non-GAAP 功能，在進階設定填入 AI API Key

之後要換語言：「進階設定」最上方的 `Language` 選單。選完會跳一個英文視窗問要不要重啟，按 Restart 就直接換好。

## Excel 結構

每間公司一個 `.xlsx`，存於 `output/` 資料夾。

| Sheet | 說明 |
|-------|------|
| `Data_Financials(Q)` | **季報三表**（IS + BS + CF，from 10-Q）。表頭 3 列為期間標籤，三表各有專屬底色，公司特有科目集中在底部 `Other (as reported)` |
| `Data_Financials(Y)` | **年報三表**（from 10-K），結構同上 |
| `Data_Ratios` | 37 個常見比率（Python 計算，**零 AI**）。A 欄英文列名（含 `(%)` / `(x)` / `(days)` / `($)` 單位後綴）、B 欄說明、C 欄算法 |
| `Data_Segments` | 營收／費用分類細項，長格式（各軸合併於一張） |
| `Data_Meta` | 申報資訊（Ticker、公司名、抓取日期、季數、財年結束月、**抓取缺漏**） |
| `Index` | 第一頁：公司抬頭、**抓取缺漏警告**、**財年起始月輸入格**、sheet 清單、品質明細 |

欄位配置（A 欄機器鍵永遠英文、B 欄跟著介面語言、C 欄永遠公司原文）、固定列位、
跨公司模板公式怎麼寫，完整規格見 `docs/ARCHITECTURE.md`「Excel Sheet Layout」。

## 財年起始月：程式猜錯時自己改

財年結束月是程式從 10-K 自動判讀的，**會出錯**。`Index!B4`（黃底那格）可以改，
`Data_Financials(Q)/(Y)` 的期間標籤會**全部自動更新**。核對方式與細節見
`docs/ARCHITECTURE.md`。

## 抓不到資料時會怎樣

某幾期沒抓到時**那幾期留空、其餘照常產出**，並主動告訴你缺了哪幾期：

> ⚠ 有 2 期沒抓到（2025-09-27、2025-03-29）。抓取期間連不上 SEC，多半是網路問題——網路穩定後重抓一次通常就補得回來。

三個地方看得到：GUI 橘字、`logs/app.log`、**Excel 第一頁 Index 的橘底那列**。
最後一個最重要——GUI 關掉就沒了，但三天後重開這份 Excel 警告還在。

網路閃斷會自動退避重試（2/4/8 秒）多半救得回來；**一期都沒抓到才不寫檔**，
你原本的 Excel 完好不動。

## 已知限制

見 `docs/ARCHITECTURE.md`「Known Issues」；待決定要不要修的項目見 `docs/TODO.md`。

## .gitignore 必含項目

- `config.json`、`.env`（機敏設定，內含 API Key / SEC Identity）
- `venv/`、`__pycache__/`、`*.pyc`
- `output/`（保留 `.gitkeep`）、`dist/`、`*.zip`
- `*.log`、`logs/`
- `company_cache.json`（執行期自建快取）
