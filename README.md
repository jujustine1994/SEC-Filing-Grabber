/*  ================================  *\
 *                                    *
 *          C  T  H                   *
 *        created by CTH              *
 *                                    *
\*  ================================  */

規則檔: windows-tool.md
類型: Windows 工具

# SEC Financial Fetcher

股票分析師工具：輸入美股代號，從 SEC EDGAR 自動抓公開財報，整理成 Excel。

## 怎麼用

雙擊 `啟動器.bat`，第一次執行會自動安裝需要的套件（約 3-5 分鐘，需要網路）。

首次啟動：

1. 跳出的 `Language` 視窗選一種顯示語言（繁中／简中／English／日本語），選完不會再問
2. 點「進階設定」，填入 SEC EDGAR Identity（你的名字 + email——SEC 規定任何自動抓取程式都要自報身分，這組字只會送給 SEC）
3. 要換語言、填 AI API Key（Non-GAAP 功能才需要），也是在「進階設定」

之後輸入股票代號、按執行，Excel 檔會出現在 `output/` 資料夾。

## 系統需求

- Windows 10/11
- 網路連線（安裝套件 + 每次抓取資料時）

## Excel 長什麼樣子

每間公司一個 `.xlsx`：

| Sheet | 內容 |
|-------|------|
| `Data_Financials(Q)` / `(Y)` | 季報 / 年報三表（損益表、資產負債表、現金流量表） |
| `Data_Ratios` | 37 個常見財務比率，Python 算好，不靠 AI |
| `Data_Segments` | 營收／費用分類細項 |
| `Data_Meta` | 抓取日期、財年結束月、有沒有缺資料 |
| `Index` | 第一頁總覽：公司抬頭、缺漏警告、可自行修正的財年起始月欄位 |

**財年結束月猜錯了怎麼辦**：程式會自動判讀，但偶爾會錯。`Index` 第一頁的黃底
欄位可以直接改，改完整本 Excel 的期間標籤會自動跟著更新。

**抓不到資料時**：那幾期會留空、其餘照常產出，並且會主動講清楚缺了哪幾期
（GUI 橘字提示 + Excel 第一頁橘底那列都看得到）。網路只是暫時不穩的話，重抓
一次通常就補得回來；真的一期都沒抓到，原本的 Excel 檔會維持不動，不會被
空白蓋掉。

## 要傳給其他人用

`git clone` 這個 repo，對方雙擊 `啟動器.bat`、一路按 Enter 就能裝好，不需要
先裝 Python。

> 舊的 zip 打包方式（`scripts\打包.bat`）**暫停用**——當初是因為對方連不上
> GitHub才改用 zip 傳送，現在恢復正常，改走 clone 就好。腳本保留著，之後又
> 連不上再撿回來用，細節見 `docs/PACKAGING.md`。

## 給開發者 / skill 呼叫

- 架構、Excel 欄位規格、已知限制：`docs/ARCHITECTURE.md`
- 指令列介面（給外部 skill 用，不經 GUI）：`docs/CLI.md`
- 待辦事項：`docs/TODO.md`
- 打包發布流程：`docs/PACKAGING.md`
- 技術棧：Python 3.13（`uv` 自動管理，不依賴系統 Python）、`edgartools`、`openpyxl`、tkinter
