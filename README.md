規則檔: windows-tool.md

# SEC Financial Fetcher

股票分析師工具：從 SEC EDGAR 抓取美國上市公司 GAAP 財報並存成 Excel。

## 執行方式

雙擊 `啟動器.bat`

## 系統需求

- Windows 10/11
- 需要網路連線（首次安裝 + 每次抓取資料）

## 首次設定

1. 雙擊 `啟動器.bat`，按照提示完成套件安裝
2. 程式啟動後點「進階設定」，填入 SEC EDGAR Identity（姓名 + 信箱）
3. 若要使用 Non-GAAP 功能，在進階設定填入 AI API Key

## Excel 結構

每間公司一個 `.xlsx`，存於 `output/` 資料夾。
- `Data_IS`：損益表
- `Data_BS`：資產負債表
- `Data_CF`：現金流量表
- `Data_Equity`：股東權益變動表
- `Data_CI`：綜合損益表
- `Data_Meta`：申報資訊

欄位說明：A 欄 = 指標名稱，B 欄起 = 各季數據（舊→新），第 1 列 = 季度標籤，第 2 列 = 申報日期。

分析用的自訂 Sheet 請命名為 `My_*`（如 `My_IS`），Python 不會碰這些 Sheet。
