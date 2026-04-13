# SEC Financial Fetcher — Architecture

## File Map

| File | Role |
|------|------|
| 啟動器.bat | 薄 BAT，呼叫 launcher.ps1 |
| launcher.ps1 | 環境檢查、uv venv、安裝套件、啟動 main.py |
| main.py | Tkinter GUI，兩個 tab + 兩個 popup |
| config.py | load_config() / save_config() |
| fetcher_gaap.py | edgartools XBRL 抓取 → StatementTable 列表 |
| fetcher_nongaap.py | Non-GAAP stub（Phase 2） |
| excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx |
| config.json | 使用者設定（gitignored） |
| config.example.json | 範本（committed） |
| company_cache.json | Ticker → 公司名快取（committed） |
| output/ | 輸出的 Excel 檔（gitignored） |

## Data Flow

1. 使用者輸入 Ticker（Tab 1）或從 Watchlist 選取（Tab 2）
2. fetcher_gaap.py 抓取全部歷史季度財報
3. 回傳 list[StatementTable]（每種財報一個）
4. excel_writer.py 開啟/建立 output/TICKER.xlsx
5. 全量改寫所有 Data_* sheets，不碰 My_* 等其他 sheets

## Key Config Variables (config.json)

- `identity`: SEC EDGAR 身份字串（必填，格式：名字 空格 信箱）
- `output_dir`: Excel 輸出路徑（預設 "output"）
- `watchlist`: [{ticker, name}, ...] 清單
- `ai.provider`: "google" / "openai" / "anthropic"
- `ai.model`: 模型名稱（預設 "gemini-flash-latest"）
- `ai.api_key`: API Key（gitignored）

## Excel Sheet Layout

每個 Data_* sheet：
- Col A：指標名稱（Concept/Label）
- Col B+：每個季度一欄（舊→新）
- Row 1：季度標籤（如 FY2024Q1）
- Row 2：申報日期（如 2024-02-01）
- Row 3+：財務數據

## StatementTable (fetcher_gaap.py 的輸出合約)

```python
@dataclass
class StatementTable:
    sheet_name: str          # "Data_IS", "Data_BS", ...
    quarter_labels: list[str]  # Row 1
    filing_dates: list[str]    # Row 2
    concepts: list[str]        # Col A, Row 3+
    values: list[list]         # values[concept_idx][quarter_idx]
```
