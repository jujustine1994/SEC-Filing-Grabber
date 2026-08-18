# CLI 使用說明（給 skill / AI 呼叫用）

不經 GUI，直接呼叫指令列取得資料。給外部 skill 或自動化流程用；一般使用者
不需要看這份，雙擊 `啟動器.bat` 就好。

```bash
# GAAP 三表 + 比率 + segment → Excel（與 GUI 產的逐格相同）
./venv/Scripts/python.exe src/cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx

# 8-K 新聞稿的 Non-GAAP 調節表（已解析、已篩過）→ JSON
./venv/Scripts/python.exe src/cli.py press-release ARLO --years 2025-2026 --tables --json
```

兩個子指令都**不呼叫任何 AI API**，只打 SEC EDGAR。

## 共通參數

| 參數 | 說明 |
|---|---|
| `--years` | `2023-2026` 或單一年份 `2024` |
| `--identity` | SEC EDGAR Identity，不給就讀 `config.json` |
| `--max-filings` | 最多抓幾筆 filing |
| `--json` | 輸出 JSON；不給路徑就印到 stdout |
| `--lang` | 產出 Excel 的顯示語言：`zh_tw` / `zh_cn` / `en` / `ja`。不給就跟 GUI 用同一個設定；只影響 B 欄與 Index 版面，A 欄機器鍵與 C 欄公司原文不變 |

`gaap` 另有 `--xlsx` / `--quarterly-only` / `--annual-only`；`press-release`
另有 `--raw`（改吐新聞稿全文，除錯用）。

`press-release` 吐的是**解析後的表格**不是原文：ARLO 一季原文 450K 字元，
篩完 4.4K。

## 兩個必讀警告

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
