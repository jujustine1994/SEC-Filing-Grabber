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

## 季度標籤（2026-08-25 改過，B5）

`press-release` 每一季吐兩個標籤，**兩個現在都是對的**，慣例也相同：

| 欄位 | 怎麼算 | 什麼時候用 |
|---|---|---|
| `label` | 發布日（`period_of_report`）+ EDGAR `fiscal_year_end` 回推名目季末 | **列清單／`--years` 篩選**用的就是它。不必下載文件 |
| `fiscal_label` | 下載後從新聞稿表格抓到的**期末日**（`period_end`）+ 財年結束月 | 最準的基準。抓不到期末日時留空，**不會**用發布日硬算 |

- `label_source`：`"announcement+fiscal_year_end"`（零下載規則）或
  `"period_of_report"`（退回舊算法的那幾季，EDGAR 給不出 `fiscal_year_end`
  或日期畸形時逐份發生）
- `label_warning`：跟著 `label_source` 走，兩種來源帶不同的說明
- `label_agrees_with_fiscal_label`：`true` / `false` / `null`（沒有
  `fiscal_label` 可比時）。**`false` 值得注意**——最可能的成因是公司改過財年，
  EDGAR 只給「現在」的 `fiscal_year_end`
- payload 頂層新增 `fiscal_year_end`（MMDD 原字串，如 `"0703"`），`fy_end_month`
  照舊保留

零下載規則 200 份實測、157 份基準可信全部與 `fiscal_label` 一致（100%）；
改之前 `label` 有 **31.5% 連年份都是錯的**。細節見
`docs/8k-period-off-by-one.md`「零下載規則」一節。

⚠ **`--years` 篩的仍然是發布日所屬的財季，不是財期本身**：篩選發生在下載
之前，那時讀不到真實期末日。現在標的財季已經對了，但發布日跨到下一個財年時
（例如財年結束後才發的年報季）邊界仍可能差一份，要精確就把範圍放寬一年，
再自己用 `fiscal_label` 篩。

⚠ **公司改過財年的舊申報可能整段標錯**：EDGAR 的 `fiscal_year_end` 只有現值。
`label_agrees_with_fiscal_label` 抓得到「選進來的那幾份有問題」，抓不到
「該選進來卻被 `--years` 漏掉」的那一類。
