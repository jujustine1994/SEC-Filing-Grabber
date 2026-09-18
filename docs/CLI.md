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

⚠ **公司改過財年的舊申報會整段標錯**：EDGAR 的 `fiscal_year_end` 只有現值。
**201 家實測只有 2 家改過財年（1.0%）**，但中招的話中得徹底——LHX 2019 年從
6 月底改成 12 月底，改制前那 30 季**全部**差 2 季；MSCI 那次只挪 31 天，不跨季所以
沒事（`scripts/check_fye_drift.py`，離線可重驗）。`label_agrees_with_fiscal_label`
抓得到「選進來的那幾份有問題」，抓不到「該選進來卻被 `--years` 漏掉」的那一類。

## `compare`：跨公司比較（AI 要畫圖就用這條）

```bash
./venv/Scripts/python.exe src/cli.py compare ARLO FORM NVDA --json out.json
./venv/Scripts/python.exe src/cli.py compare AAPL MSFT --metrics Revenue "Gross Margin (%)" --years 2023-2026
./venv/Scripts/python.exe src/cli.py compare AAPL MSFT --annual          # 比年報
./venv/Scripts/python.exe src/cli.py compare --list-metrics              # 95 個科目 + 59 個比率
```

**為什麼不要自己跑四次 `gaap` 再合併**：各家財年不同，同一個「Q2」不是同一段
時間（ARLO 的 Q2 結束在 06-28、NVDA 在 07-27）。`comparison.py` 已經把日曆季
對齊、財季對照、缺季留白這些做完了，這條就是把它接出來。自己合併等於重造
輪子，而且**對錯了不會報錯**。

輸出的 JSON：

- `periods`——對齊過的**日曆季**清單，這是唯一的期間事實來源
- `data`：`{指標: {ticker: {日曆季: 值}}}`。每個 dict 的鍵**剛好**等於 `periods`
- `fiscal_labels`：`{ticker: {日曆季: 那家自己的財季標籤}}`。各家財年不同時
  的對帳依據——同一欄 NVDA 是 FY2026Q2、AMD 是 FY2025Q2
- `period_ends`：每家每一欄的實際期末日
- `synthetic_q4`：**推算出來的 Q4**。SEC 沒有 Q4 的 10-Q，季報表裡的 Q4 一律是
  「年報 − Q1 − Q2 − Q3」算的，要標出來
- `failures`：哪幾家整家抓失敗（回傳碼 1）

⚠ 指標名是**英文機器鍵**，跟 Excel A 欄同一套。打錯會當場擋下（回傳碼 2），
不會讓你抓了十分鐘才發現整欄是空的。比率的值是百分比或倍數本身
（`Gross Margin (%)` = 44.28），**不是金額**，不要再除以 1e6。

## `db-status`：資料庫裡實際有哪些公司（TODO J7）

**AI／腳本要查資料庫內容就用這條。不連網，244 家實測 0.53 秒。**

```bash
./venv/Scripts/python.exe src/cli.py db-status                 # 人看的表格
./venv/Scripts/python.exe src/cli.py db-status --json          # 給 AI 吃
./venv/Scripts/python.exe src/cli.py db-status NVDA AMD        # 只看這幾家
./venv/Scripts/python.exe src/cli.py db-status --rebuild       # 先重算過期的 meta（很慢，見下）
```

⚠ **跟另外兩個東西分清楚，三者答的是不同問題**：

| 要問的 | 用哪個 | 連網 |
|---|---|---|
| 資料庫**實際有**哪些公司 | `db-status` | 否 |
| 名單上**要抓**哪些公司 | `update-db --list` | 否 |
| 哪幾家**不完整、該補** | `scripts/audit_local_db.py` | **是**（每家 2 次請求） |

`update-db --list` 印的是 `config["local_db_tickers"]`，那是「要保持新鮮的清單」，
跟「快取裡真的有什麼」是兩個集合，可以完全不同。

`--json` 每家給這些欄位（**英文機器鍵**）：

- `period_from` / `period_to`——**財報期間**（手上有哪幾季的數字）
- `filed_from` / `filed_to`——**SEC 收件日**，跟上面差一整個期間，不要混用
- `years`、`filings`、`size_bytes`
- `bottom`（`yes`／`no`／`unknown`）、`bottom_stale`
- `updated_at`、`stale_days`——**上次去 SEC 查這家**是多久以前。一家「已到底、
  最新期間 2026-03」如果是 120 天前查的，中間很可能已經出新財報
- `in_list`（在不在更新名單）、`meta_ok`（meta 快照跟目錄對不對得上）

**唯讀**：這條走 `read_meta()`，不會寫任何檔。meta 對不上就照實回報
`meta_ok: false`（顯示「需重算」），不當場自癒。

⚠ **`--rebuild` 很慢**：`scan_filings()` 為了讀幾個小欄位把每份 70KB 的 filing
JSON 整個 `json.load()` 進來，實測 **2.75 秒/家、244 家約 11 分鐘**。多數情況
不必跑——下一次 `update-db` 會順便把 meta 修好。

## `update-db`：更新本地財報資料庫（TODO J3）

```bash
# 名單維護（做完就存檔結束，不會順便發動抓取）
./venv/Scripts/python.exe src/cli.py update-db --list
./venv/Scripts/python.exe src/cli.py update-db --import-cached      # 快取裡已有的全加進來
./venv/Scripts/python.exe src/cli.py update-db --import-watchlist   # watchlist 全加進來
./venv/Scripts/python.exe src/cli.py update-db --add NVDA MSFT
./venv/Scripts/python.exe src/cli.py update-db --remove NVDA

# 真的跑（走更新名單，一律拓到底，只暖快取不產 Excel）
./venv/Scripts/python.exe src/cli.py update-db --json out.json
./venv/Scripts/python.exe src/cli.py update-db AAPL META   # 只跑這幾家，不動名單
```

- **沒有 `--years` / `--max-filings`**：深度是固定的（拓到底），這是設計決定
- **已經到底又沒有新財報的公司整家跳過**，只花一次 filing 清單查詢。
  實測三家全跳過是 1.0s；要抓的公司約 1.8 秒/份
- **單一公司失敗不中斷整體**，最後列出失敗與「有抓取缺漏」的公司（TODO D11：
  連續大量抓取時 SEC 會偶發失敗、靜默少格）。那幾家之後單獨重跑即可
- **中斷不會白費**：逐份即時落檔，重跑會自動跳過已完成的
- 離開碼：0 全部成功／1 有公司失敗／2 參數或設定有問題
- 這條是給 **Windows 工作排程器**掛半夜跑用的——GUI 開著過夜不可靠（更新、休眠）
