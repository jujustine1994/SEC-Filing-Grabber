# 設計：本地財報資料庫（filing 快取升級）

> 2026-09-04。CTH 指定方向：**本地資料庫「抓過不用重抓」是本專案的核心能力**。
> 本文是設計，不是實作計畫。實作計畫另外寫。

## 一、要解決什麼

CTH 的五個需求（原話）：

1. 不要每次都全部重抓
2. 新增年度可以的話只抓新的那年度
3. 容量不要太大，**維持 DataFrame 就好**
4. edgartools 的影響不應該過於頻繁
5. 讓使用者考慮是否要升級 edgartools，並**明示需要重抓**這件事。重抓可以半夜跑

## 二、結論先講：不需要重新構思架構

現有的 `filing_cache.py` 已經是正確的形狀。五個需求裡有兩個**現在就成立**：

- 「不要每次全部重抓」→ 已成立，快取的鍵是每一份 filing（`<accession>.json`）
- 「只抓新的那年度」→ 已成立且有實測：2026-09-04 抓 AAPL 時，本地原有 21 份，
  改用更深的抓取窗後變成 75 份——**保留舊的 21 份，只下載新增的 54 份**

缺的不是架構，是三塊「狀態與體驗」：更新名單、涵蓋狀態記錄、版本升級的告知。

## 三、不變的部分

```
SEC → edgartools 解析 → ★快取（只存 DataFrame）★ → 我們的比對層 → Excel／比較表
                                                    ↑ 永遠即時重算，不快取
```

- **儲存內容不變**：只存解析後的三張 DataFrame，位置不變
  （`%APPDATA%\SEC Financial Tools\filing_cache\`）
- **容量**：實測每份 filing 平均 0.085 MB。201 家拓到底約 **1.4 GB**
- **改模板仍然免費**：比對層不快取，hint／模板改動立刻在既有資料上生效。
  2026-09-04 的 H1 驗收（CF 填滿率 25%→100%）就是靠這個結構，才能不重抓就算出來
- **不存原始 XBRL**。實測可行（見附錄 A）但代價是磁碟 1.4 GB → 42 GB，
  CTH 明確否決

## 四、三層模型

原本想把「更新名單」併進既有的 `watchlist`，CTH 指出應該分開，**這個判斷是對的**：

| | 快取內容 | 更新名單（新增） | watchlist（既有） |
|---|---|---|---|
| 怎麼來的 | 掃目錄得出的**事實** | 使用者維護的**意圖** | 使用者維護 |
| 內容 | 所有抓過的（含 Tab 1 一次性查詢） | 要保持新鮮的（可含還沒抓過的） | 批次產 Excel 的對象 |
| 誰改 | 抓取／清除自動改 | 只有使用者 | 只有使用者 |

合併會壞在兩處：併進 `watchlist` → Tab 2 一按產 201 份 Excel；併進「快取裡有的」
→「未來把全部抓好」做不到，因為還沒抓的公司永遠不會被抓。

三者可重疊也可不重疊，**不強制包含關係**。

**落地位置**：更新名單放 `config.json`，跟 `watchlist` 並列（鍵名 `local_db_tickers`）。
**不放快取目錄**——`filing_cache.py` 開頭的明文原則是「事實來源是檔案本身，
不維護額外索引檔」，塞一份手動維護的名單進去會破壞這條。

**配兩個便利動作**，否則維護兩份名單很煩：
1. 「把 watchlist 全部加入更新名單」
2. 「把快取裡已有的全部加入更新名單」

## 五、`_meta.json`

一家一份，放該公司資料夾（`filing_cache/AAPL/_meta.json`）。不做全域單一檔——
分開才不會多視窗同時寫時互相蓋掉。

```json
{
  "schema_version": 1,
  "ticker": "AAPL",
  "cik": 320193,
  "file_count": 75,
  "updated_at": "2026-09-04T12:30:00+08:00",
  "edgartools_version": "5.29.0",
  "forms": {
    "10-Q": {"count": 60, "oldest": "2008-02-01", "newest": "2026-07-31",
             "reached_bottom": "xbrl_cutoff"},
    "10-K": {"count": 15, "oldest": "2008-11-05", "newest": "2025-10-31",
             "reached_bottom": "xbrl_cutoff"}
  }
}
```

**為什麼分 form**：`max_filings`（10-Q）與 `max_annual_filings`（10-K）是兩個獨立
上限，一家公司可能 10-K 到底了、10-Q 還沒。合記會誤判。

### 「到底」怎麼判定

抓取迴圈有三個停止條件（`fetcher_gaap.py:1209-1213`）：撞到 `max_filings`、
撞到 `_XBRL_CUTOFF`（2008-01-01）、清單用完。

**直覺做法是讓 builder 回報是哪一個停的——但那要穿過 3 個 builder 與 8 個
`_current_q_col()` 呼叫點，就是 TODO G13 (a) 案那個坑。改成在 builder 外面推導，
零改動：**

```
available = _list_filings(公司, form) 裡 filing_date >= 2008-01-01 的
cached    = 該資料夾裡該 form 的 accession 集合

reached_bottom = "no_more_filings"  cached ⊇ available，且原始清單沒有 2008 前的
               = "xbrl_cutoff"      cached ⊇ available，且原始清單有被 2008 擋掉的
               = null               否則（還沒抓完，下次要繼續挖）
```

`_list_filings()` 本來就回傳完整清單，所以這些資訊在「更新本地庫」跑的時候順手
就有，不必額外連網、不必動 `fetcher_gaap` 一行。

實測對照（2026-09-04 本地快取）：

| ticker | 最舊申報 | 判定 |
|---|---|---|
| AAPL / AMZN / JPM / NVDA / MSFT | 2008-0x | `xbrl_cutoff`（撞到 XBRL 起點） |
| META | 2013-02 | `no_more_filings`（2012 才上市） |
| ARLO | 2018-08 | `no_more_filings`（2018 IPO） |
| AVGO / AXP / AZO | 2021-06 | `null`（只抓了 5 年，還要挖） |

**順帶釐清一個事實**：「20 年」實際上拿不到。XBRL 從 2008 才開始，到 2026 最多
**18 年**，而且會隨時間相對變短。正確說法是「抓到底」，不是「抓 20 年」。

### 不同步怎麼自癒

原則：**掃目錄為準，meta 只是快照。**

```
讀 meta → 比對 file_count 與實際 len(dir.glob("*.json"))
  相符 → 直接用（GUI 列 201 家不必讀 881 個檔）
  不符 → 重建：讀該資料夾的 filing 算 count/oldest/newest；
         reached_bottom 保留舊值並標記過期，下次更新時重算
meta 不存在／壞掉／schema 不符 → 同上，重建
```

比對用的是**目錄列舉**（不讀檔內容），很便宜。這點很重要——201 家若每次都要讀
881 個 JSON 才能顯示清單，GUI 會卡住。

`reached_bottom` 為什麼過期了還保留：重算它要連網拿完整 filing 清單，不該為了
顯示一個列表就連 201 次網。

## 六、「更新本地庫」的行為

**對象**：更新名單（`config["local_db_tickers"]`）。
**深度**：一律拓到底（CTH 決定）。
**不產 Excel**——只暖快取。要 Excel 走既有的 Tab 2 批次。

### 每家公司的流程

```
1. _list_filings(ticker, "10-Q") / ("10-K")     ← 一次網路，很便宜
2. 跟 meta 比對：
     沒有新 filing，且兩個 form 都 reached_bottom  → 整家跳過，不進抓取
     否則                                          → 進 3
3. fetch_gaap_statements(ticker, max_filings=200, max_annual_filings=50)
   結果直接丟棄——要的是它的副作用（把快取填滿）。
   200/50 只是「大到不會是它先喊停」的餘裕值：XBRL 從 2008 起算最多 18 年，
   ≈72 份 10-Q ＋ 18 份 10-K。實際由 _XBRL_CUTOFF 或清單用完停止
4. 用步驟 1 的完整清單 + 資料夾現況重算 _meta.json
```

步驟 2 是「不要每次全部重抓」的具體實現：**第二次以後，已經到底又沒有新財報的
公司完全不會進抓取迴圈**，只花一次 filing 清單的網路。

### 失敗處理

- **單一公司失敗不中斷整體**，記錄後繼續下一家。比照 `comparison.py` 的
  `CompanyFetchError` 原則（公司層級跳過，跟同一家公司內部的科目缺漏是兩回事）
- 沿用 `collect_gaps()`，跑完列出「有抓取缺漏」的公司清單。這對應 TODO D11：
  連續大量抓取時 SEC 會偶發失敗、**靜默少格**。這些公司之後單獨重跑即可，
  第二輪會從本地快取讀已經成功的部分，只重抓失敗那幾份
- **中斷不會白費**：`save_filing()` 是逐份即時落檔的，關視窗或斷線時已抓到的
  進度都在

### 怎麼半夜跑

兩個入口：
- **GUI 按鈕**（在既有的快取面板上），適合「按下去去睡覺」
- **CLI 子命令**，適合掛 Windows 工作排程器

CLI 那條是必要的——GUI 開著過夜不可靠（更新、休眠）。

### 速率

TODO D11 的候選修法 (c)「降低連續抓取的速率」尚未做。這裡是長時間批次，
是最該套用的地方，但**不在本設計範圍**，只記一筆：若實測 201 家連續跑的缺漏率
偏高，先回頭處理 D11 (c)。

## 七、edgartools 版本

### 鎖版本

`requirements.txt` 現在是 `edgartools>=2.0.0`，**沒鎖**。任何人重跑一次
`pip install -r requirements.txt` 就可能裝到新版，**在沒打算升級的情況下讓整個
本地庫失效**。改成 `edgartools==5.29.0`。

這一條就解決了需求 4「edgar 的影響不應該過於頻繁」——不鎖的話它是隨機發生的。

### 為什麼版本一變就全滅

`load_filing()`（`filing_cache.py:206`）拿存檔時記的 `edgartools_version` 跟現在
安裝的版本做**字串完全比對**，不符就回 `None`（視同無快取）。不是 semver 比較，
所以 `5.29.0 → 5.29.1` 也全滅。

**這個嚴格度要保留。** 快取存的是「那個版本的 parser 吐出來的 DataFrame」；
edgartools 修了解析 bug 的話，舊快取裡的數字就是帶著那個 bug 的，
而且**不會報錯，只是數字錯**——對財報工具是最糟的失效模式。

### 升級體驗（需求 5）

升級是在命令列發生的，不會在 GUI 裡發生。所以偵測點是**程式啟動時發現版本與
本地資料不符**：

> 偵測到 edgartools 版本從 5.29.0 變成 5.31.0。
> 本地資料庫有 201 家、16,080 份財報（1.4 GB）是用舊版解析的，**將全部失效**。
> 繼續使用需要重新抓取，預估 N 小時。
>
> 〔今晚重抓〕〔立刻重抓〕〔取消升級〕（附回退指令 `pip install edgartools==5.29.0`）

**不提供「照用舊快取」。** 明知可能帶著舊 parser 的解析 bug 還拿來做投資判斷，
換到的只是省一晚。

## 八、GUI

擴充**既有**的快取面板（`main.py:2253` 那個 LabelFrame，現在已有列表、單家清除、
全部清除、容量顯示）。**不開新分頁。**

加上：
- 列表加欄位：涵蓋期間（最舊～最新）、份數、是否到底、上次更新
- 「更新本地庫」按鈕
- 更新名單的管理，含兩個便利動作（匯入 watchlist／匯入快取現況）

⚠ 注意：watchlist 彈窗裡現有的「快取狀態」是**公司名稱**快取，跟 filing 快取
是兩回事，UI 上要避免混淆。

## 九、測試策略

- `_meta.json` 讀寫、自癒（`file_count` 不符時重建）、schema 不符時重建 — 單元測試，`tmp_path`
- `reached_bottom` 推導 — 純函式，餵假的 filing 清單，涵蓋三種判定
- 更新名單的 config 讀寫與兩個便利動作 — 單元測試
- 版本不符的偵測 — 單元測試
- 「整家跳過」的判定 — 單元測試（有新 filing／沒新 filing × 到底／沒到底）
- GUI 對話框沿用專案現況：Tk 探針手動驗，不寫自動測試

## 十、刻意不做（YAGNI）

- **不存原始 XBRL**（容量 42 GB，CTH 否決）
- **不做容量上限／自動淘汰**（1.4 GB 不需要，而且自動刪資料違反「抓過不用重抓」的核心承諾）
- **不做「版本不符但照用」**（正確性風險）
- **不做比 filing 更細的增量**（accession 已經是最細粒度）
- **不動 `fetcher_gaap` 的抓取迴圈**（`reached_bottom` 在外面推導就夠）

## 附錄 A：為什麼不存原始 XBRL（2026-09-04 spike 實測）

技術上**可行**，封網後實測通過：

```
Filing.save(dir) → 單一 .pkl
Filing.load(pkl) → .xbrl() → Financials(xb) → 三張表
封網後 0.6s 重建，2,552 格逐格比對 0 差異  ✅
```

（`_financials_of()` 對 load 回來的 `EntityFiling` 會回 `None`——它沒有
`financials` 屬性。但 `.xbrl()` 拿得到，`Financials(xb)` 就吃 XBRL。）

**但代價（ARLO 實測，每份 filing）：**

| 存什麼 | 大小 | 倍數 | 201 家拓到底（≈16,080 份） |
|---|---|---|---|
| 解析後 DataFrame（現況） | 0.041 MB | 1× | **1.4 GB** |
| ＋ XBRL 檔案（7 個 xml/xsd） | 2.61 MB | 64× | 42 GB |
| ＋ 完整 `Filing.save()` pkl | 7.58 MB | 185× | 122 GB |

另外 `XBRL.from_directory()` 這條較省空間的路**在繁中 Windows 會炸**：
edgartools 讀檔沒指定編碼，用系統預設 cp950 去讀 UTF-8 的 XBRL，直接
`UnicodeDecodeError`；而且會誤把 `FilingSummary.xml` 當成 instance file。

結論：換到的只是「升級時不用連網」，付出 30 倍磁碟。**不做。**
