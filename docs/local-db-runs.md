# 本地財報資料庫：分批抓取記錄（TODO J5）

> universe＝`output/_hintsweep_201/tickers_joined.txt`（**201 家**）。
> 分批的挑選規則是「照字母序取還沒抓過的前 N 家」
> （`scripts/audit_local_db.py --plan-next N`），**可重現，批次之間不漏也不重複**。
>
> 每批跑完都要用 `audit_local_db.py` 收尾——`update-db` **跑一輪不保證到底**
> （實測 ACN、以及 batch 1 的 16 家都是第二輪才補齊）。

## 進度

| | 家數 |
|---|---|
| universe | **201** |
| 已完成 | **201** ✅ |
| 未開始 | **0** |

**201 家 / 13,921 份 / 0.98 GB**（每家份數中位數 75、最多 77、最少 8）。
全部通過格式驗證。實際容量比設計書估的 1.4 GB 小，因為不少公司受 J6 的
換 CIK 影響，份數不到 75。

涵蓋起始年分布（**J6 的具體規模**）：

| 起始年 | 家數 |
|---|---|
| 2008~2009（滿 17~18 年） | **165** |
| 2010~2014 | 16 |
| 2015~2019 | 15 |
| 2020~2024 | 5 |

也就是說 **165/201（82%）拿得到完整 17~18 年**，其餘 36 家受限於上市時間或
改組換 CIK。

## Batch 1（2026-09-04 ~ 09-05）

**內容**：修補既有 15 家 ＋ 新抓 100 家 ＝ **115 家**
**結果**：第一輪 4h41m59s、新增 **7,542 份**、**0 失敗 0 缺漏**（`gap_tickers` 空）、
**2.24 s/份**。快取 34 家/1,691 份 → **134 家/9,209 份/0.68 GB**。

跑之前就已完整的 19 家（沒動）：

```
AAPL ABBV ABT ACN ADP AEP AFL AIG ALL AMD AMP AMT AON APD APH ARLO AXP GOOGL META
```

修補的 15 家（原本停在舊的 5 年窗）：

```
ADBE ADI AMAT AMGN AMZN ANET AVGO AZO COHR GS JNJ JPM LITE MSFT NVDA
```

新抓的 100 家：

```
BA   BAC  BDX  BK   BKNG BLK  BMY  BSX  C    CAT  CB   CCI  CDNS CDW  CHTR
CI   CL   CMCSA CME CMG  CNC  COF  COP  COST CRM  CSCO CSX  CTAS CTSH CVX
D    DAL  DD   DDOG DE   DHR  DIS  DLR  DOV  DOW  DUK  DXCM EA   ECL  EL
EMR  EOG  EQIX ETN  EW   EXC  F    FAST FDX  FICO FIS  FORM FTNT GD   GE
GEHC GILD GIS  GLW  GM   GWW  HAL  HCA  HD   HIG  HLT  HON  HPQ  HSY  HUM
IBM  ICE  IDXX INTC INTU IP   IQV  IR   ISRG ITW  JCI  KDP  KEYS KHC  KLAC
KMB  KO   KR   LHX  LIN  LLY  LMT  LOW  LRCX LULU
```

**收尾（這一步不能省——`update-db` 跑一輪不保證到底）**：

| 步驟 | 結果 |
|---|---|
| 稽核 #1（134 家，199s） | 完整 118 家、**16 家要第二輪**（合計只缺 24 份） |
| 第二輪（那 16 家，5m11s） | 全部補齊，+24 份，0 失敗 |
| 稽核 #2（134 家，180s） | **完整 134 家、需要更新 0 家** ✅ |

要第二輪的 16 家：

```
ADI AMGN BK BKNG CCI CL CMCSA CNC CVX DLR EOG ETN JPM LHX LMT LOW
```

最終：**134 家 / 9,233 份 / 0.68 GB，全部完整且為 edgartools 5.29.0 解析。**

⚠ 「完整」＝**這個 CIK 的申報全抓到了**，不等於「拿到 18 年歷史」。
改組換過 CIK 的公司（BLK 只有 2 年、DIS/DOW/CI 7 年…）舊申報在別的 CIK，
判定正確但資料比預期少——見 **TODO J6**。

### 格式驗證（`scripts/verify_local_db.py`，2026-09-06）

份數對不代表檔案能用，所以每批跑完**另外驗一次格式**。9,233 份逐檔驗，不連網，108 秒：

| 檢查 | 結果 |
|---|---|
| 過 `load_filing()` 四道閘（JSON／schema／cik／版本） | **9,233 / 9,233＝100%** |
| 三張表反序列化（`payload_to_df` 那條路） | **0 份失敗** |
| `_meta.json` 與目錄一致（file_count、各 form count、版本） | **0 家不符** |
| 同一家出現多個 cik | **0 家** |
| 讀不出來的檔案 | **0 份** |
| **有問題的公司** | **0 家** ✅ |

**824 份「空殼」（三張表全 None）不是格式問題**，是忠實記錄上游現實：

- 年份分布 2010:85、2011:18、2012 之後只剩 6，跟 SEC 的 **XBRL 三階段強制時程**
  （2009-06 最大型 → 2010-06 其餘大型加速申報人 → 2011-06 所有其他人）完全吻合
- 剩下那幾份是分拆後的第一份年報（ABBV 2013、KEYS 2014、GEHC 2023）
- **決定性驗證**：挑 7 份 2010~2023 的空殼直接跟 SEC 重抓（COHR／KLAC／FTNT／
  DXCM／KEYS／GEHC／ABBV），**7/7 上游結果完全一樣**，edgartools 自己的訊息就是
  `No statements available in XBRL data`。**重抓也救不回來，不需要重跑**

### 端對端抽驗（5 家跑完整流程）

| ticker | 季表期數 | 涵蓋 | 關鍵列填滿率 |
|---|---|---|---|
| KO | 69 | FY2009Q2→FY2026Q2 | Revenue／GP／OI／NI／OCF／Capex 皆 96% |
| JPM | 69 | FY2009Q2→FY2026Q2 | Revenue 99%、NI 99%、OCF 96%（GP／Capex 0% 是銀行本來就沒有） |
| LULU | 65 | FY2011Q2→FY2027Q2 | 皆 98% |
| DIS | 30 | FY2019Q2→FY2026Q3 | 皆 97%（期數短是 J6 的換 CIK，不是抓漏） |
| BLK | 8 | FY2024Q3→FY2026Q2 | 皆 88%（同上） |

`Data_Financials(Q)`／`Data_Meta`／`Data_Ratios` 期數三者一致，`Data_Segments` ≤ 之。

原始資料：`output/_localdb/`（`batch1_plan.json`／`batch1_result.json`／
`audit_2026-09-04.json`／`audit_after_batch1.json`／`batch1_round2.json`／
`audit_final_batch1.json`／`verify_batch1.json`／`e2e_*.json`）。
**`output/` 是 git-ignored 的**，那些檔案只在本機。

## Batch 2（2026-09-06）：剩下的 67 家 ✅

**內容**：67 家全新公司（前置稽核確認既有 134 家仍全部完整，0 需要修補）

### ⚠ 第一次跑法失敗了——這件事以後要避免

一開始用一個 process 連跑 67 家，**在第 16 家被系統因記憶體不足中止**
（`status: killed，system is running low on memory`）。原因是 edgartools 的
內部快取會跨公司累積，我們自己的 `_parse_cache_scope()` 只涵蓋單次抓取，擋不住。

**中止沒有白費**——`save_filing()` 逐份即時落檔，已完成的 15 家都在
（134 → 150 家、9,233 → 10,324 份），重跑整家跳過。這是增量設計的價值。

**修法**：`scripts/run_localdb_batch.sh`，把名單切段、**每段一個獨立 process**，
段落結束時記憶體整個還給系統。**以後一律用這支跑，不要一個 process 硬幹。**

```bash
bash scripts/run_localdb_batch.sh output/_localdb/batch2 8 <ticker...>
```

### 結果

| | |
|---|---|
| 分段 | 9 段（每段 8 家），合計 **220 分鐘** |
| 結果 | **0 失敗、0 缺漏**（`gap_tickers` 全空） |
| 第 1~2 段 | 幾乎全跳過——被中止前已完成的那 15 家，證明中斷不白費 |
| 快取 | 134 家/9,233 份 → **201 家/13,921 份/0.98 GB** |

**收尾**：

| 步驟 | 結果 |
|---|---|
| 稽核（201 家，281s） | **完整 201 家、需要更新 0 家**——這次不需要第二輪 |
| 格式驗證（13,921 份，205s） | 四道閘 **13,921/13,921＝100%**、反序列化 0 失敗、meta 0 不符、**有問題的公司 0 家** ✅ |

（batch 1 當時有 16 家要跑第二輪，這次 0 家。分段跑每段獨立 process，
每家都跑得更乾淨。）

1,236 份空殼（2010 前 1,079、2010 後 157）**不是格式問題**，理由同 batch 1。

原始資料：`output/_localdb/batch2_plan.json`／`batch2_chunk01~09.json`／
`audit_before_batch2.json`／`audit_after_batch2.json`／`verify_batch2.json`。

## 之後怎麼維護

universe 已經全部抓完，之後只要定期跑一次把新財報補進來：

```bash
bash scripts/run_localdb_batch.sh output/_localdb/refresh 8 $(cat output/_hintsweep_201/tickers_joined.txt)
./venv/Scripts/python.exe scripts/audit_local_db.py     # 收尾確認
./venv/Scripts/python.exe scripts/verify_local_db.py    # 格式驗證
```

已經到底又沒有新財報的公司會**整家跳過**，所以沒有新財報的時候整輪只要幾分鐘。
