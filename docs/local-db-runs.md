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
| 已完成 | **134** |
| 未開始 | **67** |

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

原始資料：`output/_localdb/`（`batch1_plan.json`／`batch1_result.json`／
`audit_2026-09-04.json`／`audit_after_batch1.json`／`batch1_round2.json`）。
**`output/` 是 git-ignored 的**，那些檔案只在本機。

## Batch 2（未跑）：剩下的 67 家

```
LVS  MA   MAR  MCD  MCHP MCK  MCO  MDLZ MDT  MET  MMM  MNST MPC  MRK  MRVL
MS   MSCI MSI  MU   NDAQ NEE  NEM  NFLX NKE  NOC  NOW  NSC  NUE  NXPI ODFL
OKE  OMC  ON   ONTO ORCL ORLY OTIS OXY  PANW PAYX PEP  PFE  PG   PLD  PSX
PYPL QCOM RTX  SBUX SCHW SLB  SNOW SNPS SO   SWKS T    TGT  TMO  TSLA TXN
UNH  UNP  UPS  V    WFC  WMT  XOM
```

**推估**：67 家 × 約 75 份 ≈ 5,000 份 × 2.24 s/份 ≈ **3.1 小時**、約 0.45 GB。

怎麼跑：

```bash
# 1. 先看現在缺什麼、下一批是哪些（不下載任何 filing，134 家約 3 分鐘）
./venv/Scripts/python.exe scripts/audit_local_db.py --plan-next 100 \
    --json output/_localdb/audit_before_batch2.json

# 2. 把「需要更新的」跟「下一批」接起來跑
./venv/Scripts/python.exe src/cli.py update-db <上一步印出的 ticker> \
    --json output/_localdb/batch2_result.json

# 3. 收尾：再稽核一次，把還沒到底的再跑一輪
./venv/Scripts/python.exe scripts/audit_local_db.py
```
