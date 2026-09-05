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
