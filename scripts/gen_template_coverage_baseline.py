"""gen_template_coverage_baseline.py — 產出模板體檢基線文件。

輸出 `docs/template-coverage-baseline-<日期>.md`：每家公司的缺漏判斷、
最常出問題的列、逐列覆蓋率（現行路徑 vs companyfacts）。

**不打網路**，吃 `output/_spike/` 的快取（每家一份 companyfacts JSON ＋ 一份現行
路徑答案卷 pkl）。快取是 `spike_derive_mapping.py` 產生的，沒有的話先跑那支。
2026-08-23 樣本從 52 家擴到 102 家（新增醫療／工業／消費／金融／能源／公用事業／
REIT 各類）。

什麼時候重跑：改了 `IS/BS/CF_TEMPLATE` 的 concept 對照、改了 `data_quality`
的判斷規則、或想確認達標列數有沒有往上走。

**看數字之前先讀產出文件的第零節。** 達標列數只是體溫計，真正的 KPI 是
〔真缺口〕與〔假警報〕；而且**永遠不該以 97/97 為目標**。
"""
import sys, json, pickle, pathlib, statistics as st
from collections import Counter, defaultdict
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import data_quality as dq, fetcher_facts as ff, facts_mapping as fm
from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE, StatementTable

C = ROOT / "output" / "_spike"
TAGS = [("IS", IS_TEMPLATE), ("BS", BS_TEMPLATE), ("CF", CF_TEMPLATE)]
rows = [(tag, r[0]) for tag, T in TAGS for r in T]

# 「達標」的兩個門檻。文件開頭的說明與下面的計分共用這兩個常數，不要各寫一份
# ——改了門檻卻忘了改說明，讀的人會被誤導。
#
# ⚠ 覆蓋門檻**一定要用比例、不能用絕對家數**。原本寫死 `>=45 家`，是照 52 家
# 樣本訂的（45/52 ≈ 87%）。2026-08-23 樣本擴到 102 家時，同一個「45」等於門檻
# 從 87% 悄悄鬆到 44%，達標列數會從 47 跳到 74——**看起來大幅進步，其實是尺變了**。
MIN_CO_RATIO = 0.85   # 至少多少比例的公司抓得到
MIN_FILL = 0.9        # 抓得到的公司裡，填滿率中位數要超過多少

cur_co = Counter(); fac_co = Counter(); fill = {r: [] for _, r in rows}
# facts 路徑的填滿率——H1 要的數字。分母跟現行路徑用同一組期末日，才比得動。
fill_facts = {(tag, r): [] for tag, r in rows}
hole_hits = Counter(); contra_hits = Counter(); spor_hits = Counter()
real_gap = defaultdict(list)   # 真缺口：companyfacts 有、我們整列空白
census = Counter()             # 每個（列 × 公司）格子的三分類
per_co = []
for p in sorted(C.glob("gaap_*.pkl")):
    tk = p.stem[5:]; fp = C / f"facts_{tk}.json"
    if not fp.exists(): continue
    g = pickle.loads(p.read_bytes()); raw = json.loads(fp.read_bytes())
    ends = [e for e in g["ends"] if len(e or "") == 10]; n = max(len(ends), 1)
    seen = set()
    for i, name in enumerate(g["concepts"]):
        if name in fill and name not in seen:
            seen.add(name)
            k = sum(1 for v in g["values"][i] if isinstance(v, (int, float)))
            if k: cur_co[name] += 1; fill[name].append(k / n)
    ours_has = {name for i, name in enumerate(g["concepts"])
                if any(isinstance(v, (int, float)) for v in g["values"][i])}
    for tag, name in rows:
        for M in (fm.IS_MAPPING, fm.BS_MAPPING, fm.CF_MAPPING):
            if name in M:
                fseries = ff.resolve_row(raw, M[name], prefer="as_reported")
                facts_has = bool(fseries)
                if facts_has:
                    fac_co[(tag, name)] += 1
                    fill_facts[(tag, name)].append(
                        sum(1 for e in ends
                            if isinstance(fseries.get(e), (int, float))) / n)
                # 真缺口：這家公司**確實 tag 過**（companyfacts 讀得到），我們卻整列空白。
                # 這才是「該抓到卻沒抓到」；兩邊都沒有代表公司真的沒報，不算我們的問題。
                if facts_has and name not in ours_has:
                    real_gap[name].append(tk)   # 同名列合看，不分 IS/CF
                # 三分類，用來回答「XBRL 裡到底有沒有模板要的數字」
                if name in ours_has:      census["ours"] += 1
                elif facts_has:           census["gap"] += 1
                else:                     census["absent"] += 1
                break
        else:
            census["nomap"] += 1   # 模板有這一列，但 facts_mapping 沒有對應（如 Free Cash Flow）
    t = StatementTable(sheet_name="Data_Financials(Q)", quarter_labels=g["labels"],
                       filing_dates=[""] * len(g["labels"]), concepts=g["concepts"],
                       values=g["values"], ticker=tk,
                       labels=[""] * len(g["concepts"]), period_ends=g["ends"])
    r = dq.assess(t)
    for h in r.holed: hole_hits[h.row] += 1
    for sp in r.sporadic: spor_hits[sp.row] += 1
    for c in r.contradictions: contra_hits[c.row] += 1
    per_co.append((tk, r))

N = len(per_co)
MIN_CO = round(N * MIN_CO_RATIO)   # 門檻隨樣本數走，見上面 MIN_CO_RATIO 的說明
def med(x): return st.median(x) if x else 0.0
L = []
w = L.append
w(f"# 模板體檢：{N} 家公司的逐列覆蓋率（{date.today():%Y-%m-%d} 產出）\n")
w(f"**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（{N} 家的")
w("companyfacts JSON 與現行路徑答案卷快取），重跑不用打網路。\n")
_np = sorted(r.total_periods for _, r in per_co)
w(f"**答案卷的抓取窗**：{N} 家全部用 `spike_derive_mapping.py` 的固定參數"
  "（`max_filings=16`／`max_annual_filings=5`）重建，期數 "
  f"{_np[0]}~{_np[-1]}（中位數 {int(st.median(_np))}）。\n")
w("> 2026-08-24 那份基線的抓取窗**不一致**：AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/")
w("> META/MSFT/NVDA/TSLA/WMT 這 12 家是用全部 filing 抓的（44~69 期），其餘才是")
w("> `max_filings=16`。2026-09-04 重建時統一成 16/5，所以**那 12 家的逐列覆蓋率")
w("> 會跟舊基線對不起來——那是抓取窗變了，不是回歸**。其餘 189 家可以直接比。\n")
w("公司清單刻意涵蓋大中小型 × 跨產業，**包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）**")
w("——它們的報表結構跟製造業差很多，是檢驗模板通不通用最有效的一群。\n")
w("## 零、這份文件怎麼讀（先看這段，不然數字會誤導）\n")
w("### 「達標列數」是什麼\n")
w("一列要「達標」必須**同時**滿足兩個條件：\n")
w("```")
w(f"有值的公司數 >= {MIN_CO} 家（{MIN_CO_RATIO:.0%} 的樣本）   這一列在絕大多數公司都抓得到")
w(f"填滿率中位數 > {MIN_FILL:.0%}      抓得到的那些公司，幾乎每一季都有值")
w("```\n")
w("兩個缺一不可。只滿足前者代表「大家都有、但常常缺季」；只滿足後者代表")
w("「少數公司很完整、多數抓不到」。兩種都不能算穩。\n")
w("### 不要追求 97/97，那個目標本身是錯的\n")
w("達標門檻假設「這一列應該人人都有」，但有些列天生就不該。**達不到標不等於有 bug**：\n")
w("| 列 | 為什麼永遠達不了標 |")
w("|---|---|")
w("| `Preferred Stock` | 多數公司根本沒發特別股 |")
w("| `Minority Interest` / `Noncontrolling Interests` | 沒有非控制權益的公司就是沒有 |")
w("| `Finance Lease Liabilities, LT` | 多數公司只有營業租賃 |")
w("| `R&D Expense` | 零售、餐飲、能源業不在損益表單獨揭露 |")
w("| `Pension & Retirement Oblig.` | 大多數公司沒有確定給付制退休金 |")
w("")
w("**真正該當 KPI 的是第六節那兩個數字**：〔真缺口〕該抓到卻沒抓到幾列、")
w("〔假警報〕Index 標紅裡有幾個是誤判。達標列數只是一支粗略的體溫計。\n")
w("## 一、每家公司的缺漏判斷\n")
w("`data_quality.assess()` 的四個判斷。缺季那一欄幾乎每家都是 1，要打折看：")
w("答案卷是 `max_filings=16` 抓的，最舊那一年的 Q4 合成材料不足。\n")
w("| ticker | 期數 | 缺季 | 稀疏欄 | 有洞列 | 矛盾 | 模板不適用 |")
w("|---|---|---|---|---|---|---|")
for tk, r in sorted(per_co, key=lambda x: -len(x[1].sparse_periods)):
    w(f"| {tk} | {r.total_periods} | {sum(g.count for g in r.missing_quarters)} | "
      f"{len(r.sparse_periods)} | {len(r.holed)} | {len(r.contradictions)} | "
      f"{'**是**' if r.template_mismatch else ''} |")
w("")
_mismatch = [tk for tk, r in per_co if r.template_mismatch]
w(f"**觸發「模板不適用」的 {len(_mismatch)} 家：{', '.join(sorted(_mismatch))}**"
  "——全是金融股。IS/BS/CF 模板是為製造業設計的，銀行／券商的報表結構完全不同"
  "（存款、放款、備抵呆帳…），這是 TODO D8 記錄的已知限制，現在有量化證據。\n")
w("## 二、最常出問題的列\n")
w("### 中間有洞（同一列有些期有、有些沒有——一定是漏抓）\n")
w("| 列名 | 幾家中招 |")
w("|---|---|")
for k, v in hole_hits.most_common(15): w(f"| {k} | {v} / {N} |")
w("")
w("### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）\n")
w("2026-08-23（H3-2）從「中間有洞」拆出來的一類。當時拿 companyfacts 當真值驗 52 家、")
w("2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才")
w("到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。\n")
w("| 列名 | 幾家中招 |")
w("|---|---|")
for k, v in spor_hits.most_common(15): w(f"| {k} | {v} / {N} |")
w("")
w("### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）\n")
w("| 列名 | 幾家中招 |")
w("|---|---|")
for k, v in contra_hits.most_common(12): w(f"| {k} | {v} / {N} |")
w("")
# 2026-08-23（H3）改寫。原本這裡寫「`Current Portion of LT Debt` 25 家中招，
# 幾乎確定是 concept 對照有問題」——查證後是**錯的判斷**：INTC/PG/XOM/MU/NVDA/
# QCOM/PFE/ORCL 八家的資產負債表表面只有一條流動借款列（`us-gaap:DebtCurrent`），
# 一年內到期的長期負債併在裡面、早就進了 Short-term Debt。問題出在判斷規則，
# 不是抓取。留這段話當提醒，別再從「中招家數多」直接跳到「concept 名字錯」。
w("**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`")
w("等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，")
w("只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，")
w("先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。\n")
w("## 三、逐列覆蓋率：現行路徑 vs companyfacts\n")
w(f"「有值公司數」＝ {N} 家裡有幾家這一列拿得到資料。兩邊差 8 家以上的標 ⚠。\n")
w("「填滿」＝抓得到的那些公司裡，答案卷的期末日有幾成拿得到值（中位數）。\n")
w("| 表 | 列名 | 現行家數 | facts家數 | 差 | 現行填滿 | facts填滿 |")
w("|---|---|---|---|---|---|---|")
for tag, name in rows:
    a, b = cur_co[name], fac_co[(tag, name)]
    flag = " ⚠" if abs(a - b) >= 8 else ""
    w(f"| {tag} | {name} | {a} | {b} | {b - a:+d}{flag} | "
      f"{med(fill[name]):.0%} | {med(fill_facts[(tag, name)]):.0%} |")
w("")
# H1 專用：CF 流量列（mapping 標了 `from_ytd`）的 facts 填滿率。H1 原本記
# 「只有約 25%（一年四季只拿得到一季）」，`quarterly_from_ytd()` 上線後看這裡。
_ytd = [(tag, name) for tag, name in rows
        for M in (fm.IS_MAPPING, fm.BS_MAPPING, fm.CF_MAPPING)
        if name in M and M[name].get("from_ytd")]
_ytd_med = [med(fill_facts[k]) for k in _ytd if fill_facts[k]]
w(f"**流量列（mapping 標 `from_ytd` 的 {len(_ytd)} 列）的 facts 填滿率中位數："
  f"{st.median(_ytd_med) if _ytd_med else 0:.0%}**——這是 TODO H1 的驗收數字。")
w("H1 記錄的原始症狀是「約 25%，一年四季只拿得到一季」，成因是公司把現金流量表的")
w("項目 tag 成 YTD 累計；`fetcher_facts.quarterly_from_ytd()` 已經補上還原單季那一層。\n")
good = sum(1 for _, name in rows if cur_co[name] >= MIN_CO and med(fill[name]) > MIN_FILL)
w(f"**現行路徑達到「>={MIN_CO} 家（{MIN_CO_RATIO:.0%}）有值且填滿率 >{MIN_FILL:.0%}」的列：{good} / {len(rows)}**")
w("（這個數字不該以 97/97 為目標，理由見第零節。）\n")

# ── 四、哪些數字是推理出來的 ────────────────────────────────────────────
# 這張表是**手寫的**，因為「這一格是直接讀還是推出來的」不在答案卷 pkl 裡，
# 要從 fetcher_gaap 的後處理程式碼讀出來。改那些邏輯時要同步改這裡。
n_q4 = sum(1 for p2 in sorted(C.glob("gaap_*.pkl"))
           for lbl in pickle.loads(p2.read_bytes())["labels"] if str(lbl).endswith("Q4"))
n_all = sum(len(pickle.loads(p2.read_bytes())["labels"]) for p2 in sorted(C.glob("gaap_*.pkl")))
w("## 四、哪些數字是直接讀 XBRL、哪些是推理出來的\n")
w("**不是每一格都是從財報直接讀出來的。** 下面這些是程式算出來的，來源在")
w("`fetcher_gaap.py` 的後處理段落。看數字有疑問時先確認它屬於哪一類。\n")
w("### A. 整列都是算的\n")
w("| 列 | 算式 |")
w("|---|---|")
w("| `Free Cash Flow` | 營運現金流 − 資本支出取絕對值。**XBRL 沒有這個 tag**，本來就只能算 |")
w("")
w("### B. 抓不到才用算的（抓得到就用公司報的）\n")
w("| 列 | 算式 | 什麼情況會用到 |")
w("|---|---|---|")
w("| `Gross Profit` | 營收 − 銷貨成本 | GOOGL／AMZN 等損益表沒有毛利小計行的公司 |")
w("| `Total Non-current Assets` | 總資產 − 流動資產 | 多數公司不標 `AssetsNoncurrent` |")
w("| `Total Non-current Liabilities` | 總負債 − 流動負債 | 多數公司不標 `LiabilitiesNoncurrent` |")
w("| `Total Non-op Income/(Loss)` | 稅前淨利 − 營業利益 | 沒有營業外損益合計行的公司 |")
w("")
w("### C. 多列加總（不是挑一條）\n")
w("| 列 | 加總範圍 |")
w("|---|---|")
w("| `Debt Proceeds` | 所有借款流入（長期、短期、商業本票、可轉債…），排除淨額列 |")
w("| `Debt Repayments` | 所有還款流出，排除淨額列 |")
w("| `Investment Proceeds` | 所有投資處分／到期流入 |")
w("")
w("### D. 期間換算（影響範圍最大，最容易被忽略）\n")
w("| 什麼 | 怎麼算 |")
w("|---|---|")
w("| **現金流量表的每一個單季值** | 公司多半只 tag 年初至今累計 → 本季 YTD − 上季 YTD |")
w("| **每一個 Q4 欄** | 10-Q 只有 Q1~Q3，Q4 由年報 − Q1 − Q2 − Q3 合成（餘額列直接取年報值） |")
w("")
w(f"本次 {N} 家共 {n_all} 個期間欄，其中 **{n_q4} 欄是 Q4**（{n_q4/max(n_all,1):.0%}）——")
w("這些欄的流量列全部是合成的。Q1~Q3 不齊全時合成會失敗，那一整欄會空掉，")
w("`data_quality` 的「整欄稀疏」就是用來抓這件事的。\n")
w("## 五、XBRL 裡到底有沒有模板要的數字\n")
w(f"把「{len(rows)} 個模板列 × {N} 家公司」每一格分成三類。**判斷「有沒有」靠")
w("companyfacts**（它讀得到公司 tag 過的全部 fact，含附註層），比只看報表表面準。\n")
_tot = census['ours'] + census['gap'] + census['absent']
w("| 分類 | 格數 | 佔比 | 意思 |")
w("|---|---|---|---|")
w(f"| 我們抓到了 | {census['ours']} | {census['ours']/max(_tot,1):.0%} | 正常 |")
w(f"| **真缺口** | {census['gap']} | {census['gap']/max(_tot,1):.0%} | 公司有 tag，我們沒抓到 → 見下面 KPI 1 |")
w(f"| 公司真的沒有 | {census['absent']} | {census['absent']/max(_tot,1):.0%} | **不是問題**，這家公司就是沒報這個科目 |")
w("")
w(f"另有 {census['nomap']} 格不列入分類：那些模板列在 `facts_mapping` 裡沒有對應")
w("concept（例如 `Free Cash Flow`，XBRL 本來就沒有這個 tag），無從判斷「有沒有」。\n")
w("**所以答案是：不是每一格都存在。** 「公司真的沒有」那一類佔了相當比例，而且")
w("**那是正常的**——沒發特別股、沒有非控制權益、不揭露 R&D 的公司本來就不該有值。")
w("值得追的只有中間那一類。\n")
w("## 六、兩個真正的 KPI\n")
w("### KPI 1 — 真缺口：該抓到卻沒抓到\n")
w("判準：**這家公司確實 tag 過**（companyfacts 讀得到），我們卻整列空白。")
w("兩邊都沒有的不算——那是公司真的沒報，不是我們的問題。\n")
w("| 列名 | 幾家真缺 | 哪幾家 |")
w("|---|---|---|")
gap_rank = sorted(real_gap.items(), key=lambda x: -len(x[1]))
for k, v in gap_rank[:20]:
    w(f"| {k} | {len(v)} / {N} | {', '.join(sorted(v)[:8])}{' …' if len(v) > 8 else ''} |")
w("")
w(f"**真缺口總計：{sum(len(v) for v in real_gap.values())} 個（列 × 公司）組合，"
  f"分布在 {len(real_gap)} 個模板列。**")
w("")
w("榜首那幾列全部是 TODO D10（只寫在附註、沒印在報表表面）——這是**已知的暫時性")
w("限制**，不是新 bug。要壓低這個數字只有兩條路：接一條讀附註的路徑，或接受它。\n")
w("### KPI 2 — 假警報：Index 標紅裡有幾個是誤判\n")
w("標紅只有兩類：〔矛盾〕整列空白但相關欄位顯示該有、〔中間有洞〕。")
w("「零星有值」刻意不標紅（H3-2），所以不算在內。\n")
w("| | 家次 |")
w("|---|---|")
w(f"| 標紅：矛盾 | {sum(contra_hits.values())} |")
w(f"| 標紅：中間有洞 | {sum(hole_hits.values())} |")
w(f"| **標紅合計** | **{sum(contra_hits.values()) + sum(hole_hits.values())}** |")
w(f"| 降級為零星有值（不標紅） | {sum(spor_hits.values())} |")
w("")
w("**要壓低的是標紅合計裡的誤判比例**，不是把標紅壓到 0——真缺口該標就要標。")
w("驗證方式：對標紅的列抽樣，走 ARCHITECTURE「三步排查順序」確認是哪一類。\n")
w("## 七、怎麼重跑\n")
w("```")
w("venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢")
w("venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒")
w("```")
out = ROOT / "docs" / f"template-coverage-baseline-{date.today():%Y-%m-%d}.md"
out.write_text("\n".join(L), encoding="utf-8")
print("written", out, len(L), "lines")
