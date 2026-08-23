"""gen_template_coverage_baseline.py — 產出模板體檢基線文件。

輸出 `docs/template-coverage-baseline-<日期>.md`：每家公司的缺漏判斷、
最常出問題的列、逐列覆蓋率（現行路徑 vs companyfacts）。

**不打網路**，吃 `output/_spike/` 的快取（52 家的 companyfacts JSON 與現行
路徑的答案卷）。那些快取是 `spike_derive_mapping.py` 產生的，沒有的話先跑那支。

什麼時候重跑：改了 `IS/BS/CF_TEMPLATE` 的 concept 對照、改了 `data_quality`
的判斷規則、或想確認「40/97」這個覆蓋率數字有沒有往上走。
"""
import sys, json, pickle, pathlib, statistics as st
from collections import Counter
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import data_quality as dq, fetcher_facts as ff, facts_mapping as fm
from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE, StatementTable

C = ROOT / "output" / "_spike"
TAGS = [("IS", IS_TEMPLATE), ("BS", BS_TEMPLATE), ("CF", CF_TEMPLATE)]
rows = [(tag, r[0]) for tag, T in TAGS for r in T]

cur_co = Counter(); fac_co = Counter(); fill = {r: [] for _, r in rows}
hole_hits = Counter(); contra_hits = Counter(); spor_hits = Counter()
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
    for _, name in rows:
        for M in (fm.IS_MAPPING, fm.BS_MAPPING, fm.CF_MAPPING):
            if name in M:
                if ff.resolve_row(raw, M[name], prefer="as_reported"): fac_co[name] += 1
                break
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
def med(x): return st.median(x) if x else 0.0
L = []
w = L.append
w(f"# 模板體檢：{N} 家公司的逐列覆蓋率（{date.today():%Y-%m-%d} 產出）\n")
w("**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（52 家的")
w("companyfacts JSON 與現行路徑答案卷快取），重跑不用打網路。\n")
w("公司清單刻意涵蓋大中小型 × 跨產業，**包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）**")
w("——它們的報表結構跟製造業差很多，是檢驗模板通不通用最有效的一群。\n")
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
w("**金融股與 REIT 全數觸發「模板不適用」**（稀疏欄佔 90~100%）。這是 TODO D8")
w("記錄的已知限制，現在有量化證據。\n")
w("## 二、最常出問題的列\n")
w("### 中間有洞（同一列有些期有、有些沒有——一定是漏抓）\n")
w("| 列名 | 幾家中招 |")
w("|---|---|")
for k, v in hole_hits.most_common(15): w(f"| {k} | {v} / {N} |")
w("")
w("### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）\n")
w("2026-08-23（H3-2）從「中間有洞」拆出來的一類。拿 companyfacts 當真值驗 52 家、")
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
w("「有值公司數」＝ 52 家裡有幾家這一列拿得到資料。兩邊差 ≥8 家的標 ⚠。\n")
w("| 表 | 列名 | 現行 | facts | 差 |")
w("|---|---|---|---|---|")
for tag, name in rows:
    a, b = cur_co[name], fac_co[name]
    flag = " ⚠" if abs(a - b) >= 8 else ""
    w(f"| {tag} | {name} | {a} | {b} | {b - a:+d}{flag} |")
w("")
good = sum(1 for _, name in rows if cur_co[name] >= 45 and med(fill[name]) > 0.9)
w(f"**現行路徑達到「≥45 家有值且填滿率 >90%」的列：{good} / {len(rows)}**\n")
w("## 四、怎麼重跑\n")
w("```")
w("venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢")
w("venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒")
w("```")
out = ROOT / "docs" / f"template-coverage-baseline-{date.today():%Y-%m-%d}.md"
out.write_text("\n".join(L), encoding="utf-8")
print("written", out, len(L), "lines")
