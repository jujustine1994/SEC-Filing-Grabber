"""spike_validate_facts.py — companyfacts 資料本身的獨立驗證（TODO G11 第三步）。

`spike_derive_mapping.py` 是拿現行路徑當答案卷去反推 mapping，那有個先天限制：
**現行路徑錯的地方，比對也會跟著錯**。這支從另外四個角度驗證 companyfacts
本身可不可信，不依賴現行路徑，所以只要打 companyfacts（每家 0.5 秒），
可以一次驗二三十家。

四項檢查：

1. **會計恆等式**：資產 = 負債 + 權益。這是硬約束，對不上就代表期間對齊或
   取版規則有問題。用它驗「instant 期間分類」與「重編取版」是否正確。

2. **四季加總 = 年度**：Q1+Q2+Q3+Q4 是否等於同財年的年度值。這是驗
   「80~100 天算單季」這條規則最直接的方式——如果把半年報或 YTD 誤收進來，
   加總一定爆掉。

3. **SEC 自己的 `frame` vs 我們的期中點判準**：每筆 fact 上的 `frame`
   （如 `CY2025Q2`）是 SEC 官方的日曆季正規化。拿它跟 `fiscal_input`
   的 `basis="span"` 比，**這是對 F6/G2 跨公司對齊決策的獨立驗證**。

4. **重編頻率**：同一期間 as_reported 與 latest 差多少、多常差。這決定
   `prefer` 的預設值該選哪個（TODO G11 的待決事項之一）。

用法：
    venv/Scripts/python.exe scripts/spike_validate_facts.py                # 預設 24 家
    venv/Scripts/python.exe scripts/spike_validate_facts.py NVDA AAPL
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import requests

import config
import fetcher_facts as ff
import fiscal_input as fi
from edgar import Company, set_identity

CACHE = ROOT / "output" / "_spike"
CACHE.mkdir(parents=True, exist_ok=True)

# 刻意涵蓋不同財年結束月、不同產業、不同規模。金融業（JPM/GS）的報表結構
# 跟製造業差很多，是最容易踩到「模板假設」的一群，一定要放進來。
DEFAULT_TICKERS = [
    # 12 月結算
    "AMD", "INTC", "TSLA", "KO", "JPM", "GS", "JNJ", "AMZN", "GOOGL", "META",
    # 1 月結算
    "NVDA", "WMT", "CRM", "MRVL",
    # 其他結算月
    "AAPL",   # 9
    "MSFT",   # 6
    "AVGO",   # 11
    "COST",   # 8/9
    "PANW",   # 7
    "LITE",   # 6/7
    "COHR",   # 6
    "ORCL",   # 5
    "ADBE",   # 11/12
    "NKE",    # 5
]

_EQUITY_CONCEPTS = ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
                    "StockholdersEquity"]
# 夾層權益（可贖回非控制權益／可贖回特別股）。US-GAAP 上它既不在 Liabilities
# 也不在 StockholdersEquity 裡，是資產負債表中間獨立的一塊。不加它，TSLA
# 這類有 redeemable NCI 的公司會全部判成「不平」——那是檢查寫錯，不是資料錯。
_MEZZANINE_CONCEPTS = ["TemporaryEquityCarryingAmountAttributableToParent",
                       "TemporaryEquityCarryingAmountIncludingPortionAttributableToNoncontrollingInterests",
                       "RedeemableNoncontrollingInterestEquityCarryingAmount"]
_FLOW_CONCEPTS = ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
                  "NetIncomeLoss"]


def _facts(ticker: str, identity: str) -> dict | None:
    p = CACHE / f"facts_{ticker}.json"
    if not p.exists():
        try:
            cik = int(Company(ticker).cik)
        except Exception as e:
            print(f"  [{ticker}] 查不到 CIK: {type(e).__name__}", file=sys.stderr)
            return None
        r = requests.get(ff.COMPANYFACTS_URL.format(cik=cik),
                         headers={"User-Agent": identity}, timeout=60)
        if r.status_code != 200:
            print(f"  [{ticker}] HTTP {r.status_code}", file=sys.stderr)
            return None
        p.write_bytes(r.content)
    return json.loads(p.read_bytes())


def _series(raw, concept, kind, prefer="as_reported"):
    return ff.series_for_concept(raw, concept, kind=kind, prefer=prefer)


def _first_series(raw, concepts, kind):
    for c in concepts:
        s = _series(raw, c, kind)
        if s:
            return c, s
    return None, {}


# ── 1. 會計恆等式 ───────────────────────────────────────────────────────────

def check_balance_sheet(raw) -> tuple[int, int, list]:
    assets = _series(raw, "Assets", "instant")
    liab = _series(raw, "Liabilities", "instant")
    _, equity = _first_series(raw, _EQUITY_CONCEPTS, "instant")
    _, mezz = _first_series(raw, _MEZZANINE_CONCEPTS, "instant")
    ok = bad = 0
    examples = []
    for end in sorted(set(assets) & set(liab) & set(equity)):
        lhs = assets[end]
        rhs = liab[end] + equity[end] + mezz.get(end, 0.0)
        # XBRL 有些公司年度值報到百萬、季度值報到千，兩邊精度不同。
        # 用相對誤差 0.1% 當門檻——真正的不平會差得比這多得多。
        if abs(lhs - rhs) <= 1e-3 * max(abs(lhs), 1.0):
            ok += 1
        else:
            bad += 1
            if len(examples) < 3:
                examples.append((end, lhs, rhs))
    return ok, bad, examples


# ── 2. 四季加總 = 年度 ──────────────────────────────────────────────────────

def check_quarters_sum_to_year(raw, fy_start_month) -> tuple[int, int, list]:
    ok = bad = 0
    examples = []
    for concept in _FLOW_CONCEPTS:
        q = _series(raw, concept, "quarter")
        a = _series(raw, concept, "annual")
        if not q or not a:
            continue
        by_fy = {}
        for end, val in q.items():
            fy = fi.fiscal_year_of(end, fy_start_month)
            by_fy.setdefault(fy, []).append(val)
        for end, annual in a.items():
            fy = fi.fiscal_year_of(end, fy_start_month)
            qs = by_fy.get(fy, [])
            if len(qs) != 4:
                continue          # 那一年沒收齊四季，這裡不判定
            # 同上：精度不同造成的尾數差不算錯，用相對誤差 0.1%
            if abs(sum(qs) - annual) <= 1e-3 * max(abs(annual), 1.0):
                ok += 1
            else:
                bad += 1
                if len(examples) < 3:
                    examples.append((concept, fy, sum(qs), annual))
        break                      # 一個 concept 夠了，不重複算
    return ok, bad, examples


# ── 3. SEC 的 frame vs 我們的期中點判準 ─────────────────────────────────────

def check_frame_vs_span(raw) -> tuple[int, int, list]:
    """SEC 的 `frame` 是它自己的日曆季正規化，拿來獨立驗證我們的 span 判準。

    只看 duration 的季度 fact（`CY2025Q2` 這種）。instant 的 frame 帶 `I` 結尾
    （`CY2026Q1I`），語意不同，不比。
    """
    agree = disagree = 0
    examples = []
    for node in raw.get("facts", {}).get("us-gaap", {}).values():
        for fact in node.get("units", {}).get("USD", []):
            frame = fact.get("frame", "")
            if not frame.startswith("CY") or frame.endswith("I") or "Q" not in frame:
                continue
            if ff.classify_period(fact) != "quarter":
                continue
            ours = fi.calendar_quarter_of(fact["end"], basis="span")
            if frame[2:] == ours:
                agree += 1
            else:
                disagree += 1
                if len(examples) < 3:
                    examples.append((fact.get("start"), fact["end"], frame[2:], ours))
    return agree, disagree, examples


# ── 4. 重編頻率 ─────────────────────────────────────────────────────────────

def check_restatements(raw) -> tuple[int, int, list]:
    same = differ = 0
    examples = []
    for name, node in raw.get("facts", {}).get("us-gaap", {}).items():
        buckets = {}
        for fact in node.get("units", {}).get("USD", []):
            if ff.classify_period(fact) != "quarter":
                continue
            buckets.setdefault(fact["end"], []).append(fact)
        for end, items in buckets.items():
            if len(items) < 2:
                continue
            a = ff.pick_fact(items, prefer="as_reported")["val"]
            b = ff.pick_fact(items, prefer="latest")["val"]
            # 只算「實質不同」：純精度變更（229,724,000 → 230,000,000）
            # 不算重編，那是公司改了申報精度，不是改了數字
            if abs(a - b) <= 1e-3 * max(abs(a), abs(b), 1.0):
                same += 1
            else:
                differ += 1
                if len(examples) < 3:
                    examples.append((name, end, a, b))
    return same, differ, examples


def main(argv):
    tickers = argv[1:] or DEFAULT_TICKERS
    identity = config.load_config().get("identity")
    if not identity:
        print("沒有 SEC Identity", file=sys.stderr)
        return 1
    set_identity(identity)

    tot = Counter()
    print(f"{'ticker':8}{'資產=負債+權益':>16}{'四季=年度':>14}"
          f"{'frame vs span':>16}{'重編筆數':>12}")
    print("-" * 70)
    notes = []
    for t in tickers:
        raw = _facts(t, identity)
        if raw is None:
            continue
        # 財年起始月從年度 fact 的期末日推——不依賴現行路徑的偵測
        _, annual = _first_series(raw, _FLOW_CONCEPTS, "annual")
        fy_end_month = int(sorted(annual)[-1][5:7]) if annual else 12
        fy_start = fi.fy_start_month(fy_end_month)

        b_ok, b_bad, b_ex = check_balance_sheet(raw)
        s_ok, s_bad, s_ex = check_quarters_sum_to_year(raw, fy_start)
        f_ok, f_bad, f_ex = check_frame_vs_span(raw)
        r_same, r_diff, r_ex = check_restatements(raw)

        tot.update({"b_ok": b_ok, "b_bad": b_bad, "s_ok": s_ok, "s_bad": s_bad,
                    "f_ok": f_ok, "f_bad": f_bad, "r_same": r_same, "r_diff": r_diff})
        print(f"{t:8}{b_ok:>7}/{b_ok + b_bad:<8}{s_ok:>6}/{s_ok + s_bad:<7}"
              f"{f_ok:>8}/{f_ok + f_bad:<7}{r_diff:>7}/{r_same + r_diff:<5}")
        for tag, ex in [("BS 不平", b_ex), ("四季≠年度", s_ex),
                        ("frame≠span", f_ex), ("重編", r_ex)]:
            for e in ex:
                notes.append(f"  [{t}] {tag}: {e}")

    print("-" * 70)
    print(f"{'合計':8}{tot['b_ok']:>7}/{tot['b_ok'] + tot['b_bad']:<8}"
          f"{tot['s_ok']:>6}/{tot['s_ok'] + tot['s_bad']:<7}"
          f"{tot['f_ok']:>8}/{tot['f_ok'] + tot['f_bad']:<7}"
          f"{tot['r_diff']:>7}/{tot['r_same'] + tot['r_diff']:<5}")
    if notes:
        print("\n--- 不一致的例子 ---")
        for n in notes[:40]:
            print(n)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
