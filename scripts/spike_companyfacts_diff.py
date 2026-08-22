"""spike_companyfacts_diff.py — companyfacts 路徑 vs 現行解 filing 路徑的逐格比對。

TODO G11 的決策依據。**不改任何現有程式**，只是把兩條路跑同一批公司，
把差異列出來讓 CTH 看數據決定要不要切換。

用法：
    venv/Scripts/python.exe scripts/spike_companyfacts_diff.py NVDA AMD AAPL

輸出四段：
    1. 耗時對照（這是 G11 的主要賣點，要有實測數字）
    2. 逐格比對：兩邊都有值時數字對不對，各自獨有的期間有多少
    3. 覆蓋率：每個模板列在兩條路各拿到幾期
    4. Q4 來源：10-K 直接 tag 的 Q4 有多少（那些不用再合成）

比對的是 `Data_Financials(Q)` 的模板列。segments / overflow / B 欄標籤
不在比對範圍——companyfacts 結構上就拿不到那些（見 fetcher_facts 的模組說明）。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import requests

import config
import fetcher_facts as ff
from edgar import Company, set_identity
from fetcher_gaap import (
    BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE,
    fetch_gaap_statements,
)

# 兩邊數字差多少才算「對不上」。XBRL 是整數分位，理論上要完全相等；
# 留一點餘裕是為了容忍浮點往返，不是為了容忍真正的差異。
_REL_TOL = 1e-9

# BS 是時點值，IS/CF 是期間值
_TEMPLATES = [("IS", IS_TEMPLATE, "quarter"),
              ("BS", BS_TEMPLATE, "instant"),
              ("CF", CF_TEMPLATE, "quarter")]


def _facts_series(raw, template, kind, prefer):
    """模板 → {列名: {期末日: 值}}。

    模板的 `std_concept` 是 edgartools 的 standard_concept 命名，`fallback`
    通常才是真正的 us-gaap element name——兩個都試，primary 有值就不看 fallback。
    """
    out = {}
    for row in template:
        row_name, std_concept, fallback = row[0], row[1], row[2]
        out[row_name] = ff.series_for_concept(
            raw, std_concept, kind=kind, prefer=prefer,
            fallbacks=[fallback] if fallback else None)
    return out


def _gaap_series(q_tbl):
    """現行路徑的 StatementTable → {列名: {期末日: 值}}。

    同名列（IS 與 CF 都有 `SBC`）取第一個出現的，跟 Excel 上看到的一致。
    """
    ends = q_tbl.period_ends or []
    out = {}
    for i, name in enumerate(q_tbl.concepts):
        if name in out or not name:
            continue
        series = {}
        for j, end in enumerate(ends):
            if len(end or "") == 10 and q_tbl.values[i][j] is not None:
                series[end] = q_tbl.values[i][j]
        out[name] = series
    return out


def compare(ticker: str, identity: str, prefer: str = "as_reported") -> None:
    print(f"\n{'=' * 72}\n{ticker}\n{'=' * 72}")

    cik = Company(ticker).cik
    t0 = time.time()
    raw = requests.get(COMPANYFACTS := ff.COMPANYFACTS_URL.format(cik=int(cik)),
                       headers={"User-Agent": identity}, timeout=60).json()
    t_facts = time.time() - t0

    t0 = time.time()
    tables = fetch_gaap_statements(ticker, identity,
                                   fetch_quarterly=True, fetch_annual=True)
    t_gaap = time.time() - t0

    print(f"\n[1] 耗時   companyfacts {t_facts:6.2f}s   "
          f"解 filing {t_gaap:7.2f}s   差 {t_gaap / max(t_facts, 1e-9):.0f}x")

    q_tbl = next(t for t in tables if t.sheet_name == "Data_Financials(Q)")
    gaap = _gaap_series(q_tbl)

    facts = {}
    for _, template, kind in _TEMPLATES:
        facts.update(_facts_series(raw, template, kind, prefer))

    same = diff = only_facts = only_gaap = 0
    diff_examples = []
    for row_name, f_series in facts.items():
        g_series = gaap.get(row_name, {})
        for end in set(f_series) | set(g_series):
            fv, gv = f_series.get(end), g_series.get(end)
            if fv is not None and gv is not None:
                if abs(fv - gv) <= _REL_TOL * max(abs(fv), abs(gv), 1.0):
                    same += 1
                else:
                    diff += 1
                    if len(diff_examples) < 12:
                        diff_examples.append((row_name, end, gv, fv))
            elif fv is not None:
                only_facts += 1
            else:
                only_gaap += 1

    total = same + diff + only_facts + only_gaap
    print(f"\n[2] 逐格比對（共 {total} 格）")
    print(f"    兩邊都有且相同 {same:5}   兩邊都有但不同 {diff:5}")
    print(f"    只有 facts 有  {only_facts:5}   只有現行路徑有 {only_gaap:5}")
    if diff_examples:
        print("\n    不同的例子（列名 / 期末日 / 現行 / facts）:")
        for n, e, gv, fv in diff_examples:
            print(f"      {n[:32]:34} {e}  {gv:>18,.0f}  {fv:>18,.0f}")

    print("\n[3] 模板列覆蓋率（現行 → facts）")
    for row_name in facts:
        g_n, f_n = len(gaap.get(row_name, {})), len(facts[row_name])
        if g_n or f_n:
            flag = "  <<<" if f_n > g_n + 2 or g_n > f_n + 2 else ""
            print(f"    {row_name[:34]:36} {g_n:3} → {f_n:3}{flag}")

    rev = ff.series_for_concept(raw, "Revenues", kind="quarter", prefer=prefer,
                                fallbacks=["RevenueFromContractWithCustomerExcludingAssessedTax"])
    node = (raw.get("facts", {}).get("us-gaap", {}).get("Revenues")
            or raw.get("facts", {}).get("us-gaap", {})
                  .get("RevenueFromContractWithCustomerExcludingAssessedTax") or {})
    from_10k = {f["end"] for f in node.get("units", {}).get("USD", [])
                if ff.classify_period(f) == "quarter" and f.get("form") == "10-K"}
    print(f"\n[4] Revenue 單季共 {len(rev)} 期，其中 {len(from_10k)} 期由 10-K 直接 tag"
          f"（這些是 Q4，不用再合成）")
    print(f"    us-gaap concept 總數 {len(raw.get('facts', {}).get('us-gaap', {}))}")


def main(argv):
    tickers = argv[1:] or ["NVDA"]
    identity = config.load_config().get("identity")
    if not identity:
        print("沒有 SEC Identity，先在進階設定填好再跑", file=sys.stderr)
        return 1
    set_identity(identity)
    for t in tickers:
        try:
            compare(t, identity)
        except Exception as e:
            print(f"[{t}] 失敗 {type(e).__name__}: {e}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
