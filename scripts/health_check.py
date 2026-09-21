# -*- coding: utf-8 -*-
"""health_check.py — 全庫體檢：這個工具現在到底多可信（唯讀，不改任何東西）。

## 這支在答什麼

**「我輸入一個代號，出來的 Excel 有幾成可信？」**

`docs/template-coverage-baseline-2026-08-24.md` 是現有的體檢基線，但那份文件
**自己標註已經過時**（H6 放寬四條 label_hint、201 家多救回 37 家，基線沒跟著
更新，因為要重跑就得真的抓 201 家 SEC）。這支改吃**本地財報資料庫**，
不必重抓，而且涵蓋的是實際在用的那 215 家。

## 判斷邏輯直接用正式路徑的函式

`data_quality.assess()` 就是 Excel 的 Index 頁標紅用的那一套，
`fetch_gaap_statements()` 就是按下「執行」跑的那條路。**所以體檢結論跟打開
Excel 看到的一致，不可能分岔。**

## ⚠ 一定要用正式參數跑

`max_filings` 調小會讓 `missing_quarters` 冒出一堆假缺口、`template_mismatch`
變成 True——實測 ARLO 用 `max_filings=6` 跑出 6 段假缺季。**抽樣會把體檢數字
整個扭曲**，所以這支一律用預設參數，跟使用者按下「執行」時一樣。

## 中斷續跑

逐家即時落檔到 `output/_health/<ticker>.json`，重跑時已完成的整家跳過。
215 家約 1.5 小時（快取全命中，主要成本是比對層的 CPU 與每家一次清單請求）。

    venv/Scripts/python.exe scripts/health_check.py              # 全庫
    venv/Scripts/python.exe scripts/health_check.py AAPL NVDA    # 指定幾家
    venv/Scripts/python.exe scripts/health_check.py --report     # 只出摘要，不重跑
"""
from __future__ import annotations

import collections
import json
import pathlib
import sys
import time

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))

# ⚠ Windows 主控台預設 cp950，符號編不出去會讓整支腳本掛掉。專案慣例。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

import config                      # noqa: E402
import data_quality                # noqa: E402
import fetcher_gaap as fg          # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parent.parent
CACHE = ROOT / "local_db" / "filing_cache"
OUT = ROOT / "output" / "_health"

# 表頭列不是模板科目，不列入覆蓋率統計。
_HEADER_ROWS = {"Fiscal Quarter", "Calendar Quarter", "Period End", "",
                "Income Statement", "Balance Sheet", "Cash Flow Statement"}

# ⚠ **模板列與 overflow 列一定要分開統計。** 產出的表裡除了 97 個模板列，
# 還有每家公司自己的科目（overflow，例如 ARLO 的「Separation expense」）。
# 混在一起算覆蓋率的話，overflow 只有一家有值是**正常現象**，卻會把
# 「幾家抓得到」的分母整個拉歪——實測三家就冒出 190 個「列」，而模板只有 97 個。
_TEMPLATE_ROWS = frozenset(
    row[0] for tpl in (fg.IS_TEMPLATE, fg.BS_TEMPLATE, fg.CF_TEMPLATE) for row in tpl)


def _has(value) -> bool:
    """有沒有值。NaN 用 `v != v` 判，不 import numpy。"""
    return value is not None and value != "" and not (
        isinstance(value, float) and value != value)


# 三張表合併在同一張 sheet，區塊標題就是分界線。
# ⚠ 現金流那塊的標題是 **"Cash Flow"**，不是 "Cash Flow Statement"——寫錯的話
# 整個 CF 區塊抓不到，跨表比對會安靜地變成「比了 0 期」（實測踩到）。
_SECTION_TITLES = ("Income Statement", "Balance Sheet", "Cash Flow",
                   "Cash Flow Statement")
_CF_SECTION = ("Cash Flow", "Cash Flow Statement")


def _sectioned_rows(table):
    """`{區塊: {列名: 值}}`。

    ⚠ **一定要按區塊分，不可以用列名當鍵。** `Net Income` 在損益表與現金流量表
    **各有一列**，用名字取會拿 IS 那列當 CF 那列——`diag_celldiff2.py` 的註解
    記過這個坑（2026-08-24 因此憑空生出 3,659 個假異動）。
    """
    out, current = {}, None
    for i, name in enumerate(table.concepts or []):
        if name in _SECTION_TITLES:
            current = name
            out.setdefault(current, {})
            continue
        if current is None or name in _HEADER_ROWS:
            continue
        out[current].setdefault(name, table.values[i] if i < len(table.values) else [])
    return out


def _cross_checks(table) -> dict:
    """跨表一致性與會計恆等式。**自己跟自己對，不依賴任何外部基準。**

    只做定義明確的兩項：

    - **Net Income**：損益表與現金流量表第一行本來就該是同一個數字
    - **會計恆等式**：`Total Assets` 對得上 `Total Liabilities & Equity`

    ⚠ **沒有做「BS 的 Cash vs CF 的 Ending Cash」**——BS 那一列的
    `standard_concept` 是 `CashAndMarketableSecurities`（**含有價證券**），
    跟 CF 的期末現金定義不同，直接比會產生大量假警報。
    """
    sec = _sectioned_rows(table)
    is_rows = sec.get("Income Statement", {})
    bs_rows = sec.get("Balance Sheet", {})
    cf_rows = {}
    for key in _CF_SECTION:
        if sec.get(key):
            cf_rows = sec[key]
            break
    n = len(table.period_ends or table.quarter_labels or [])

    def compare(a, b, tol=1.0):
        """兩列逐期比。回傳 (兩邊都有值的期數, 對不上的期數)。"""
        both = mismatch = 0
        for i in range(n):
            x = a[i] if i < len(a) else None
            y = b[i] if i < len(b) else None
            if not (_has(x) and _has(y)):
                continue
            try:
                x, y = float(x), float(y)
            except (TypeError, ValueError):
                continue
            both += 1
            if abs(x - y) > tol:
                mismatch += 1
        return both, mismatch

    ni_both, ni_bad = compare(is_rows.get("Net Income", []),
                              cf_rows.get("Net Income", []))
    eq_both, eq_bad = compare(bs_rows.get("Total Assets", []),
                              bs_rows.get("Total Liabilities & Equity", []))
    return {
        "net_income_compared": ni_both, "net_income_mismatch": ni_bad,
        "balance_compared": eq_both, "balance_mismatch": eq_bad,
    }


def check_one(ticker: str, identity: str) -> dict:
    """跑一家，回傳這家的體檢結果。走正式路徑，不另寫一套判斷。"""
    started = time.time()
    with fg.collect_gaps() as gaps:
        tables = fg.fetch_gaap_statements(ticker, identity)

    row = {"ticker": ticker, "elapsed": round(time.time() - started, 1),
           "gaps": [{"where": g.where, "kind": g.kind, "exc": g.exc_name}
                    for g in gaps.gaps],
           "sheets": {}}

    for table in tables:
        name = getattr(table, "sheet_name", "")
        if name not in ("Data_Financials(Q)", "Data_Financials(Y)"):
            continue
        report_ = data_quality.assess(table)
        rows = []
        for i, concept in enumerate(table.concepts or []):
            if concept in _HEADER_ROWS:
                continue
            values = table.values[i] if i < len(table.values) else []
            rows.append({"row": concept,
                         "template": concept in _TEMPLATE_ROWS,
                         "filled": sum(1 for v in values if _has(v)),
                         "periods": len(values)})
        row["sheets"][name] = {
            "cross": _cross_checks(table),
            "total_periods": report_.total_periods or 0,
            "empty_rows": report_.empty_but_plausible,
            "holed": len(report_.holed),
            "sporadic": len(report_.sporadic),
            "contradictions": len(report_.contradictions),
            "missing_quarters": sum(g.count for g in report_.missing_quarters),
            "template_mismatch": bool(report_.template_mismatch),
            "rows": rows,
        }
    return row


def build_report() -> int:
    """把落檔的結果彙總成摘要。"""
    files = sorted(OUT.glob("*.json"))
    if not files:
        print("還沒有任何結果，先跑一次 health_check.py")
        return 1
    rows = []
    for path in files:
        try:
            rows.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            continue

    errored = [r["ticker"] for r in rows if r.get("error")]
    rows = [r for r in rows if not r.get("error")]
    print(f"=== 全庫體檢摘要（{len(rows)} 家）===")
    if errored:
        print(f"    另有 {len(errored)} 家跑失敗：{errored[:10]}")
    print()

    # ── 1. 抓取缺漏 ──
    kinds = collections.Counter(g["kind"] for r in rows for g in r["gaps"])
    with_gaps = [r["ticker"] for r in rows if r["gaps"]]
    missing_cur = collections.Counter(
        r["ticker"] for r in rows
        for g in r["gaps"] if g["exc"] == "MissingCurrentPeriod")
    print("【抓取缺漏】")
    print(f"  有缺漏的公司      {len(with_gaps)} / {len(rows)} 家")
    print(f"  缺漏筆數分類      {dict(kinds)}")
    print(f"  當期整期讀不出來  {sum(missing_cur.values())} 筆、{len(missing_cur)} 家"
          f"  {missing_cur.most_common(8)}")
    print()

    # ── 2. 每家的品質判斷（Index 頁標紅用的同一套）──
    for sheet in ("Data_Financials(Q)", "Data_Financials(Y)"):
        have = [r for r in rows if sheet in r["sheets"]]
        if not have:
            continue

        def pick(field):
            return sorted(((r["ticker"], r["sheets"][sheet][field]) for r in have
                           if r["sheets"][sheet][field]), key=lambda x: -x[1])

        mismatch = [r["ticker"] for r in have if r["sheets"][sheet]["template_mismatch"]]
        annual = sheet.endswith("(Y)")
        print(f"【{sheet}】{len(have)} 家"
              + ("   ⚠ Excel 的 Index 頁只判季表，年表這幾個數字僅供參考" if annual else ""))
        print(f"  模板對不上        {len(mismatch):3d} 家  {mismatch[:8]}")
        fields = [("有〔矛盾〕標紅", "contradictions"), ("有〔中間有洞〕", "holed")]
        if not annual:
            # ⚠ **年表不可以算「缺季」。** `missing_quarters()` 是用「相鄰期末日
            # 差幾季」判的，年表相鄰差 365 天＝4 季，於是每兩年報 3 季缺口——
            # AAPL 17 年會報 48 段，**100% 假警報**。正式路徑
            # （`excel_formatter._quality()`）只對 `Data_Financials(Q)` 做判斷，
            # 年表根本不判，所以 Excel 的 Index 頁不會出現這個假警報。
            fields.append(("有缺季", "missing_quarters"))
        for label, field in fields:
            hits = pick(field)
            print(f"  {label:16s}  {len(hits):3d} 家  {hits[:6]}")
        print()

    # ── 2.5 跨表一致性（自己跟自己對，不依賴外部基準）──
    sheet = "Data_Financials(Q)"
    have = [r for r in rows if sheet in r["sheets"] and r["sheets"][sheet].get("cross")]
    if have:
        ni_bad = [(r["ticker"], r["sheets"][sheet]["cross"]["net_income_mismatch"])
                  for r in have if r["sheets"][sheet]["cross"]["net_income_mismatch"]]
        eq_bad = [(r["ticker"], r["sheets"][sheet]["cross"]["balance_mismatch"])
                  for r in have if r["sheets"][sheet]["cross"]["balance_mismatch"]]
        ni_tot = sum(r["sheets"][sheet]["cross"]["net_income_compared"] for r in have)
        eq_tot = sum(r["sheets"][sheet]["cross"]["balance_compared"] for r in have)
        print("【跨表一致性】不依賴外部基準，自己跟自己對")
        print(f"  Net Income（損益表 vs 現金流量表）  比了 {ni_tot} 期，"
              f"對不上 {sum(n for _, n in ni_bad)} 期、{len(ni_bad)} 家")
        print(f"    {sorted(ni_bad, key=lambda x: -x[1])[:8]}")
        print(f"  會計恆等式（總資產 vs 負債+權益）   比了 {eq_tot} 期，"
              f"對不上 {sum(n for _, n in eq_bad)} 期、{len(eq_bad)} 家")
        print(f"    {sorted(eq_bad, key=lambda x: -x[1])[:8]}")
        print()

    # ── 3. 模板列覆蓋率（最重要的一段）──
    sheet = "Data_Financials(Q)"
    have = [r for r in rows if sheet in r["sheets"]]
    if not have:
        return 0
    per_row_cos = collections.Counter()
    per_row_fill = collections.defaultdict(list)
    all_rows = set()
    overflow = collections.Counter()
    for r in have:
        for item in r["sheets"][sheet]["rows"]:
            # ⚠ overflow 列（公司自己的科目）不算進覆蓋率，只計數。
            if not item.get("template"):
                if item["filled"]:
                    overflow[r["ticker"]] += 1
                continue
            all_rows.add(item["row"])
            if item["filled"]:
                per_row_cos[item["row"]] += 1
                per_row_fill[item["row"]].append(
                    item["filled"] / max(item["periods"], 1))

    total = len(have)
    print(f"【模板列覆蓋率】{sheet}，{total} 家、{len(all_rows)} 個模板列"
          f"（overflow 另計：中位數每家 "
          f"{sorted(overflow.values())[len(overflow) // 2] if overflow else 0} 列）")
    print(f"  {'列名':<34s} {'幾家有值':>10s} {'填滿率中位數':>12s}")
    worst = sorted(all_rows, key=lambda n: per_row_cos.get(n, 0))[:25]
    for name in worst:
        count = per_row_cos.get(name, 0)
        fills = sorted(per_row_fill.get(name, []))
        median = fills[len(fills) // 2] if fills else 0.0
        print(f"  {name[:34]:<34s} {count:>6d}/{total:<4d} {median:>11.0%}")
    print()
    good = sum(1 for name in all_rows if per_row_cos.get(name, 0) >= total * 0.85)
    never = [n for n in all_rows if per_row_cos.get(n, 0) == 0]
    print(f"  ≥85% 的公司抓得到的列：{good} / {len(all_rows)}")
    print(f"  一家都抓不到的列：{len(never)}  {sorted(never)[:10]}")
    print()
    print("⚠ 不要把「達標列數」當 KPI——有些列天生就不該人人都有（多數公司沒發")
    print("   特別股、沒有非控制權益、零售業不單獨揭露 R&D）。詳見")
    print("   docs/template-coverage-baseline-2026-08-24.md 第零節。")
    return 0


def main(argv) -> int:
    if "--report" in argv:
        return build_report()
    OUT.mkdir(parents=True, exist_ok=True)
    targets = [a for a in argv if not a.startswith("-")] or \
        sorted(d.name for d in CACHE.iterdir() if d.is_dir())
    identity = config.load_config()["identity"]
    for i, ticker in enumerate(targets, 1):
        path = OUT / f"{ticker}.json"
        if path.exists():
            continue                       # 中斷續跑：已完成的整家跳過
        try:
            row = check_one(ticker, identity)
        except Exception as exc:           # noqa: BLE001 — 一家壞掉不能拖垮整批
            row = {"ticker": ticker, "error": f"{type(exc).__name__}: {exc}",
                   "gaps": [], "sheets": {}}
            print(f"[{i}/{len(targets)}] {ticker:6s} ERROR {type(exc).__name__}",
                  flush=True)
        else:
            stats = row["sheets"].get("Data_Financials(Q)", {})
            print(f"[{i}/{len(targets)}] {ticker:6s} {row['elapsed']:5.1f}s "
                  f"期數={stats.get('total_periods', 0):3d} "
                  f"空列={stats.get('empty_rows', 0):3d} "
                  f"缺漏={len(row['gaps'])}", flush=True)
        path.write_text(json.dumps(row, ensure_ascii=False), encoding="utf-8")
    print()
    return build_report()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
