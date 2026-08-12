"""
smoke_test_10.py - batch live smoke test, 10 companies GAAP fetch, check key rows.
Usage: python scripts/smoke_test_10.py
"""

import sys
import io
import traceback
from pathlib import Path

# Force UTF-8 output on Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# 讓 import 找到專案根目錄的模組
_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_ROOT / "src"))

from config import load_config
from fetcher_gaap import fetch_gaap_statements

# ── 測試名單 ─────────────────────────────────────────────────────────────────
# 涵蓋：Dec FY / Sep FY / Jun FY / Jan FY、老公司（AMD）、CF overflow 公司（COHR）
TICKERS = ["AAPL", "MSFT", "TSLA", "AMD", "NVDA", "GOOGL", "META", "WMT", "COHR", "AMZN"]

# 要檢查的欄位（Std Name，對應 Data_Financials(Q) A 欄）
CHECK_ROWS_IS = [
    "Revenue",
    "Gross Profit",
    "Operating Income",
    "Net Income",
]
CHECK_ROWS_CF = [
    "Operating Cash Flow",
    "Capex",
    "Free Cash Flow",
]

# ── 顏色（Windows Terminal / PowerShell 支援 ANSI）────────────────────────────
RED   = "\033[91m"
GRN   = "\033[92m"
YLW   = "\033[93m"
RST   = "\033[0m"

def check(val) -> str:
    if val is None:
        return f"{RED}NONE{RST}"
    return f"{GRN}OK  {RST}"


def run_smoke_test():
    cfg = load_config()
    identity = cfg.get("identity", "")
    if not identity:
        print(f"{RED}ERROR: identity 未設定，請先在進階設定填入 SEC EDGAR Identity{RST}")
        sys.exit(1)

    results = []

    for ticker in TICKERS:
        print(f"\n{'─'*60}")
        print(f"Fetching {ticker} ...")
        try:
            tables = fetch_gaap_statements(
                ticker,
                identity,
                max_filings=cfg.get("max_filings", 80),
                ai_config=cfg.get("ai", {}),
            )
        except Exception as e:
            print(f"{RED}[ERROR] {ticker}: {e}{RST}")
            results.append({
                "ticker": ticker,
                "status": "ERROR",
                "error": str(e),
                "labels": [],
                "checks": {},
            })
            continue

        # 找 Data_Financials(Q)
        q_table = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
        if q_table is None:
            print(f"{YLW}[WARN] {ticker}: 找不到 Data_Financials(Q){RST}")
            results.append({
                "ticker": ticker,
                "status": "NO_TABLE",
                "error": "Data_Financials(Q) missing",
                "labels": [],
                "checks": {},
            })
            continue

        labels  = q_table.quarter_labels
        concepts = q_table.concepts
        values   = q_table.values   # list[list], values[row_idx][col_idx]

        # 最新季度 = 最後一欄
        latest_label = labels[-1] if labels else "N/A"
        latest_col   = len(labels) - 1

        # 建 concept → row_idx 的 mapping
        concept_idx = {c: i for i, c in enumerate(concepts)}

        checks = {}
        all_rows = CHECK_ROWS_IS + CHECK_ROWS_CF
        for row_name in all_rows:
            idx = concept_idx.get(row_name)
            if idx is None:
                checks[row_name] = None   # 欄位根本不存在於 concepts
            else:
                row_vals = values[idx]
                val = row_vals[latest_col] if latest_col < len(row_vals) else None
                checks[row_name] = val

        print(f"  Latest quarter: {latest_label}  (total {len(labels)} qtrs)")
        for row_name in all_rows:
            val = checks.get(row_name)
            print(f"  {row_name:<25} {check(val)}  {'' if val is None else f'{val:,.0f}'}")

        results.append({
            "ticker": ticker,
            "status": "OK",
            "error": None,
            "labels": [latest_label],
            "checks": checks,
        })

    # ── 彙總表 ────────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("  Summary: missing key rows per company")
    print(f"{'='*60}")
    header = f"{'Ticker':<8} {'Latest':<12} " + " ".join(f"{r[:6]:<7}" for r in CHECK_ROWS_IS + CHECK_ROWS_CF)
    print(header)
    print("─" * len(header))

    all_rows = CHECK_ROWS_IS + CHECK_ROWS_CF
    issue_tickers = []
    for r in results:
        ticker  = r["ticker"]
        status  = r["status"]
        if status == "ERROR":
            print(f"{ticker:<8} {RED}ERROR: {r['error'][:40]}{RST}")
            issue_tickers.append(ticker)
            continue
        if status == "NO_TABLE":
            print(f"{ticker:<8} {YLW}NO_TABLE{RST}")
            issue_tickers.append(ticker)
            continue

        latest = r["labels"][0] if r["labels"] else "N/A"
        checks = r["checks"]
        has_issue = any(v is None for v in checks.values())
        if has_issue:
            issue_tickers.append(ticker)

        row_str = ""
        for row_name in all_rows:
            val = checks.get(row_name)
            if val is None:
                row_str += f"{RED}{'NONE':<7}{RST}"
            else:
                row_str += f"{GRN}{'OK':<7}{RST}"

        print(f"{ticker:<8} {latest:<12} {row_str}")

    print(f"{'='*60}")
    if issue_tickers:
        print(f"{YLW}Issues found: {', '.join(issue_tickers)}{RST}")
    else:
        print(f"{GRN}All companies: key rows all present OK{RST}")


if __name__ == "__main__":
    run_smoke_test()
