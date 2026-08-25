"""驗證 `cli.py press-release` 的 `fiscal_label`（TODO D4 後半，方案 B+）。

15 家 × 8 季 = 120 份 Item 2.02 8-K，檢查兩件事：

1. **期末日抓取率**：`_period_end_from_tables()` 有沒有抓到 `period_end`
2. **偏移一致性**：列清單階段的 `label` 與下載後算出的 `fiscal_label` 的差，
   **B5（2026-08-25）之後應該全部是 0**——列清單改走零下載規則（發布日 +
   EDGAR `fiscal_year_end` 回推名目季末）之後，兩條路本來就該對齊

第 2 點是重點，而且它同時是 B5 的端對端驗收與 `fiscal_label` 的回歸：
`fiscal_label` 那一側完全沒動，出現非 0 偏移就代表其中一條路壞了。
B5 之前每家是各自的常數偏移（-3 ~ +1，由財年結束月決定），舊的期望值留在
`LEGACY_OFFSETS` 當對照。

**零 AI**，只打 SEC EDGAR。約 3 分鐘。

    ./venv/Scripts/python.exe scripts/verify_8k_fiscal_labels.py
"""
import re
import sys
from collections import Counter
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT / "src"))
import cli
from press_release_tables import parse_tables

cli._force_utf8_io()                    # 主控台是 cp950，中文與符號要先轉 UTF-8
# 沒填過進階設定（config.json 不存在）時，第一個參數當 identity 用。
IDENTITY = sys.argv[1] if len(sys.argv) > 1 else cli.resolve_identity(None)
TICKERS = ["AAPL", "AVGO", "ARLO", "AMD", "INTC", "NOW", "COST", "MSFT",
           "MU", "PANW", "WDC", "ORCL", "QRVO", "CRM", "NVDA"]
# B5 之前的偏移（列清單標籤 − 實際財期，單位：季）。留著當對照，不再是期望值。
LEGACY_OFFSETS = {"AAPL": 0, "AVGO": 0, "ARLO": 1, "AMD": 1, "INTC": 1, "NOW": 1,
                  "COST": -1, "MSFT": -1, "MU": -1, "PANW": -1, "WDC": -1,
                  "ORCL": -2, "QRVO": -2, "CRM": -3, "NVDA": -3}

# B5 之後：列清單的 label 與 fiscal_label 必須完全一致，每一家都是 0。
EXPECTED = {ticker: 0 for ticker in LEGACY_OFFSETS}


def ordinal(label):
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", label or "")
    return int(m.group(1)) * 4 + int(m.group(2)) - 1 if m else None


total = empty = 0
for ticker in TICKERS:
    fye = cli._fiscal_year_end(ticker, IDENTITY)
    fym = cli._fy_end_month_from_mmdd(fye)
    filings = cli._earnings_filings(ticker=ticker, identity=IDENTITY,
                                    start_year=None, end_year=None, max_filings=8,
                                    fiscal_year_end=fye)
    offsets, misses = Counter(), []
    for label, filing in filings:
        total += 1
        try:
            html = cli._press_release_html(filing)
        except Exception as exc:
            misses.append(f"{label}:{type(exc).__name__}")
            empty += 1
            continue
        if not html:
            misses.append(f"{label}:no_pr")
            continue
        tables = parse_tables(html)
        end = cli._period_end_from_tables(
            tables, str(filing.filing_date) or str(filing.period_of_report))
        new = cli._fiscal_label(end, fym)
        if not end or not new:
            misses.append(f"{label}:no_date")
            empty += 1
            continue
        a, b = ordinal(label), ordinal(new)
        offsets[a - b if a is not None and b is not None else "?"] += 1
    exp = EXPECTED[ticker]
    ok = "OK " if list(offsets) == [exp] else "!! "
    print(f"{ok}{ticker:5} fye={str(fye):>4} 偏移={dict(offsets)} 期望={exp} "
          f"（B5 前是 {LEGACY_OFFSETS[ticker]}）問題={misses}", flush=True)

print(f"\n共 {total} 份，抓不到期末日 {empty} 份")
