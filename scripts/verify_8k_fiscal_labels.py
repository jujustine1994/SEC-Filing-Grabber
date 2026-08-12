"""驗證 `cli.py press-release` 的 `fiscal_label`（TODO D4 後半，方案 B+）。

15 家 × 8 季 = 120 份 Item 2.02 8-K，檢查兩件事：

1. **期末日抓取率**：`_period_end_from_tables()` 有沒有抓到 `period_end`
2. **偏移一致性**：新標籤與舊標籤（發布日換算）的差，同一家應該是**常數**，
   且等於 `docs/8k-period-off-by-one.md` 用新聞稿內文獨立推出的偏移

第 2 點是重點：偏移由財年結束月決定，同一家 8 季全部一樣。出現混雜就代表
期末日抓錯了某幾季。

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
# 報告表格的偏移（現行標籤 − 實際財期，單位：季）
EXPECTED = {"AAPL": 0, "AVGO": 0, "ARLO": 1, "AMD": 1, "INTC": 1, "NOW": 1,
            "COST": -1, "MSFT": -1, "MU": -1, "PANW": -1, "WDC": -1,
            "ORCL": -2, "QRVO": -2, "CRM": -3, "NVDA": -3}


def ordinal(label):
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", label or "")
    return int(m.group(1)) * 4 + int(m.group(2)) - 1 if m else None


total = empty = 0
for ticker in TICKERS:
    fym = cli._fy_end_month(ticker, IDENTITY)
    filings = cli._earnings_filings(ticker=ticker, identity=IDENTITY,
                                    start_year=None, end_year=None, max_filings=8)
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
    print(f"{ok}{ticker:5} fym={str(fym):>4} 偏移={dict(offsets)} 期望={exp} "
          f"問題={misses}", flush=True)

print(f"\n共 {total} 份，抓不到期末日 {empty} 份")
