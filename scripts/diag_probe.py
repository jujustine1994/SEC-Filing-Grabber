"""diag_probe.py — 印出某家公司某張報表裡符合條件的列（最常用的排查工具）。

用法：
    venv/Scripts/python.exe scripts/diag_probe.py <TICKER> <is|bs|cf> <正則> [幾份filing]

輸出每一列的 `concept` / `standard_concept` / `label` 與當期數值。
**ARCHITECTURE「三步排查順序」的第 2、3 步就靠這支**——先確認那一列在不在報表
dataframe 裡，再看 matcher 為什麼沒命中。2026-08-23 的 H3 修復全部從這裡開始。
"""
import sys, re, pathlib
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import config
from edgar import Company, set_identity
set_identity(config.load_config()["identity"])
import fetcher_gaap as fg

tk, which, pat = sys.argv[1], sys.argv[2], re.compile(sys.argv[3], re.I)
n = int(sys.argv[4]) if len(sys.argv) > 4 else 2
c = Company(tk)
filings = list(c.get_filings(form="10-Q", amendments=False))[:n]
with fg._parse_cache_scope():
    for f in filings:
        obj = fg._filing_obj(f)
        fin = fg._financials_of(obj)
        if fin is None: print(tk, f.filing_date, "no financials"); continue
        stmt = {"is": fin.income_statement, "bs": fin.balance_sheet, "cf": fin.cashflow_statement}[which]()
        if stmt is None: print(tk, f.filing_date, "no stmt"); continue
        df = stmt.to_dataframe()
        mask = fg._consolidated_mask(df)
        print(f"=== {tk} {f.filing_date} rows={len(df)} cons={mask.sum()} cols={[x for x in df.columns if x not in fg.META_COLS]}")
        for i, r in df[mask].iterrows():
            blob = f"{r.get('concept')} {r.get('standard_concept')} {r.get('label')}"
            if pat.search(str(blob)):
                vals = [r.get(x) for x in df.columns if x not in fg.META_COLS][:1]
                print(f"  [{i}] concept={r.get('concept')!r} std={r.get('standard_concept')!r} label={r.get('label')!r} v={vals}")
