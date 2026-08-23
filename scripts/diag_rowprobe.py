"""diag_rowprobe.py — 某個模板列在多家公司的命中情況，並列出所有可能的候選。

用法：
    venv/Scripts/python.exe scripts/diag_rowprobe.py "<模板列名>" <候選正則> <TICKERS逗號分隔> [幾份filing]

對每家公司印出「現行 `_match_is_row()` 命中了什麼」與「dataframe 裡還有哪些
長得像的列」。判斷一個模板列的對照要不要改，看這支的輸出最快。
"""
import sys, re, pathlib
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import config
from edgar import Company, set_identity
set_identity(config.load_config()["identity"])
import fetcher_gaap as fg

ROW, PAT, TKS = sys.argv[1], re.compile(sys.argv[2], re.I), sys.argv[3].split(",")
N = int(sys.argv[4]) if len(sys.argv) > 4 else 1
SRC = {r[0]: r for T in (fg.IS_TEMPLATE, fg.BS_TEMPLATE, fg.CF_TEMPLATE) for r in T}
name, std, fb, source, match, hint, lbl_fb = SRC[ROW]
print(f"### {ROW}: std={std!r} fallback={fb!r} match={match} hint={hint!r} src={source}\n")
for tk in TKS:
    try:
        filings = list(Company(tk).get_filings(form="10-Q", amendments=False))[:N]
    except Exception as e:
        print(tk, "ERR", e); continue
    with fg._parse_cache_scope():
        for f in filings:
            try:
                fin = fg._financials_of(fg._filing_obj(f))
                stmt = {"IS": lambda: fin.income_statement(), "BS": lambda: fin.balance_sheet(),
                        "CF": lambda: fin.cashflow_statement()}[source]()
                df = stmt.to_dataframe()
            except Exception as e:
                print(f"{tk} {f.filing_date} ERR {type(e).__name__}: {e}"); continue
            idx = fg._match_is_row(df, std, fb, label_fallback=lbl_fb,
                                   match=match, label_hint=hint)
            got = f"[{idx}] {df.loc[idx,'concept']!r} / {df.loc[idx,'label']!r}" if idx is not None else "*** NO MATCH ***"
            print(f"{tk} {f.filing_date} -> {got}")
            m = fg._consolidated_mask(df)
            for i, r in df[m].iterrows():
                if PAT.search(f"{r.get('concept')} {r.get('label')}"):
                    print(f"    cand [{i}] concept={r.get('concept')!r} std={r.get('standard_concept')!r} label={r.get('label')!r}")
