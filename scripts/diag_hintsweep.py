"""diag_hintsweep.py — 掃出「`label_hint` 太窄把正確答案濾掉」的模板列。

用法：
    venv/Scripts/python.exe scripts/diag_hintsweep.py <TICKERS逗號分隔>

對每個有 hint 的模板列，比較「有 hint」與「沒 hint」的命中結果，列出被 hint
殺掉的案例。**2026-08-23 的 H3 就是靠這支發現 hint 是最大的系統性成因**
（`Cash Taxes Paid` 少 14 家、`Deferred Revenue, current` 少 12 家…）。

⚠ 有些 hint 是**有必要的**（擋掉現金流量表最下面的租賃補充揭露列），
不要看到被殺就拿掉，要逐條看它擋的是什麼。
"""
import sys, pathlib
from collections import defaultdict
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import config
from edgar import Company, set_identity
set_identity(config.load_config()["identity"])
import fetcher_gaap as fg

TKS = sys.argv[1].split(",")
killed = defaultdict(list)
for tk in TKS:
    try: f = list(Company(tk).get_filings(form="10-Q", amendments=False))[:1][0]
    except Exception as e: print(tk, "ERR", e); continue
    with fg._parse_cache_scope():
        fin = fg._financials_of(fg._filing_obj(f))
        dfs = {}
        for key, g in (("IS","income_statement"),("BS","balance_sheet"),("CF","cashflow_statement")):
            try: dfs[key] = getattr(fin, g)().to_dataframe()
            except Exception: pass
        for T in (fg.IS_TEMPLATE, fg.BS_TEMPLATE, fg.CF_TEMPLATE):
            for name, std, fb, src, match, hint, label_fb in T:
                if not hint or src not in dfs: continue
                df = dfs[src]
                with_hint = fg._match_is_row(df, std, fb, label_fb, match=match, label_hint=hint)
                without   = fg._match_is_row(df, std, fb, label_fb, match=match, label_hint=None)
                if with_hint is None and without is not None:
                    killed[(name, hint)].append(f"{tk}:{df.loc[without,'concept']}|{df.loc[without,'label']}")
                elif with_hint is not None and without is not None and with_hint != without:
                    killed[(name, hint)].append(f"{tk}:DIFF hint={df.loc[with_hint,'concept']} vs {df.loc[without,'concept']}")
for (name, hint), hits in sorted(killed.items(), key=lambda x: -len(x[1])):
    print(f"\n{name}  hint={hint!r}  killed {len(hits)}/{len(TKS)}")
    for h in hits: print("   ", h)
