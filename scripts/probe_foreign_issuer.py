# -*- coding: utf-8 -*-
"""probe_foreign_issuer.py — 外國私人發行人到底抓不抓得到（TODO D9）。

    ./venv/Scripts/python.exe scripts/probe_foreign_issuer.py TM SONY HMC NMR

回答三個問題，全部用實測不靠猜：

  1. **有沒有向 SEC 申報**——ADR Level I 走 Rule 12g3-2(b) 豁免，根本沒資料；
     Level II／III 才必須提交 20-F／6-K
  2. **用哪個會計準則**——`us-gaap` 的話現有科目對照表直接能用，`ifrs-full`
     才要另建一套（這是 D9 真正的工程量落點，不是「外國公司」這件事本身）
  3. **季報（6-K）解不解析得出來**——6-K 是大雜燴，同樣是 6-K 可能是財報也
     可能是股東會通知。實測 ARM：33 份裡只有 9 份有財報，判準是
     **`R*.htm` 檔數量 > 1**（有財報 61~72 個、沒有的剛好 1 個，33/33 分離）

⚠ **不要抽樣一份就下結論**。2026-09-18 查 ARM 時第一次剛好抽到沒財報的那份，
差點做出「6-K 沒有財報資料」的錯誤結論，掃完 33 份才看到真相。
"""
from __future__ import annotations

import collections
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

# ⚠ Windows 主控台預設 cp950，`⚠`／`✅` 這類符號編不出去會讓整支腳本掛掉
# （`UnicodeEncodeError`）。`errors="replace"` 是保險：真的編不出去的字印成
# `?`，不要讓一個符號炸掉整趟輸出。專案慣例，見 `cli.py` 的 `_force_utf8_io()`。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass


from config import load_config              # noqa: E402
from edgar import Company, set_identity     # noqa: E402

MAX_6K = 12        # 每家掃幾份 6-K。ARM 實測 33 份裡 9 份有財報，12 份夠看出比例


def _namespaces(filing) -> tuple[int, dict]:
    """(facts 總數, 命名空間分布)。拿不到回 (0, {})。"""
    try:
        x = filing.xbrl()
        df = x.facts.to_dataframe() if x is not None else None
        if df is None:
            return 0, {}
        col = "concept" if "concept" in df.columns else df.columns[0]
        ns = collections.Counter(str(v).split("_")[0].split(":")[0] for v in df[col])
        return len(df), dict(ns.most_common(4))
    except Exception:                          # noqa: BLE001
        return 0, {}


def _r_file_count(filing) -> int:
    """`R*.htm` 的數量——SEC 財報檢視器從 XBRL 三表產生的，有三表才會有一堆。"""
    try:
        return sum(1 for a in filing.attachments
                   if str(getattr(a, "document_type", "")) == "HTML"
                   and str(getattr(a, "document", "")).startswith("R"))
    except Exception:                          # noqa: BLE001
        return 0


def _can_parse(filing) -> str:
    """真的走一次既有的解析路徑，回傳三表的列數或失敗原因。"""
    try:
        fin = getattr(filing.obj(), "financials", None)
        if fin is None:
            return "financials=None"
        shapes = []
        for name in ("income_statement", "balance_sheet", "cashflow_statement"):
            st = getattr(fin, name)()
            df = st.to_dataframe() if st is not None else None
            shapes.append(str(df.shape[0]) if df is not None else "-")
        return "／".join(shapes) + " 列"
    except Exception as exc:                   # noqa: BLE001
        return f"{type(exc).__name__}"


def probe(ticker: str) -> None:
    print(f"\n{'=' * 64}\n{ticker}")
    try:
        c = Company(ticker)
    except Exception as exc:                   # noqa: BLE001
        print(f"  ❌ SEC 查無此代號（{type(exc).__name__}）"
              f"——多半是 ADR Level I，走 Rule 12g3-2(b) 豁免")
        return
    print(f"  {c.name}｜CIK {c.cik}")

    counts = collections.Counter(str(f.form) for f in c.get_filings())
    key = {k: v for k, v in counts.items() if k in ("10-K", "10-Q", "20-F", "6-K", "40-F")}
    print(f"  申報類型：{key or '（沒有財報類申報）'}")

    if counts.get("10-Q"):
        print("  ✅ 有 10-Q，現有流程本來就抓得到，不屬於 D9")
        return

    for form in ("20-F", "40-F"):
        fl = list(c.get_filings(form=form))[:1]
        if not fl:
            continue
        n, ns = _namespaces(fl[0])
        gaap = "us-gaap" in ns
        print(f"  {form} 最新一份（{fl[0].filing_date}）：{n} 筆　{ns}")
        print(f"     → {'✅ us-gaap，現有科目對照表可用' if gaap else '⚠ IFRS，要另建對照表'}")
        print(f"     → 解析：{_can_parse(fl[0])}")

    sixk = list(c.get_filings(form="6-K"))[:MAX_6K]
    if not sixk:
        print("  6-K：沒有")
        return
    rich = []
    for f in sixk:
        nr = _r_file_count(f)
        if nr > 1:
            rich.append((f, nr))
    print(f"  6-K：掃了 {len(sixk)} 份，其中 {len(rich)} 份的 R 檔 > 1（＝含財報）")
    for f, nr in rich[:2]:
        n, ns = _namespaces(f)
        print(f"     {f.filing_date}　R檔 {nr}　{n} 筆　{ns}")
        print(f"        → 解析：{_can_parse(f)}")


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        print(__doc__)
        return 2
    set_identity(load_config()["identity"])
    for ticker in argv:
        probe(ticker.strip().upper())
    return 0


if __name__ == "__main__":
    sys.exit(main())
