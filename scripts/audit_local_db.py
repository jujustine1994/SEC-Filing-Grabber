# -*- coding: utf-8 -*-
"""
audit_local_db.py — 本地財報資料庫的體檢（TODO J5 的前置與收尾工具）。

回答一個問題：**快取裡這些公司，哪些已經是完整且最新的？**

「完整且最新」＝三個條件同時成立（跟 `local_db.plan_ticker()` 判斷「整家跳過」
的邏輯是同一套，所以這支的結論跟 `update-db` 實際會不會跳過**保證一致**）：

  1. 10-Q 與 10-K **都到底了**（`reached_bottom` 不是 null）
  2. SEC 上**沒有新的 filing** 是本地沒有的
  3. 快取是**現在這個 edgartools 版本**解出來的

只做「列 filing 清單」這一種網路請求（每家兩次，很便宜），**不下載任何 filing**。
34 家約一分鐘。

    ./venv/Scripts/python.exe scripts/audit_local_db.py              # 體檢快取裡所有公司
    ./venv/Scripts/python.exe scripts/audit_local_db.py --json a.json
    ./venv/Scripts/python.exe scripts/audit_local_db.py AAPL META    # 只看這幾家
    ./venv/Scripts/python.exe scripts/audit_local_db.py --plan-next 100 \\
        --universe output/_hintsweep_201/tickers_joined.txt

`--plan-next N` 會從 universe 檔案裡挑出「還沒抓過的前 N 家」並印出來，
給 J5 分批跑用。**照字母序挑**，所以同一份 universe 每次挑出來的都一樣，
分批之間不會漏掉也不會重複。

回傳碼：0＝全部完整；1＝有不完整的（清單印在最後，可直接餵給 `update-db`）。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import filing_cache          # noqa: E402
import local_db              # noqa: E402
from config import load_config   # noqa: E402


def audit_one(ticker: str, identity: str) -> dict:
    """一家公司的體檢結果。任何例外都收進 `error`，不中斷整批。"""
    row = {"ticker": ticker, "cached": 0, "complete": False, "reason": "",
           "new_filings": 0, "forms": {}, "version_ok": None, "error": ""}
    row["cached"] = len(local_db.cached_accessions(ticker))
    try:
        listings, _cik = local_db._default_list_filings(ticker, identity)
    except Exception as exc:                       # noqa: BLE001
        row["error"] = f"{type(exc).__name__}: {exc}"
        row["reason"] = "列清單失敗"
        return row

    meta = local_db.load_meta(ticker)
    current = filing_cache.edgartools_version()
    cached = local_db.cached_accessions(ticker)
    version_ok = bool(current) and (
        not cached or (meta or {}).get("edgartools_version") == current)
    plan = local_db.plan_ticker(listings, cached, version_ok=version_ok)

    row["version_ok"] = version_ok
    row["new_filings"] = plan["new_count"]
    row["forms"] = {f: plan["forms"][f]["reached_bottom"] for f in local_db.FORMS}
    row["available"] = {f: plan["forms"][f]["available"] for f in local_db.FORMS}
    row["complete"] = plan["skip"]
    row["meta_version"] = (meta or {}).get("edgartools_version")

    if plan["skip"]:
        row["reason"] = "完整"
    elif not version_ok:
        row["reason"] = f"版本不符（快取 {row['meta_version']}，現在 {current}）"
    elif plan["new_count"]:
        row["reason"] = f"有 {plan['new_count']} 份新 filing 沒抓"
    else:
        unfinished = [f for f in local_db.FORMS
                      if plan["forms"][f]["reached_bottom"] is None]
        row["reason"] = f"還沒到底：{'／'.join(unfinished)}"
    return row


def plan_next(universe_path: str, count: int) -> list[str]:
    """從 universe 挑「還沒抓過的前 N 家」。照字母序，所以分批可重現。"""
    raw = Path(universe_path).read_text(encoding="utf-8")
    universe = local_db.normalize_tickers(raw.replace(",", " ").split())
    cached = {r["ticker"] for r in filing_cache.list_cached_tickers()}
    todo = sorted(t for t in universe if t not in cached)
    return todo[:count]


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="本地財報資料庫體檢")
    ap.add_argument("tickers", nargs="*", help="只看這幾家；不給就看快取裡全部")
    ap.add_argument("--identity", help="SEC EDGAR Identity（預設讀 config.json）")
    ap.add_argument("--json", metavar="PATH", help="把結果寫成 JSON")
    ap.add_argument("--plan-next", type=int, metavar="N",
                    help="從 universe 挑出還沒抓過的前 N 家並印出")
    ap.add_argument("--universe", metavar="PATH",
                    default="output/_hintsweep_201/tickers_joined.txt",
                    help="universe 清單檔（預設 201 家那份）")
    args = ap.parse_args(argv)

    identity = (args.identity or load_config().get("identity") or "").strip()
    if not identity:
        print("沒有 SEC EDGAR Identity。用 --identity 指定，或先在 GUI 填一次。",
              file=sys.stderr)
        return 2

    targets = local_db.normalize_tickers(args.tickers) or [
        r["ticker"] for r in filing_cache.list_cached_tickers()]

    started = time.monotonic()
    rows = []
    for i, ticker in enumerate(targets, 1):
        row = audit_one(ticker, identity)
        rows.append(row)
        mark = "OK  " if row["complete"] else "需要更新"
        print(f"[{i}/{len(targets)}] {ticker:<6} {row['cached']:>3} 份  "
              f"{mark}  {row['reason']}", flush=True)

    incomplete = [r["ticker"] for r in rows if not r["complete"]]
    complete = [r["ticker"] for r in rows if r["complete"]]
    print(f"\n體檢完成（{time.monotonic() - started:.0f}s）："
          f"完整 {len(complete)} 家、需要更新 {len(incomplete)} 家")
    if incomplete:
        print("需要更新的：" + " ".join(incomplete))

    payload = {"checked": len(rows), "complete": complete,
               "incomplete": incomplete, "rows": rows}

    if args.plan_next:
        nxt = plan_next(args.universe, args.plan_next)
        payload["next_batch"] = nxt
        payload["universe"] = args.universe
        print(f"\n下一批（universe 裡還沒抓過的前 {args.plan_next} 家，"
              f"實際挑到 {len(nxt)} 家）：")
        print(" ".join(nxt))

    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
        print(f"\n寫入 {args.json}")

    return 1 if incomplete else 0


if __name__ == "__main__":
    sys.exit(main())
