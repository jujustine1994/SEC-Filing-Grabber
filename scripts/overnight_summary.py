# -*- coding: utf-8 -*-
"""overnight_summary.py — 整夜更新跑完的收尾摘要（`overnight_update_db.ps1` 呼叫）。

    ./venv/Scripts/python.exe scripts/overnight_summary.py <輸出目錄> <時間戳>

讀第二輪各段的 `--json`，回答「早上起來只想知道的那幾件事」：第二輪跳過了幾家
（愈多愈好＝第一輯就抓齊了）、哪幾家失敗、哪幾家有缺漏、資料庫現在多大。

**獨立成檔而不是塞在 PowerShell 的 here-string 裡**：那樣要同時應付 PowerShell
的 `$` 插值與 Python 的 f-string，2026-09-18 實測踩到語法錯誤，而且 here-string
裡的 Python 沒辦法單獨測試。
"""
from __future__ import annotations

import json
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


import filing_cache          # noqa: E402
import local_db              # noqa: E402

# 這幾家是 20-F 外國申報人（只有年報、沒有 10-Q），這個工具抓的是 10-K/10-Q。
# 它們失敗是**預期的**，不要讓人早上看到紅字以為系統壞了。
KNOWN_20F = {"ASML", "STM", "IFX"}


def _missing_chunks(out_dir: Path, stamp: str, pass_name: str,
                    found) -> list[tuple[str, list[str]]]:
    """manifest 上有、但沒有產出 `--json` 的段落 → [(段號, 那段的 ticker)]。

    `--json` 是 `cli.py` 跑完才寫的，所以沒有它就代表那個 process 沒跑完——
    最可能是被系統因記憶體不足砍掉（2026-09-06 實測過，分段就是為了這個）。
    """
    manifest_path = out_dir / f"{stamp}_{pass_name}_manifest.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return []                        # 沒有 manifest 就不做這項檢查
    done = {p.name.rsplit("chunk", 1)[-1].split(".")[0].lstrip("0") or "0"
            for p in found}
    return [(num, tickers) for num, tickers in sorted(manifest.items(),
                                                      key=lambda kv: int(kv[0]))
            if num.lstrip("0") not in done]


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if len(argv) < 2:
        print("用法：overnight_summary.py <輸出目錄> <時間戳>", file=sys.stderr)
        return 2
    out_dir, stamp = Path(argv[0]), argv[1]

    updated = skipped = failed = 0
    fails: list[tuple[str, str]] = []
    gaps: list[str] = []
    pass_name = "pass2"
    chunks = sorted(out_dir.glob(f"{stamp}_pass2_chunk*.json"))
    if not chunks:                       # 第二輪被跳過（-SkipSecondPass）就看第一輪
        pass_name = "pass1"
        chunks = sorted(out_dir.glob(f"{stamp}_pass1_chunk*.json"))

    # ⚠ 整段 process 被系統 OOM 砍掉時不會寫 `--json`，只看 json 會**靜默少報
    # 一整段**（8 家憑空消失）。分段的存在本來就是為了緩解 OOM，所以這是實際
    # 會發生的情境，不是理論風險。比對 manifest 才看得出來。
    crashed = _missing_chunks(out_dir, stamp, pass_name, chunks)

    for path in chunks:
        try:
            # utf-8-sig 不是多餘的：正常流程是 `cli.py` 用 Python 寫的（無 BOM），
            # 但這些檔案也可能被 PowerShell 或手動編輯過而帶上 BOM，那樣用
            # `utf-8` 讀會丟 ValueError → 被下面吞掉 → **整段統計靜默歸零**。
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError):
            continue
        updated += data.get("updated", 0)
        skipped += data.get("skipped", 0)
        failed += data.get("failed", 0)
        gaps += data.get("gap_tickers") or []
        fails += [(r["ticker"], r.get("error") or "")
                  for r in data.get("results", []) if r.get("status") == "failed"]

    print(f"最後一輪：更新 {updated} 家、跳過 {skipped} 家、失敗 {failed} 家")
    print("（「跳過」愈多愈好——代表前一輪就抓齊了，這正是收斂的樣子）")

    if crashed:
        lost = [t for _, tickers in crashed for t in tickers]
        print(f"\n⚠⚠ 有 {len(crashed)} 段整段沒有產出，{len(lost)} 家沒跑到 ⚠⚠")
        print("   （那個 process 沒跑完就結束了，最可能是記憶體不足被系統砍掉）")
        for num, tickers in crashed:
            print(f"    第 {num} 段：{', '.join(tickers)}")
        print("   → 這幾家不會出現在上面的統計裡。單獨補跑：")
        print("     venv\\Scripts\\python.exe src\\cli.py update-db "
              + " ".join(lost[:8]) + (" …" if len(lost) > 8 else ""))
        print("   → 或直接再雙擊一次 過夜更新資料庫.bat，它會從沒抓完的接下去")

    if fails:
        expected = [t for t, _ in fails if t in KNOWN_20F]
        real = [(t, e) for t, e in fails if t not in KNOWN_20F]
        if expected:
            print(f"\n預期內的失敗（20-F 外國申報人，沒有 10-Q，不是 bug）："
                  f"{', '.join(sorted(expected))}")
        if real:
            print(f"\n⚠ 要看的失敗 {len(real)} 家：")
            for ticker, err in real:
                print(f"    {ticker}：{err[:120]}")
            print("    → 單獨重跑就好："
                  "venv\\Scripts\\python.exe src\\cli.py update-db "
                  + " ".join(t for t, _ in real[:6]))

    if gaps:
        # D11：連續大量抓取時 SEC 會偶發失敗、靜默少格，單獨重抓通常就補回來
        print(f"\n⚠ 有抓取缺漏（D11）、建議單獨重跑：{', '.join(sorted(set(gaps)))}")

    rows = local_db.overview_rows({})
    summary = local_db.overview_summary(rows)
    print(f"\n資料庫現況：{summary['companies']} 家 / {summary['filings']} 份 / "
          f"{summary['size_bytes'] / 1e9:.2f} GB")
    print(f"位置：{filing_cache.cache_root()}")

    incomplete = [r["ticker"] for r in rows if r["bottom"] != local_db.BOTTOM_YES]
    if incomplete:
        print(f"\n還沒抓到底的 {len(incomplete)} 家："
              f"{', '.join(incomplete[:15])}"
              + ("…" if len(incomplete) > 15 else ""))
        print("  → 再跑一次 過夜更新資料庫.bat 就會從這些接下去")
    return 0


if __name__ == "__main__":
    sys.exit(main())
