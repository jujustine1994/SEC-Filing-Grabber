# -*- coding: utf-8 -*-
"""check_fy_labels.py — 財年標籤的內部一致性檢查（G13(b) 階段一，唯讀）。

## 這支在答什麼

**「52/53 週財年制造成的財年錯位，全庫到底有多少家中招？」**

2026-09-20 抽查 15 家時發現 JNJ 與 CDNS 中招（JNJ 3 個財年、CDNS 2 個），
但**全庫 215 家的真正影響面還不知道**。TODO G13(b) 的修法要動核心模型
（`fiscal_quarter_of()` 有 7 個呼叫點、會動到所有公司的期間標籤），
**值不值得動要先看數字**——這支就是那個數字。

## 判準：內部一致性，不依賴任何外部來源

同一個 `FYxxxx` 標籤底下的四季，按**期末日**排序之後，Q 序號就該是 1、2、3、4。
排出來不是這個順序，代表那個標籤底下混了不同財年的資料。JNJ FY2021 實例：

    2021-01-03 → FY2021Q4     ← 1 月，卻標成 Q4
    2021-04-04 → FY2021Q1
    2021-07-04 → FY2021Q2
    2021-10-03 → FY2021Q3

排序後 Q 序號是 4、1、2、3。`2021-01-03` 其實是 **FY2020** 的結束日
（那份 10-K filed 2021-02-22 報的是 2020 財年）。

⚠ **刻意不去問「正確答案是什麼」**，只問「現有標籤自己有沒有矛盾」。
要判斷 `2021-01-03` 到底該叫 FY2020 還 FY2021，得知道那家公司的財年命名慣例
（NVDA 的 FY2026 結束於 2026-01，JNJ 的 2021-01-03 卻屬於 FY2020——**慣例因
公司而異，沒有通則**）。內部一致性檢查繞開這個問題：不管慣例是什麼，
同一個財年的四季在時間上一定是連續的。

## 兩階段：離線篩選 → pipeline 精確驗證

**⚠ 離線捷徑單獨用會漏判，這是實測踩到的坑。** 第一版只讀快取欄名，對 JNJ
回報「0 個錯位」，但完整 pipeline 量到 3 個——因為 JNJ 的 `2021-01-03` 在快取
裡標的是 `(FY)`（它是年報期末日），正式路徑會把年報的 FY 欄**合成成 Q4** 放進
季表，離線捷徑沒有複製那段邏輯。同理它報的 COP／DE／EXC／GS／JCI／NOC／ONTO
其實是「同一天有多個期間欄」的重複問題，不是財年錯位。

所以分兩階段：

- **階段一（離線、快）**：篩「財年結束月會跳月」的公司——那是 52/53 週制的
  特徵，只有這些才可能中招。全庫 215 家掃完約 20 秒。
- **階段二（走 pipeline、慢但準）**：對候選跑真正的 `fetch_gaap_statements()`，
  拿它算出來的 `quarter_labels` 與 `period_ends` 判定。**結論不會跟正式路徑
  分岔，因為用的就是正式路徑。**

用法：

    venv/Scripts/python.exe scripts/check_fy_labels.py                 # 階段一：全庫篩選
    venv/Scripts/python.exe scripts/check_fy_labels.py --verify        # 兩階段都跑
    venv/Scripts/python.exe scripts/check_fy_labels.py --verify JNJ    # 只驗指定幾家
"""
from __future__ import annotations

import collections
import json
import pathlib
import re
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))

# ⚠ Windows 主控台預設 cp950，`⚠`／`✅` 編不出去會讓整支腳本掛掉。
# errors="replace" 是保險。專案慣例，見 cli.py 的 _force_utf8_io()。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

from fiscal_input import fiscal_quarter_of, fy_start_month   # noqa: E402

CACHE = pathlib.Path(__file__).resolve().parent.parent / "local_db" / "filing_cache"
FY_COL = re.compile(r"^(\d{4})-(\d{2})-(\d{2})\s+\(FY\)")
Q_COL = re.compile(r"^(\d{4}-\d{2}-\d{2})\s+\(Q\d\)")
LABEL = re.compile(r"^FY(\d{4})Q([1-4])$")


def _cols(entry, key):
    payload = (entry.get("dataframes") or {}).get(key)
    return payload["data"]["columns"] if payload else []


def scan(ticker: str) -> dict:
    """一家公司：財年結束月 + 所有季末日 → 標籤矛盾清單。"""
    directory = CACHE / ticker
    fy_months, q_ends = collections.Counter(), set()
    for path in directory.glob("*.json"):
        try:
            entry = json.load(open(path, encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not entry.get("has_financials"):
            continue
        for col in _cols(entry, "income_statement"):
            m = FY_COL.match(str(col))
            if m:
                fy_months[int(m.group(2))] += 1
            m = Q_COL.match(str(col))
            if m:
                q_ends.add(m.group(1))
    if not fy_months or not q_ends:
        return {"ticker": ticker, "skipped": "沒有年報或季報的期間欄"}

    # 多數決：52/53 週制的公司財年結束月會在兩個月份之間跳，取常見的那個
    # ——這正是問題的根源，所以這裡故意用跟正式路徑一樣的「單一月份」模型。
    fy_end_month = fy_months.most_common(1)[0][0]
    start_month = fy_start_month(fy_end_month)

    # ⚠ **這裡刻意不做錯位判定**——離線資料少了「年報 FY 欄合成 Q4」那一段，
    # 判出來的結果跟正式路徑不一致（見模組說明）。這一階段只回報特徵。
    return {"ticker": ticker, "fy_end_month": fy_end_month,
            "months_seen": dict(fy_months), "quarters": len(q_ends),
            "is_week_based": len(fy_months) > 1}


def verify(ticker: str, identity: str) -> list:
    """階段二：走真正的 pipeline，回報這家有哪些財年錯位。

    判準：同一個 `FYxxxx` 標籤底下的四季，按期末日排序後 Q 序號該是 1、2、3、4。
    **用的就是正式路徑算出來的 `quarter_labels`**，所以結論不會分岔。
    """
    import fetcher_gaap

    tables = fetcher_gaap.fetch_gaap_statements(ticker, identity, max_filings=40)
    pairs = []
    for table in tables:
        if table.sheet_name != "Data_Financials(Q)":
            continue
        pairs = [(e, l) for e, l in zip(table.period_ends, table.quarter_labels)
                 if e and LABEL.match(str(l))]
    by_fy = collections.defaultdict(list)
    for end, label in pairs:
        m = LABEL.match(str(label))
        by_fy[m.group(1)].append((end, int(m.group(2))))

    bad = []
    for fy, items in sorted(by_fy.items()):
        if len(items) < 4:
            continue                      # 不完整的財年無從判斷順序
        order = sorted(items)
        if [q for _, q in order] != [1, 2, 3, 4]:
            bad.append((fy, order))
    return bad


def main(argv) -> int:
    do_verify = "--verify" in argv
    argv = [a for a in argv if a != "--verify"]
    targets = argv or sorted(d.name for d in CACHE.iterdir() if d.is_dir())

    # ── 階段一：離線篩選 ──
    rows = [scan(t) for t in targets]
    scanned = [r for r in rows if not r.get("skipped")]
    candidates = [r["ticker"] for r in scanned if r["is_week_based"]]
    print(f"階段一（離線）：掃 {len(scanned)} 家")
    print(f"  財年結束月會跳月的（52/53 週制特徵）：{len(candidates)} 家")
    print(f"  {candidates}")
    if skipped := [r["ticker"] for r in rows if r.get("skipped")]:
        print(f"  跳過（資料不足）{len(skipped)} 家")

    if not do_verify:
        print()
        print("加 --verify 跑階段二（走正式 pipeline 精確判定，每家約 10~30 秒）")
        return 0

    # ── 階段二：走 pipeline ──
    import config
    identity = config.load_config()["identity"]
    print()
    print(f"階段二（pipeline）：驗證 {len(candidates)} 家…")
    bad_cos, total_bad = {}, 0
    for i, ticker in enumerate(candidates, 1):
        try:
            bad = verify(ticker, identity)
        except Exception as exc:                      # noqa: BLE001
            print(f"  [{i}/{len(candidates)}] {ticker:6s} ERROR "
                  f"{type(exc).__name__}", flush=True)
            continue
        if bad:
            bad_cos[ticker] = bad
            total_bad += len(bad)
            print(f"  [{i}/{len(candidates)}] {ticker:6s} ✗ {len(bad)} 個財年錯位",
                  flush=True)
            for fy, order in bad[:2]:
                cells = "  ".join(f"{e}=Q{q}" for e, q in order)
                print(f"            FY{fy}: {cells}")
        else:
            print(f"  [{i}/{len(candidates)}] {ticker:6s} OK", flush=True)

    print()
    print(f"=== 結論 ===")
    print(f"  52/53 週制候選   {len(candidates)} 家")
    print(f"  真的錯位的公司   {len(bad_cos)} 家  {sorted(bad_cos)}")
    print(f"  錯位的財年總數   {total_bad} 個")
    return 1 if total_bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
