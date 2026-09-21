# -*- coding: utf-8 -*-
"""classify_holes.py — 把「中間有洞」分成真缺與正常（唯讀，不改任何東西）。

## 這支在答什麼

全庫體檢顯示 **209/215 家有〔中間有洞〕**，那是 Excel 的 Index 頁會標紅給
使用者看的東西。但**沒人知道那裡面有多少是真缺**——如果 97% 的公司都標紅，
這個警告就等於沒用。

## ⚠ 為什麼不用 companyfacts

`docs/PITFALLS.md` 陷阱一：companyfacts 彙整的是這家公司歷來**所有 filing**
的所有 fact，包含「後續財報把去年同期當比較欄報」的那些。我們只從那一期
自己的 filing 取，所以凡是「當期空殼、後續補報」都會被誤判成真缺——實測讓
「真缺 57%」那個數字整個不可信。

## 改用的判準：本地快取裡「有沒有任何一份 filing 報過這一期」

同一個期末日會出現在好幾份 filing 裡（當期那份，加上後續幾份的比較欄）。
所以問題可以改寫成：

    這一期、這個科目，**本地快取的任何一份 filing** 裡有值嗎？

- **有** → 資料在我們手上，是比對層沒撈出來 → **真缺，可以補**
- **沒有** → 公司從來沒報過 → **正常，不該標紅**

這個判準的好處：
- **完全離線**，不打 SEC
- 用的是**我們自己已經抓下來的 filing**，不是外部彙整
- 科目對照直接用 `IS/BS/CF_TEMPLATE` 的 `std_concept`，跟正式路徑同一套

⚠ **「真缺」是上界，不是實數**（2026-09-22 實測修正）。這支只比
`standard_concept`，而模板的比對規則還有 `fallback_suffix`／`label_hint`／
`match` 等過濾條件。**同一個正規化名稱可能是完全不同的科目**——NEM 2015-06-30
的損益表裡 `CostOfGoodsAndServicesSold` 的 label 是 **'Exploration'（探勘費用）**，
正式路徑正確地沒把它當成 `Cost of Revenue`，這支卻會判成「資料在手上、沒撈出來」。

所以這支的定位是**篩出值得人工排查的候選**，不是「這些格子可以直接補」。
判斷單一案例請用 `scripts/diag_rowprobe.py`（它會列出所有長得像的候選）。

⚠ **它也不回答「補進去對不對」**，只回答「資料在不在手上」。真要補，還要決定
用哪一份 filing 的值（當期那份 vs 後續比較欄，遇到重編財報會不一樣）——
那是產品判斷，不在這支的範圍。

    venv/Scripts/python.exe scripts/classify_holes.py            # 全庫
    venv/Scripts/python.exe scripts/classify_holes.py LHX APH    # 指定幾家
"""
from __future__ import annotations

import collections
import json
import pathlib
import re
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

import config                      # noqa: E402
import data_quality                # noqa: E402
import fetcher_gaap as fg          # noqa: E402
import filing_cache                # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parent.parent
CACHE = ROOT / "local_db" / "filing_cache"
OUT = ROOT / "output" / "_holes"

DATE = re.compile(r"^(\d{4}-\d{2}-\d{2})(?:\s+\((\w+)\))?")

# 模板列 → (std_concept, 來源表)。直接用正式路徑的對照表，不另寫一套。
_ROW_SPEC: dict[str, tuple[str, str]] = {}
for _tpl, _src in ((fg.IS_TEMPLATE, "income_statement"),
                   (fg.BS_TEMPLATE, "balance_sheet"),
                   (fg.CF_TEMPLATE, "cashflow_statement")):
    for _row in _tpl:
        # 同名列（Net Income 在 IS 與 CF 都有）以先出現的為準——這支只用來
        # 判斷「資料在不在」，兩張表都找得到就算在。
        _ROW_SPEC.setdefault(_row[0], (_row[1], _src))


def _has(v) -> bool:
    return v is not None and v != "" and not (isinstance(v, float) and v != v)


def build_index(ticker: str) -> tuple[dict, dict]:
    """掃這家所有快取 filing 的三張表，回 (季度索引, 年度索引)。

    同一個期末日會出現在好幾份 filing（當期那份 ＋ 後續幾份的比較欄），
    **任何一份有值就算有**——這支問的是「資料在不在我們手上」。

    ⚠ **季度欄與年度欄一定要分開。** 季表的洞要的是**單季值**，年度欄
    （`(FY)`）的值是整年的，不能拿來說「這一季的資料在手上」——實測 ARLO
    的 `Acquisitions` 在 `FY2018Q4` 是空的，快取裡找到的卻是
    `'2018-12-31 (FY)'`（年度）。混在一起會高估真缺。
    """
    quarterly: dict[str, set] = collections.defaultdict(set)
    annual: dict[str, set] = collections.defaultdict(set)
    for path in (CACHE / ticker).glob("*.json"):
        if not filing_cache.ACCESSION_RE.match(path.stem):
            continue
        try:
            entry = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not entry.get("has_financials"):
            continue
        for key in ("income_statement", "balance_sheet", "cashflow_statement"):
            payload = (entry.get("dataframes") or {}).get(key)
            if not payload:
                continue
            df = filing_cache.payload_to_df(payload)
            if df is None or "standard_concept" not in df.columns:
                continue
            period_cols = []
            for col in df.columns:
                m = DATE.match(str(col))
                if m:
                    period_cols.append((col, m.group(1),
                                        (m.group(2) or "").upper()))
            if not period_cols:
                continue
            for _, row in df.iterrows():
                std = str(row.get("standard_concept") or "")
                if not std or std == "nan":
                    continue
                for col, end, kind in period_cols:
                    if not _has(row.get(col)):
                        continue
                    (annual if kind == "FY" else quarterly)[end].add(std)
    return quarterly, annual


def holes_of(table):
    """[(列名, 期末日)] —— 首末有值之間的洞。`holed_rows()` 不記位置。"""
    sparse = {sp.period_end for sp in data_quality.sparse_periods(table)}
    ends = list(table.period_ends or [])
    out = []
    for name, vals in data_quality._rows(table):
        idx = [i for i, v in enumerate(vals) if _has(v)]
        if len(idx) < 2:
            continue
        for i in range(idx[0], idx[-1] + 1):
            if _has(vals[i]) or i >= len(ends) or not ends[i]:
                continue
            if ends[i] in sparse:
                continue           # 整期稀疏，歸另一類，不算這一列的洞
            out.append((name, ends[i]))
    return out


def classify(ticker: str, identity: str) -> dict:
    tables = fg.fetch_gaap_statements(ticker, identity)
    table = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
    if table is None:
        return {"ticker": ticker, "error": "no quarterly table"}

    quarterly, annual = build_index(ticker)
    real = normal = unknown = annual_only = 0
    per_row = collections.Counter()
    samples = []
    for name, end in holes_of(table):
        spec = _ROW_SPEC.get(name)
        if not spec:
            unknown += 1           # overflow 列或模板外的列，沒有對照
            continue
        std = spec[0]
        if not std:
            unknown += 1
            continue
        if std in quarterly.get(end, ()):
            real += 1                      # 季度欄裡有 → 資料真的在手上
            per_row[name] += 1
            if len(samples) < 5:
                samples.append(f"{name}@{end}")
        elif std in annual.get(end, ()):
            annual_only += 1               # 只有年度值，不等於這一季的值
        else:
            normal += 1
    return {"ticker": ticker, "real": real, "normal": normal,
            "annual_only": annual_only,
            "unknown": unknown, "per_row": dict(per_row.most_common(10)),
            "samples": samples}


def main(argv) -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    targets = [a for a in argv if not a.startswith("-")] or \
        sorted(d.name for d in CACHE.iterdir() if d.is_dir())
    identity = config.load_config()["identity"]
    for i, ticker in enumerate(targets, 1):
        path = OUT / f"{ticker}.json"
        if path.exists():
            continue
        try:
            row = classify(ticker, identity)
        except Exception as exc:               # noqa: BLE001
            row = {"ticker": ticker, "error": f"{type(exc).__name__}: {exc}"}
            print(f"[{i}/{len(targets)}] {ticker:6s} ERROR {type(exc).__name__}",
                  flush=True)
        else:
            if not row.get("error"):
                tot = row["real"] + row["normal"] + row["unknown"]
                print(f"[{i}/{len(targets)}] {ticker:6s} 洞 {tot:4d} → "
                      f"真缺 {row['real']:4d} / 正常 {row['normal']:4d} / "
                      f"無對照 {row['unknown']:3d}", flush=True)
        path.write_text(json.dumps(row, ensure_ascii=False), encoding="utf-8")

    # ── 摘要 ──
    rows = []
    for path in sorted(OUT.glob("*.json")):
        try:
            rows.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            continue
    rows = [r for r in rows if not r.get("error")]
    if not rows:
        return 1
    real = sum(r["real"] for r in rows)
    normal = sum(r["normal"] for r in rows)
    annual_only = sum(r.get("annual_only", 0) for r in rows)
    unknown = sum(r["unknown"] for r in rows)
    total = real + normal + annual_only + unknown or 1
    print()
    print(f"=== 「中間有洞」的分類（{len(rows)} 家）===")
    print(f"  洞總數                        {total}")
    print(f"  真缺（季度資料在手上、沒撈出）{real:6d}  ({100*real/total:.1f}%)")
    print(f"  只有年度值（不等於這一季）    {annual_only:6d}  ({100*annual_only/total:.1f}%)")
    print(f"  正常（公司從來沒報過）        {normal:6d}  ({100*normal/total:.1f}%)")
    print(f"  無對照（overflow 等）         {unknown:6d}  ({100*unknown/total:.1f}%)")
    print()
    worst = sorted(rows, key=lambda r: -r["real"])[:15]
    print("真缺最多的 15 家：")
    for r in worst:
        if r["real"]:
            print(f"  {r['ticker']:6s} {r['real']:4d}  {list(r['per_row'])[:4]}")
    print()
    agg = collections.Counter()
    for r in rows:
        agg.update(r["per_row"])
    print("真缺最多的 15 個科目：")
    for name, n in agg.most_common(15):
        print(f"  {name[:40]:40s} {n:5d}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
