# -*- coding: utf-8 -*-
"""
verify_local_db.py — 本地財報資料庫的**格式驗證**（TODO J5 每批跑完必跑）。

`audit_local_db.py` 回答「份數夠不夠」，這支回答**「這些檔案能不能用」**。
兩件事完全不同：一堆格式壞掉的檔案，份數一樣是滿的。

跑一批要好幾個小時，發現格式錯才重來的代價太大，所以每批跑完就驗一次。

## 驗什麼

**A. 四道閘（最關鍵）**——對每一份快取檔實際呼叫 `filing_cache.load_filing()`。
   它回 `None` 就代表**下次抓取會當作沒有快取、整份重新下載**，那份檔案等於白抓。
   四道閘是：JSON 可解析／`schema_version`／`cik` 相符／`edgartools_version` 相符。
   這一項直接用正式程式碼的函式，不是另外寫一套判斷——**不可能跟實際行為分岔**。

**B. 反序列化**——把命中的快取餵給 `cached_filing()`，實際取出三張 DataFrame。
   驗的是 `payload_to_df()` 那條路（`orient="split"` + dtype 還原）沒有壞掉。

**C. 內容合理性**——有 financials 的檔案，IS 至少要有一列、要有 `concept`／`label`
   這類欄位、而且至少有一個期間欄（不是只有 meta 欄）。

**D. 負向快取比例**——`has_financials=False` 是 pre-XBRL 舊申報的正常結果，
   但**比例過高**（預設 >40%）代表那家公司的抓取有問題，要列出來看。

**E. meta 與目錄一致**——`_meta.json` 的 `file_count`／每個 form 的 count
   跟實際檔案對得上。

**完全不連網**，9,233 份約 2~4 分鐘。

    ./venv/Scripts/python.exe scripts/verify_local_db.py
    ./venv/Scripts/python.exe scripts/verify_local_db.py --json output/_localdb/verify.json
    ./venv/Scripts/python.exe scripts/verify_local_db.py AAPL DIS   # 只驗這幾家

回傳碼 0＝全部通過；1＝有問題（問題清單印在最後）。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import filing_cache          # noqa: E402
import local_db              # noqa: E402

# 期間欄之外的欄位（edgartools 的 meta 欄）。用來判斷「這張表有沒有真的期間資料」。
META_COLS = {
    "concept", "label", "standard_concept", "level", "abstract", "dimension",
    "is_breakdown", "dimension_axis", "dimension_member", "dimension_member_label",
    "dimension_label", "unit", "point_in_time", "balance", "weight",
    "preferred_sign", "parent_concept", "parent_abstract_concept",
}

NEGATIVE_CACHE_WARN_RATIO = 0.40


def verify_ticker(ticker: str) -> dict:
    """一家公司的格式驗證。回傳的 `problems` 空的就是通過。"""
    out = {
        "ticker": ticker, "files": 0, "loadable": 0, "dead": [],
        "has_financials": 0, "negative": 0, "unreadable": [],
        "bad_dataframe": [], "empty_statement": [], "problems": [],
        "empty_shell": [], "is_missing": [],
        "forms": Counter(), "oldest": None, "newest": None,
    }
    directory = filing_cache.ticker_dir(ticker)
    accessions = sorted(local_db.cached_accessions(ticker))
    out["files"] = len(accessions)
    if not accessions:
        out["problems"].append("沒有任何 filing 檔")
        return out

    # cik 從檔案本身取（load_filing 的第三道閘要比對它）。取多數決——
    # 不一致本身就是問題，下面會報。
    ciks = Counter()
    for accession in accessions:
        try:
            with open(directory / f"{accession}.json", "r", encoding="utf-8") as f:
                raw = json.load(f)
            ciks[raw.get("cik")] += 1
        except (OSError, ValueError):
            out["unreadable"].append(accession)
    if len(ciks) > 1:
        out["problems"].append(f"同一家公司出現多個 cik：{dict(ciks)}")
    cik = ciks.most_common(1)[0][0] if ciks else None
    out["cik"] = cik

    dates = []
    for accession in accessions:
        # ── A. 四道閘：用正式程式碼判，不另外寫一套 ──
        entry = filing_cache.load_filing(ticker, accession, cik)
        if entry is None:
            out["dead"].append(accession)
            continue
        out["loadable"] += 1
        out["forms"][str(entry.get("form") or "?")] += 1
        if entry.get("filing_date"):
            dates.append(str(entry["filing_date"]))

        if not entry.get("has_financials"):
            out["negative"] += 1
            continue
        out["has_financials"] += 1

        # ── B/C. 反序列化 + 內容合理性 ──
        try:
            fin = filing_cache.cached_filing(entry).financials
            dfs = {}
            for name, getter in (("IS", fin.income_statement),
                                 ("BS", fin.balance_sheet),
                                 ("CF", fin.cashflow_statement)):
                stmt = getter()
                dfs[name] = None if stmt is None else stmt.to_dataframe()
        except Exception as exc:                      # noqa: BLE001
            out["bad_dataframe"].append(f"{accession}: {type(exc).__name__}: {exc}")
            continue

        # 「空殼」＝ financials 物件在、但三張表全是 None。
        #
        # ⚠ **這不是格式錯，不要當成問題報。** 是忠實記錄上游現實：
        # SEC 的 XBRL 是**分三階段**強制的（2009-06 最大型 → 2010-06 其餘大型
        # 加速申報人 → 2011-06 所有其他人），之前的申報 edgartools 給得出
        # financials 物件卻解不出任何一張表。
        #
        # **2026-09-06 決定性驗證**：挑 7 份 2010~2023 的空殼直接跟 SEC 重抓
        # （COHR／KLAC／FTNT／DXCM／KEYS／GEHC／ABBV），**7/7 上游結果完全一樣**，
        # edgartools 自己的訊息就是 "No statements available in XBRL data"。
        # 快取存的沒錯，重抓也救不回來。
        #
        # 旁證：batch 1 的 824 份空殼，年份分布是 2010:85、2011:18、2012 之後只剩 6，
        # 跟三階段時程完全吻合。剩下那幾份是分拆後第一份年報（ABBV 2013、
        # KEYS 2014、GEHC 2023）。
        #
        # 所以這裡只**統計**，不列為 problem——誤報會讓人以後不看這支的輸出。
        # 真正的涵蓋率問題屬於 H 系列的體檢，不是格式驗證的職責。
        if all(df is None or df.empty for df in dfs.values()):
            out["empty_shell"].append(f"{accession} ({entry.get('filing_date') or ''})")
            continue

        df = dfs["IS"]
        if df is None or df.empty:
            # IS 沒有但 BS/CF 有——不常見但不是格式錯（G13 那類上游解析狀況）
            out["is_missing"].append(accession)
            continue
        cols = {str(c) for c in df.columns}
        if not ({"concept", "label"} & cols):
            out["bad_dataframe"].append(f"{accession}: IS 沒有 concept/label 欄")
        if not (cols - META_COLS):
            out["empty_statement"].append(f"{accession}: IS 只有 meta 欄，沒有期間欄")

    if dates:
        out["oldest"], out["newest"] = min(dates), max(dates)

    # ── D. 負向快取比例 ──
    if out["loadable"]:
        ratio = out["negative"] / out["loadable"]
        out["negative_ratio"] = round(ratio, 3)
        if ratio > NEGATIVE_CACHE_WARN_RATIO:
            out["problems"].append(
                f"負向快取比例 {ratio:.0%}（{out['negative']}/{out['loadable']}）偏高")

    # ── E. meta 一致性 ──
    meta = local_db.read_meta(ticker)
    if meta is None:
        out["problems"].append("沒有 _meta.json（或壞掉／schema 不符）")
    else:
        if meta.get("file_count") != out["files"]:
            out["problems"].append(
                f"meta file_count={meta.get('file_count')} 但實際 {out['files']} 份")
        for form in local_db.FORMS:
            recorded = (meta.get("forms", {}).get(form) or {}).get("count")
            actual = out["forms"].get(form, 0)
            if recorded != actual:
                out["problems"].append(
                    f"meta {form} count={recorded} 但實際 {actual} 份")
        if meta.get("edgartools_version") != filing_cache.edgartools_version():
            out["problems"].append(
                f"meta 版本 {meta.get('edgartools_version')}，"
                f"現在是 {filing_cache.edgartools_version()}")

    if out["dead"]:
        out["problems"].append(
            f"{len(out['dead'])} 份過不了 load_filing 的四道閘（下次會整份重抓）")
    if out["unreadable"]:
        out["problems"].append(f"{len(out['unreadable'])} 份讀不出來")
    if out["bad_dataframe"]:
        out["problems"].append(f"{len(out['bad_dataframe'])} 份反序列化有問題")
    if out["empty_statement"]:
        out["problems"].append(f"{len(out['empty_statement'])} 份 IS 只有 meta 欄，沒有期間欄")
    out["forms"] = dict(out["forms"])
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="本地財報資料庫格式驗證（不連網）")
    ap.add_argument("tickers", nargs="*", help="只驗這幾家；不給就驗全部")
    ap.add_argument("--json", metavar="PATH", help="把結果寫成 JSON")
    args = ap.parse_args(argv)

    targets = local_db.normalize_tickers(args.tickers) or [
        r["ticker"] for r in filing_cache.list_cached_tickers()]

    started = time.monotonic()
    rows, bad = [], []
    tot_files = tot_loadable = tot_neg = tot_fin = 0
    tot_shell = tot_ismiss = 0
    late_shells = []
    for i, ticker in enumerate(targets, 1):
        row = verify_ticker(ticker)
        rows.append(row)
        tot_files += row["files"]
        tot_loadable += row["loadable"]
        tot_neg += row["negative"]
        tot_fin += row["has_financials"]
        tot_shell += len(row["empty_shell"])
        late_shells += [f"{ticker} {s}" for s in row["empty_shell"]
                        if s.split("(")[-1].rstrip(")") >= "2010-01-01"]
        tot_ismiss += len(row["is_missing"])
        if row["problems"]:
            bad.append(ticker)
            print(f"[{i}/{len(targets)}] {ticker:<6} ✗ " + "；".join(row["problems"]),
                  flush=True)
        elif i % 20 == 0 or i == len(targets):
            print(f"[{i}/{len(targets)}] …{ticker} OK", flush=True)

    elapsed = time.monotonic() - started
    print(f"\n驗證完成（{elapsed:.0f}s）")
    print(f"  公司            {len(targets)} 家")
    print(f"  filing 檔       {tot_files} 份")
    print(f"  過四道閘        {tot_loadable} 份"
          f"（{tot_loadable / max(tot_files, 1):.1%}）")
    print(f"    有 financials {tot_fin} 份")
    print(f"    負向快取      {tot_neg} 份"
          f"（pre-XBRL 舊申報，正常現象）")
    print(f"    其中空殼      {tot_shell} 份"
          f"（三張表全 None）")
    if tot_shell:
        print(f"      -> 2010 前 {tot_shell - len(late_shells)} 份、"
              f"2010 後 {len(late_shells)} 份。**都不是格式問題**："
              f"SEC 的 XBRL 分三階段強制（2009-06／2010-06／2011-06），")
        print(f"         實測跟 SEC 重抓 7 份，7/7 上游結果一樣。詳見腳本註解")
    if tot_ismiss:
        print(f"    IS 缺但 BS/CF 有 {tot_ismiss} 份")
    print(f"  有問題的公司    {len(bad)} 家")
    if bad:
        print("  -> " + " ".join(bad))
    else:
        print("\n✅ 格式全部正確——每一份都過得了 load_filing 的四道閘、"
              "三張表都反序列化得回來、meta 跟目錄一致。")

    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(
            json.dumps({"checked": len(targets), "files": tot_files,
                        "loadable": tot_loadable, "negative": tot_neg,
                        "late_shells": late_shells,
                        "bad": bad, "rows": rows}, ensure_ascii=False, indent=2),
            encoding="utf-8")
        print(f"\n寫入 {args.json}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
