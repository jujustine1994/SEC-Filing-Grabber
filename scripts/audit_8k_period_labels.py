"""audit_8k_period_labels.py — 量化 Item 2.02 8-K 季度標籤的 off-by-one（TODO D4 前半）。

問題：`fetcher_nongaap._period_to_quarter_label()` 是拿 8-K 的 `period_of_report`
去分季，但 EDGAR 那欄放的是**發布/事件日**，不是財報所屬的財期結束日。
實測 INTC `20260723` 被標成 `FY2026Q3`，但那份新聞稿報的是 FY2026 Q2。

方法（**純文字比對，不呼叫任何 AI**）：
  1. 從 EDGAR 列出每家的 Item 2.02 8-K（listing metadata，不下載文件）
  2. 下載新聞稿原文，用 regex 抓公司自己寫的財期
     — 「second quarter of fiscal year 2026」之類的措辭
     — 「three months ended June 28, 2026」的期末日
  3. 跟現行標籤比對，統計差了幾季

同時查另一個同根因的問題：同一個日曆季內發布兩份 Item 2.02 8-K 時，
`_list_earnings_filings()` 的 dedupe「保留最舊那份」會把後一份直接丟掉。

用法：
    ./venv/Scripts/python.exe scripts/audit_8k_period_labels.py [--quarters 8] [輸出.json]

原文會快取到 `--cache-dir`（預設 `8k_audit_html/`），改分析規則可重跑不必重抓。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8")

import config
from edgar import Company, set_identity

from fetcher_nongaap import _period_to_quarter_label

# 跨產業、跨市值、含非 12 月結算（AAPL 9 月、COST 8 月、WDC/MU/LRCX/QRVO 6~7 月、
# INTC 12 月但財季錯開、CRM/PANW/NOW 1~7 月）。非 12 月結算才看得出標籤是
# 「日曆季 vs 財季」的差異還是真的整季錯位。
TICKERS = [
    "AAPL", "MSFT", "NVDA", "AVGO", "ORCL", "CRM", "PANW", "NOW",
    "INTC", "AMD", "MU", "LRCX", "QRVO", "WDC", "COST", "ARLO",
]

_ORDINALS = {"first": 1, "second": 2, "third": 3, "fourth": 4}

# 「second quarter of fiscal year 2026」「fourth quarter and fiscal year 2025」
_FQ_RE = re.compile(
    r"\b(first|second|third|fourth)[\s-]+quarter\b[^.]{0,40}?"
    r"\bfiscal(?:\s+year)?\s+(\d{4})",
    re.IGNORECASE,
)
# 財年寫在前面的講法：「fiscal year 2026 second quarter」
_FQ_RE_ALT = re.compile(
    r"\bfiscal(?:\s+year)?\s+(\d{4})\s+(first|second|third|fourth)\s+quarter\b",
    re.IGNORECASE,
)
# 12 月結算的公司不寫 fiscal：「Third Quarter 2024 Results」。這條較寬鬆，
# 放在 fiscal 版之後才試，避免把「compared to the third quarter of 2023」抓進來。
# INTC / AMD 寫「Third-Quarter 2024」，連字號不是空白。
_FQ_RE_PLAIN = re.compile(
    r"\b(first|second|third|fourth)[\s-]+quarter\b[^.]{0,30}?\b(20\d\d)\b",
    re.IGNORECASE,
)
# 期末日。措辭差異極大——實測就有「quarter ended」「52-week fiscal year ended」
# 「fiscal 2024, which ended」「fiscal fourth quarter and fiscal year ended」，
# 所以只認 `ended <日期>` 這個共同結構，不去猜前面接什麼。
_ENDED_RE = re.compile(
    r"\bended\s+([A-Z][a-z]+\s+\d{1,2},\s*\d{4})",
    re.IGNORECASE,
)


def _press_release_text(filing) -> str | None:
    try:
        ek = filing.obj()
    except Exception as exc:
        print(f"      obj() 失敗: {type(exc).__name__}", file=sys.stderr)
        return None
    for pr in (getattr(ek, "press_releases", None) or []):
        try:
            text = pr.text()
        except Exception:
            continue
        if text:
            return unicodedata.normalize("NFKC", text)
    return None


def stated_fiscal_quarter(text: str) -> str | None:
    """公司自己在新聞稿寫的財期，如 `FY2026Q2`。抓不到回 None。

    只認開頭 8,000 字元。標題與第一段一定會寫「reports second quarter fiscal
    2026 results」；再往後就會混進去年同期與財測的措辭，抓到的是別的季。
    """
    head = text[:8000]
    m = _FQ_RE.search(head)
    if m:
        return f"FY{m.group(2)}Q{_ORDINALS[m.group(1).lower()]}"
    m = _FQ_RE_ALT.search(head)
    if m:
        return f"FY{m.group(1)}Q{_ORDINALS[m.group(2).lower()]}"
    m = _FQ_RE_PLAIN.search(head)
    if m:
        return f"FY{m.group(2)}Q{_ORDINALS[m.group(1).lower()]}"
    return None


_MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], 1)}


def stated_period_end(text: str, not_after: str = "") -> str | None:
    """新聞稿寫的**本期**期末日，正規化成 YYYY-MM-DD。抓不到回 None。

    取所有「... ended <日期>」裡**不晚於發布日、且最晚**的那個。

    兩個限制條件缺一不可：
      - 不能取第一個：多數新聞稿把去年同期排在前面，抓到的是比較期間。
      - 只取最晚的也不行：ARLO 的 Q3 新聞稿裡有「fiscal year ended
        December 31」這種講法，日期比本季期末晚，推出來的財年結束月會變成
        3 月（實際是 12 月），整家公司的比對就全錯。發布日之後的期末日
        一定不是本期。
    """
    dates = []
    for m in _ENDED_RE.finditer(text[:20000]):
        parsed = _parse_date(m.group(1))
        if parsed and (not not_after or parsed <= not_after):
            dates.append(parsed)
    return max(dates) if dates else None


def _parse_date(text: str) -> str | None:
    m = re.match(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", text.strip())
    if m is None:
        return None
    month = _MONTHS.get(m.group(1).lower())
    return f"{m.group(3)}-{month:02d}-{int(m.group(2)):02d}" if month else None


def fy_end_month_from(period_end: str, stated: str) -> int | None:
    """由「期末日 + 公司自稱的財季」反推財年結束月。

    比另外下載 10-K 判斷（`_detect_fy_end_month`）便宜得多，而且用同一份新聞稿
    的兩個獨立欄位互相印證：8 季全部推出同一個月份才採信。
    NVDA Q1 FY2027 結束在 4 月 → (4−3) = 1 月；AAPL Q1 結束在 12 月 → 9 月。
    """
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", stated or "")
    if m is None or not period_end:
        return None
    end_month = int(period_end[5:7])
    return ((end_month - 3 * int(m.group(2)) - 1) % 12) + 1


def gaap_style_label(period_end: str, fy_end_month: int) -> str | None:
    """用 `fetcher_gaap._col_to_quarter_label()` 的慣例替期末日產標籤。

    這才是有意義的比較基準：`Data_NonGAAP` 與 `Data_Q` 在同一本活頁簿裡，
    兩張表的 `FY2026Q1` 必須指同一段期間，否則使用者橫著看就是錯的。
    """
    if not period_end or not fy_end_month:
        return None
    year, month = int(period_end[:4]), int(period_end[5:7])
    quarter = ((month - fy_end_month - 1) % 12) // 3 + 1
    if fy_end_month < 12 and month > fy_end_month:
        year += 1
    return f"FY{year}Q{quarter}"


def _quarter_ordinal(label: str) -> int | None:
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", label or "")
    return int(m.group(1)) * 4 + int(m.group(2)) - 1 if m else None


def audit(tickers: list[str], quarters: int, cache_dir: Path) -> dict:
    cache_dir.mkdir(parents=True, exist_ok=True)
    per_filing: list[dict] = []
    dupes: list[dict] = []

    for i, ticker in enumerate(tickers, 1):
        print(f"[{i}/{len(tickers)}] {ticker}", flush=True)
        try:
            company = Company(ticker)
            filings = [f for f in company.get_filings(form="8-K", amendments=False)
                       if "2.02" in str(getattr(f, "items", "") or "")][:quarters]
        except Exception as exc:
            print(f"      清單失敗: {type(exc).__name__}", file=sys.stderr)
            continue

        # 依**發布日由舊到新**掃，才能重現 _list_earnings_filings() 的 dedupe
        # ——它是 `for label, filing in reversed(candidates)`，同一標籤保留最舊
        # 那份。方向弄反的話 kept/dropped 會整個對調，結論剛好相反。
        seen: dict[str, str] = {}
        for filing in reversed(filings):
            period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
            if len(period) < 8:
                continue
            label = _period_to_quarter_label(period)
            accession = str(filing.accession_no)

            # 同一標籤兩份 8-K → dedupe 會丟掉其中一份
            if label in seen:
                dupes.append({"ticker": ticker, "label": label,
                              "kept": seen[label], "dropped": accession,
                              "dropped_period_of_report": period})
            else:
                seen[label] = accession

            cached = cache_dir / f"{ticker}_{accession}.txt"
            if cached.exists():
                text = cached.read_text(encoding="utf-8")
            else:
                text = _press_release_text(filing)
                if text:
                    cached.write_text(text, encoding="utf-8")
            if not text:
                per_filing.append({"ticker": ticker, "accession": accession,
                                   "period_of_report": period, "label": label,
                                   "stated": None, "period_end": None,
                                   "offset": None, "note": "no_press_release"})
                continue

            per_filing.append({
                "ticker": ticker,
                "accession": accession,
                "period_of_report": period,
                "label": label,
                "stated": stated_fiscal_quarter(text),
                "period_end": stated_period_end(
                    text, f"{period[:4]}-{period[4:6]}-{period[6:8]}"),
            })

    # 財年結束月：先用每一份各自反推，同一家取眾數，並記錄推得一不一致
    fy_end: dict[str, int] = {}
    fy_end_conflict: dict[str, dict] = {}
    for ticker in {r["ticker"] for r in per_filing}:
        votes = Counter(
            m for m in (
                fy_end_month_from(r["period_end"], r["stated"])
                for r in per_filing if r["ticker"] == ticker
            ) if m
        )
        if votes:
            fy_end[ticker] = votes.most_common(1)[0][0]
            if len(votes) > 1:
                fy_end_conflict[ticker] = dict(votes)

    # 基準用公司自己寫的財期（`stated`），不是從期末日換算的 `derived`。
    #
    # 兩者在慣例上等價——`_col_to_quarter_label()` 對非 12 月結算的公司會把
    # 年份往後推一年，推出來就是公司自己的財年編號（NVDA 4 月底那季，兩邊
    # 都是 FY2027Q1）。但 `derived` 是看月份的，52/53 週制的公司財年結束日
    # 會在月底前後跳（COST 有幾年落在 9 月 1 日），換算就會差一季。
    # 公司自己寫的不會有這個問題，`derived` 留著當獨立交叉驗證。
    agree = disagree = 0
    for r in per_filing:
        derived = gaap_style_label(r["period_end"], fy_end.get(r["ticker"], 0))
        r["derived_label"] = derived
        if r["stated"] and derived:
            if r["stated"] == derived:
                agree += 1
            else:
                disagree += 1
        expected = r["stated"] or derived
        r["expected_label"] = expected
        a, b = _quarter_ordinal(r["label"]), _quarter_ordinal(expected or "")
        r["offset"] = (a - b) if (a is not None and b is not None) else None

    offsets = Counter(r["offset"] for r in per_filing if r["offset"] is not None)
    by_ticker: dict[str, Counter] = {}
    for r in per_filing:
        if r["offset"] is not None:
            by_ticker.setdefault(r["ticker"], Counter())[r["offset"]] += 1

    return {
        "n_filings": len(per_filing),
        "n_compared": sum(offsets.values()),
        "offset_histogram": dict(sorted(offsets.items())),
        "by_ticker": {t: dict(sorted(c.items())) for t, c in by_ticker.items()},
        "cross_check": {"stated_eq_derived": agree, "stated_ne_derived": disagree},
        "fy_end_month": fy_end,
        "fy_end_month_conflict": fy_end_conflict,
        "duplicate_labels": dupes,
        "filings": per_filing,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default="8k_period_audit.json")
    ap.add_argument("--quarters", type=int, default=8, help="每家最多查幾份 8-K")
    ap.add_argument("--cache-dir", default="8k_audit_html")
    ap.add_argument("--tickers", nargs="*", default=None)
    args = ap.parse_args()

    set_identity(config.load_config()["identity"])
    result = audit(args.tickers or TICKERS, args.quarters, Path(args.cache_dir))
    Path(args.out).write_text(json.dumps(result, ensure_ascii=False, indent=2),
                              encoding="utf-8")

    print("\n" + "=" * 72)
    print(f"抓到 {result['n_filings']} 份、成功比對 {result['n_compared']} 份")
    cc = result["cross_check"]
    print(f"交叉驗證（公司自述財期 vs 期末日換算）："
          f"一致 {cc['stated_eq_derived']} / 不一致 {cc['stated_ne_derived']}")
    print("\n偏移量分布（正數＝現行標籤比 Data_Q 的同名欄晚幾季）：")
    for off, n in result["offset_histogram"].items():
        print(f"  {off:+d} 季  {n:3d} 份")
    print("\n各家（財年結束月 / 偏移分布）：")
    for t, hist in sorted(result["by_ticker"].items()):
        print(f"  {t:6s} FY末={result['fy_end_month'].get(t, '?'):>2}月  {hist}")
    if result["fy_end_month_conflict"]:
        print("\n⚠ 財年結束月推導不一致（同一家不同季推出不同月份）：")
        for t, votes in result["fy_end_month_conflict"].items():
            print(f"  {t}: {votes}")
    print(f"\n同標籤重複（dedupe 會丟掉）：{len(result['duplicate_labels'])} 筆")
    for d in result["duplicate_labels"]:
        print(f"  {d['ticker']} {d['label']}  保留 {d['kept']} / 丟掉 {d['dropped']}")
    print(f"\n完整結果：{args.out}")


if __name__ == "__main__":
    main()
