"""cli.py — 給外部 skill 用的指令列介面（TODO B1）。

    ./venv/Scripts/python.exe src/cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx
    ./venv/Scripts/python.exe src/cli.py press-release ARLO --years 2025-2026 --tables --json

**薄封裝**：這裡只做參數解析與輸出格式化，抓取邏輯一律轉呼叫既有核心函式，
GUI 與核心一行都沒動。輸出組裝走 `output_tables.append_ratio_table`，跟 GUI
同一份程式碼——CLI 產的 Excel 與 GUI 產的必須逐格相同。

**零 AI**：兩個子指令都只打 EDGAR，不呼叫任何 LLM API。`gaap` 的 E2 診斷在
`override_engine.E2_LLM_ENABLED = False` 已從源頭關掉；`press-release` 走
`press_release_tables` 的確定性解析。

`press-release` 的輸出是「已解析並篩過的表格」而不是新聞稿原文：ARLO 一季
原文 450K 字元，篩完 4.4K。skill 讀的是後者。

網路那一層集中在 `_gaap_tables` / `_earnings_filings` / `_press_release_html`
三個函式，測試把它們換掉就能完全離線跑。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import unicodedata
from pathlib import Path
from typing import Any

from config import load_config
import i18n
from errsafe import _exc_status
from excel_writer import check_output_writable, write_statements
from fiscal_input import (fiscal_quarter_of, fy_start_month,
                          quarter_label_from_announcement)
from output_tables import append_ratio_table, has_any_data
from press_release_tables import PressTable, filter_nongaap, parse_tables


class CliError(Exception):
    """使用者輸入有問題（不是程式壞掉）。main() 會印訊息並回傳非 0。"""


# ── 參數解析 ────────────────────────────────────────────────────────────────

_YEARS_RE = re.compile(r"^(\d{4})?(-)?(\d{4})?$")


def parse_years(text: str | None) -> tuple[int | None, int | None]:
    """`"2023-2026"` → (2023, 2026)、`"2024"` → (2024, 2024)、`"2020-"` → (2020, None)。"""
    if text is None:
        return (None, None)
    m = _YEARS_RE.match(text.strip())
    if m is None:
        raise CliError(f"--years 格式錯誤：{text!r}（要 2023-2026 或 2024）")
    start, dash, end = m.group(1), m.group(2), m.group(3)
    if start is None and end is None:
        raise CliError(f"--years 格式錯誤：{text!r}（要 2023-2026 或 2024）")
    if not dash:
        return (int(start), int(start))
    s = int(start) if start else None
    e = int(end) if end else None
    if s is not None and e is not None and s > e:
        raise CliError(f"--years 起始年大於結束年：{text!r}")
    return (s, e)


def resolve_identity(explicit: str | None) -> str:
    """`--identity` 沒給就用 config.json 裡進階設定填的那組。"""
    if explicit:
        return explicit
    identity = (load_config().get("identity") or "").strip()
    if not identity:
        raise CliError(
            "沒有 SEC EDGAR Identity。用 --identity \"姓名 信箱\" 指定，"
            "或先在 GUI 的「進階設定」填一次。"
        )
    return identity


# ── 網路層（測試會換掉這三個）─────────────────────────────────────────────

def _gaap_tables(**kwargs) -> list:
    from fetcher_gaap import fetch_gaap_statements
    return fetch_gaap_statements(**kwargs)


def _earnings_filings(ticker: str, identity: str, start_year: int | None,
                      end_year: int | None, max_filings: int,
                      fiscal_year_end: str | None = None) -> list[tuple[str, Any]]:
    from edgar import Company, set_identity
    from fetcher_nongaap import _list_earnings_filings

    set_identity(identity)
    return _list_earnings_filings(Company(ticker), start_year=start_year,
                                  end_year=end_year, max_filings=max_filings,
                                  fiscal_year_end=fiscal_year_end)


def _fiscal_year_end(ticker: str, identity: str) -> str | None:
    """公司財年結束日，EDGAR 的原始 MMDD 字串（`"0926"` = 9 月 26 日）。查不到回 None。

    走 EDGAR submissions 的 `fiscalYearEnd`，**一個 ticker 一次請求**，不必為了
    問財年而下載 10-K。`fetcher_gaap._detect_fy_end_month()` 是另一條路，但它要
    `filing.obj()` 最多三份 10-K，慢得多。

    **回傳完整 MMDD 不是只有月份**：列清單階段的季度標籤靠這個「日」回推名目
    季末（B5），只給月份會退化——52/53 週制的真實季末離月底最多 20 天。
    """
    from edgar import Company, set_identity

    set_identity(identity)
    raw = str(getattr(Company(ticker), "fiscal_year_end", "") or "").strip()
    return raw if len(raw) == 4 and raw.isdigit() else None


def _fy_end_month_from_mmdd(mmdd: str | None) -> int | None:
    """`"0926"` → 9。認不得回 None。

    查不到就回 None 而不是預設 12：非 12 月結算的公司會被整批標錯一到三季。
    """
    raw = str(mmdd or "").strip()
    if len(raw) != 4 or not raw.isdigit():
        return None
    month = int(raw[:2])
    return month if 1 <= month <= 12 else None


def _press_release_html(filing) -> str:
    """8-K 的新聞稿附件 HTML。取不到就回空字串（很多 8-K 沒附新聞稿）。"""
    eight_k = filing.obj()
    for pr in (getattr(eight_k, "press_releases", None) or []):
        html = pr.html()
        if html:
            return unicodedata.normalize("NFKC", html)
    return ""


# ── gaap 子指令 ─────────────────────────────────────────────────────────────

def _sheet_payload(tbl) -> dict[str, Any]:
    labels = list(getattr(tbl, "labels", []) or [])
    return {
        "sheet_name": tbl.sheet_name,
        "quarter_labels": list(tbl.quarter_labels),
        "filing_dates": list(tbl.filing_dates),
        "period_ends": list(getattr(tbl, "period_ends", []) or []),
        "rows": [
            {
                "concept": concept,
                "label": labels[i] if i < len(labels) else "",
                "values": list(values),
            }
            for i, (concept, values) in enumerate(zip(tbl.concepts, tbl.values))
        ],
    }


def cmd_gaap(args: argparse.Namespace) -> int:
    identity = resolve_identity(args.identity)
    start_year, end_year = parse_years(args.years)

    # 抓之前先確認寫得進去。失敗點本來在最後一步的 wb.save()——檔案被 Excel
    # 開著時要白等 24 秒才看到一個裸的 PermissionError。GUI 早就有這道檢查，
    # CLI 漏了（2026-08-08 實際踩到）。
    if args.xlsx:
        lock_msg = check_output_writable(args.xlsx)
        if lock_msg:
            raise CliError(lock_msg)

    from fetcher_gaap import collect_gaps, offline_report, reset_offline_tracking

    reset_offline_tracking()
    with collect_gaps() as gaps:
        tables = _gaap_tables(
            ticker=args.ticker,
            identity=identity,
            max_filings=args.max_filings,
            start_year=start_year,
            end_year=end_year,
            fetch_quarterly=not args.annual_only,
            fetch_annual=not args.quarterly_only,
        )
    # 一期都沒抓到就不寫檔——空殼 Excel 會蓋掉使用者原本好好的舊檔。
    # tables 本身不會是空 list（結構表仍在），所以要看有沒有實質資料。
    if not has_any_data(tables):
        print(f"[{args.ticker}] 沒有抓到任何資料，未寫出檔案", file=sys.stderr)
        return 1
    # 缺漏走 stderr：stdout 可能是 --json 的資料流，混進去會壞掉 pipeline。
    if gaps.has_gaps:
        print(f"[{args.ticker}] {gaps.summary()}", file=sys.stderr)
    # TODO J9：靠本地清單撐過去的一定要講出來，否則使用者會以為看到的是最新的
    _warn_if_offline(offline_report())

    append_ratio_table(tables)

    if args.xlsx:
        out = Path(args.xlsx)
        out.parent.mkdir(parents=True, exist_ok=True)
        write_statements(tables, out)
        print(f"[{args.ticker}] 寫入 {out}（{len(tables)} 張 sheet）", file=sys.stderr)

    if args.json:
        payload = {
            "ticker": args.ticker,
            "sheets": [_sheet_payload(t) for t in tables],
        }
        _emit_json(payload, args.json)
    return 0


# ── press-release 子指令 ────────────────────────────────────────────────────

# 退回舊算法（EDGAR 給不出 fiscal_year_end）的那幾季才帶這句。標籤是用 Item
# 2.02 8-K 的 period_of_report 換算的，而 EDGAR 那欄放的是**發布日**不是財期
# 結束日，實測 119 份只有 16 份對。調查報告：docs/8k-period-off-by-one.md。
_LABEL_WARNING = (
    # 全形／特殊符號要避開：Windows 主控台是 cp950，U+2212（真減號）與 ⚠ 都
    # 編不進去，`--json -` 印到 stdout 會整個 UnicodeEncodeError 掛掉。
    "label 由 period_of_report（發布日）換算，已知有系統性 off-by-one："
    "偏 -3 到 +1 季，偏多少由財年結束月決定。**改用同一季的 fiscal_label**"
    "（由 period_end + 財年結束月算出，與 Data_Q 的財季同一套慣例）；"
    "fiscal_label 是空的才退回看表格裡的期間表頭。"
)

# 走零下載規則（B5）的那幾季帶這句。label 已經與 fiscal_label 對齊——200 份
# 實測、157 份基準可信全部一致——但下載後算的 fiscal_label 仍然是最準的那個。
_LABEL_NOTE = (
    "label 由發布日 + EDGAR fiscal_year_end 回推名目季末算出（零下載規則，"
    "157 份實測與 fiscal_label 100% 一致）。fiscal_label 是下載後從新聞稿的"
    "期末日算的，仍然最準；兩者一致與否看 label_agrees_with_fiscal_label。"
)

_LABEL_SOURCE_ANNOUNCEMENT = "announcement+fiscal_year_end"
_LABEL_SOURCE_PERIOD = "period_of_report"


_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# 「June 28, 2026」「Sept. 30, 2025」「Dec 31, 2026」全收。月份只認前三個字母，
# 所以 `Sept.` 這種四字母縮寫也吃得下。ISO 日期另外一條。
_TEXT_DATE_RE = re.compile(
    r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+"
    r"(\d{1,2})\s*,?\s+(\d{4})\b",
    re.IGNORECASE,
)
_ISO_DATE_RE = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")

# 發布日不是期末日。新聞稿的日期戳與安全港聲明都寫申報當天，而財報最快也要
# 期末後兩週才發，所以申報日前 3 天內的日期一律不算期末日。
_RELEASE_DATE_SLACK_DAYS = 3


def _dates_in(text: str) -> list[str]:
    """文字裡所有日期 → ISO 字串清單（順序不保證）。認不得的日期直接跳過。"""
    from datetime import date

    found: list[str] = []
    for m in _TEXT_DATE_RE.finditer(text):
        month = _MONTHS[m.group(1).lower()[:3]]
        try:
            found.append(date(int(m.group(3)), month, int(m.group(2))).isoformat())
        except ValueError:
            continue
    for m in _ISO_DATE_RE.finditer(text):
        try:
            found.append(date(*(int(g) for g in m.groups())).isoformat())
        except ValueError:
            continue
    return found


def _cutoff_date(filing_date: str) -> str:
    """申報日 → 期末日的上限（申報日往前 3 天）。認不得日期就回空字串＝不設限。"""
    from datetime import date, timedelta

    found = _dates_in(filing_date)
    if not found:
        return ""
    y, m, d = (int(x) for x in found[0].split("-"))
    return (date(y, m, d) - timedelta(days=_RELEASE_DATE_SLACK_DAYS)).isoformat()


def _period_end_from_tables(tables: list[PressTable], not_after: str) -> str:
    """新聞稿表格裡的日期 → 本期財期結束日。抓不到回空字串。

    規則只有一條：**取不晚於「申報日前 3 天」的最新日期**。

    - 去年同期比較欄（`Three Months Ended June 29, 2025`）更早，取最新自動排除
    - 財測的未來日期（`Fiscal Quarter Ending May 3, 2026`）晚於申報日，被排除
    - 發布日（日期戳、安全港聲明的 `speak only as of August 4, 2026`）＝申報當天，
      被那 3 天的緩衝排除。財報最快也要期末後兩週才發，不會誤傷真的期末日

    **試過「優先採信 `ended` 後面那個日期」，更糟，已放棄。** 15 家 120 份實測，
    三家因此標錯：AMD 的安全港聲明是 `as of <發布日>`、INTC 的資產負債表是
    `(as of <去年年底>)`、AVGO 的註腳是 `for the fiscal quarter ended <上一季>`，
    而它們真正的期末日都只是**沒有引導詞的表頭**（colspan 展開後就剩日期本身）。
    關鍵字在新聞稿裡指向的往往不是本期。

    每一欄還要**直向串起來**再找一次：NVDA／INTC 把日期排成上下兩列
    （`April 26,` 一列、`2026` 下一列），只看單一儲存格會整家抓不到
    （實測 NVDA 三季全空）。
    """
    cutoff = _cutoff_date(not_after)
    texts: list[str] = []
    for table in tables:
        texts.extend(cell for row in table.rows for cell in row)
        texts.extend(
            " ".join(row[col] for row in table.rows if col < len(row))
            for col in range(table.n_cols)
        )

    dates = [d for text in texts for d in _dates_in(text)]
    if cutoff:
        dates = [d for d in dates if d <= cutoff]
    return max(dates) if dates else ""


def _fiscal_label(period_end: str, fy_end_month: int | None) -> str:
    """期末日 + 財年結束月 → `FY2026Q2`。任一個缺就回空字串。

    直接沿用 `fiscal_input.fiscal_quarter_of()`——那是 Excel 第 1/3 列公式的
    Python 規格，本來就把期末日往前推 15 天再取年月。**不可以用期末日的月份
    直接推財季**：COST／WDC／PANW 用 52/53 週制，期末日在月底前後浮動最多
    6 天（WDC FY2026 Q2 結束在 2026-01-02），看月份會整整差一季。
    """
    if not period_end or fy_end_month is None:
        return ""
    return fiscal_quarter_of(period_end, fy_start_month(fy_end_month))


def _quarter_payload(label: str, filing, html: str, raw: bool,
                     fy_end_month: int | None,
                     fiscal_year_end: str | None = None) -> dict[str, Any]:
    period_of_report = str(getattr(filing, "period_of_report", ""))

    # label 是列清單階段算好的（`_list_earnings_filings()`）。這裡重算一次同一
    # 條規則，只為了判斷那個 label 是新規則來的、還是逐份退回了舊算法——兩者
    # 要帶不同的警告。純日期運算，不打網路。
    from_rule = quarter_label_from_announcement(period_of_report, fiscal_year_end)
    by_rule = bool(from_rule) and from_rule == label

    entry: dict[str, Any] = {
        "label": label,
        "label_source": _LABEL_SOURCE_ANNOUNCEMENT if by_rule else _LABEL_SOURCE_PERIOD,
        "label_warning": _LABEL_NOTE if by_rule else _LABEL_WARNING,
        "period_of_report": period_of_report,
        "filing_date": str(getattr(filing, "filing_date", "")),
        "accession": str(getattr(filing, "accession_no", "")),
    }

    # 期末日與正確財季一律要算，--raw 也不例外：兩種模式吐同一組 key，
    # skill 不必分兩種情況處理。--raw 是除錯路徑，多解析一次表格划算。
    tables = parse_tables(html)
    not_after = entry["filing_date"] or entry["period_of_report"]
    entry["period_end"] = _period_end_from_tables(tables, not_after)
    entry["fiscal_label"] = _fiscal_label(entry["period_end"], fy_end_month)
    entry["fiscal_label_source"] = "period_end"

    # 兩條路算出來的財季一致嗎？`fiscal_label` 是下載後從真實期末日算的，是最準
    # 的基準。不一致代表零下載規則對這一份不成立，最可能的成因是公司改過財年
    # （EDGAR 只給「現在」的 fiscal_year_end）——那是規則 C 唯一沒有對策的結構
    # 性風險。⚠ 這個旗標只抓得到「選進來的有問題」：`--years` 篩在下載之前，
    # 被漂移害到而根本沒被選中的那幾份不會走到這裡。
    entry["label_agrees_with_fiscal_label"] = (
        entry["label"] == entry["fiscal_label"] if entry["fiscal_label"] else None
    )

    if raw:
        entry["text"] = html
        entry["chars"] = len(html)
        return entry

    kept = filter_nongaap(tables)
    entry["n_tables_total"] = len(tables)
    entry["n_tables_kept"] = len(kept)
    entry["tables"] = [t.to_dict() for t in kept]
    entry["chars"] = sum(len(t.text()) for t in kept)
    return entry


def _render_quarter(entry: dict[str, Any]) -> str:
    title = entry.get("fiscal_label") or entry["label"]
    head = (f"### {title}  (period_end={entry.get('period_end') or '?'}, "
            f"period_of_report={entry['period_of_report']}, "
            f"filed={entry['filing_date']})")
    if "text" in entry:
        return f"{head}\n{entry['text']}"
    parts = [head, f"⚠ {entry['label_warning']}"]
    for tbl in entry["tables"]:
        parts.append("")
        if tbl["caption"]:
            parts.append(f"[{tbl['caption']}]")
        parts.append("\n".join(" | ".join(r) for r in tbl["rows"]))
    return "\n".join(parts)


def cmd_press_release(args: argparse.Namespace) -> int:
    identity = resolve_identity(args.identity)
    start_year, end_year = parse_years(args.years)

    # 一個 ticker 問一次，不是每份申報問一次。**要在列清單之前查**：`--years`
    # 就是在列清單那裡篩的，而季度標籤靠這個 MMDD 才算得準（B5）。
    fiscal_year_end = _fiscal_year_end(args.ticker, identity)
    fy_end_month = _fy_end_month_from_mmdd(fiscal_year_end)
    if fy_end_month is None:
        print(f"[{args.ticker}] 查不到財年結束月，label 退回發布日換算、"
              f"fiscal_label 留空", file=sys.stderr)

    filings = _earnings_filings(ticker=args.ticker, identity=identity,
                                start_year=start_year, end_year=end_year,
                                max_filings=args.max_filings,
                                fiscal_year_end=fiscal_year_end)

    quarters: list[dict[str, Any]] = []
    skipped: list[dict[str, str]] = []
    for label, filing in filings:
        try:
            html = _press_release_html(filing)
        except Exception as exc:
            # 只印類型 + status。例外訊息挾帶完整 URL，不可 f"{exc}"。
            print(f"[{args.ticker}] {label} 下載失敗 -> "
                  f"{type(exc).__name__}{_exc_status(exc)}", file=sys.stderr)
            skipped.append({"label": label, "error": type(exc).__name__})
            continue
        if not html:
            skipped.append({"label": label, "error": "no_press_release"})
            continue
        quarters.append(_quarter_payload(label, filing, html, args.raw,
                                         fy_end_month, fiscal_year_end))

    payload = {"ticker": args.ticker, "fy_end_month": fy_end_month,
               "fiscal_year_end": fiscal_year_end,
               "quarters": quarters, "skipped": skipped}

    if args.json:
        _emit_json(payload, args.json)
    else:
        print("\n\n".join(_render_quarter(q) for q in quarters))
        if skipped:
            print(f"\n跳過：{', '.join(s['label'] for s in skipped)}", file=sys.stderr)
    return 0


# ── update-db（TODO J3）─────────────────────────────────────────────────────
#
# CLI 這條是**必要的**，不是 GUI 的附屬品：整批拓到底要好幾小時，GUI 開著過夜
# 不可靠（Windows 更新、休眠），這條才掛得上工作排程器。

def _fmt_hms(seconds: float) -> str:
    seconds = int(seconds)
    return f"{seconds // 3600}h{seconds % 3600 // 60:02d}m{seconds % 60:02d}s"


def cmd_update_db(args: argparse.Namespace) -> int:
    """把更新名單上的公司一路抓到底，只暖快取不產 Excel。

    名單維護（`--import-watchlist` / `--import-cached` / `--add` / `--remove`）
    做完就存檔並印出名單**直接結束**，不順便發動抓取——「改名單」跟「跑幾小時的
    抓取」混在同一次執行裡，手滑的代價差太多。
    """
    import local_db
    from config import CONFIG_PATH, load_config, save_config

    cfg = load_config()
    touched = False
    if args.import_watchlist:
        added = local_db.import_from_watchlist(cfg)
        print(f"從 watchlist 加入 {len(added)} 家：{', '.join(added) or '（無新增）'}")
        touched = True
    if args.import_cached:
        added = local_db.import_from_cache(cfg)
        print(f"從快取現況加入 {len(added)} 家：{', '.join(added) or '（無新增）'}")
        touched = True
    if args.add:
        added = local_db.add_tickers(cfg, args.add)
        print(f"加入 {len(added)} 家：{', '.join(added) or '（無新增）'}")
        touched = True
    if args.remove:
        removed = [t for t in args.remove if local_db.remove_ticker(cfg, t)]
        print(f"移除 {len(removed)} 家：{', '.join(removed) or '（無異動）'}")
        touched = True
    if touched:
        save_config(cfg, args.config_path or CONFIG_PATH)

    targets = local_db.normalize_tickers(args.tickers) or local_db.get_update_list(cfg)
    if args.list or touched:
        print(f"更新名單（{len(local_db.get_update_list(cfg))} 家）："
              + (", ".join(local_db.get_update_list(cfg)) or "（空）"))
        return 0
    if not targets:
        raise CliError("更新名單是空的。先用 --import-watchlist／--import-cached／"
                       "--add TICKER 建一份，或直接在指令後面列 ticker")

    identity = resolve_identity(args.identity)
    stale = local_db.stale_cache_summary()
    if stale["companies"]:
        print(f"⚠ 有 {stale['n_companies']} 家、{stale['n_filings']} 份快取是舊版 "
              f"edgartools（{'／'.join(stale['old_versions'])}）解出來的，"
              f"目前是 {stale['current']}——這些會全部重抓", file=sys.stderr)

    started = time.monotonic()

    def on_progress(event: dict) -> None:
        kind = event.get("event")
        if kind == "ticker_start":
            print(f"[{event['index'] + 1}/{event['total']}] {event['ticker']} …",
                  file=sys.stderr, flush=True)
        elif kind == "ticker_done":
            note = f"新增 {event.get('new_filings', 0)} 份"
            if event.get("gaps"):
                note += f"、缺漏 {event['gaps']} 期"
            if event.get("error"):
                note = event["error"]
            print(f"    -> {event['status']}（{note}）", file=sys.stderr, flush=True)

    report = local_db.update_local_db(targets, identity, progress=on_progress)
    elapsed = time.monotonic() - started

    payload = {
        "elapsed_seconds": round(elapsed, 1),
        "updated": report.updated,
        "skipped": report.skipped,
        "failed": report.failed,
        "stopped": report.stopped,
        # D11：連續大量抓取時 SEC 會偶發失敗、**靜默少格**。這幾家之後單獨重跑
        # 即可——第二輪會從本地快取讀已經成功的部分，只重抓失敗那幾份。
        "gap_tickers": report.gap_tickers,
        "results": [{"ticker": r.ticker, "status": r.status,
                     "new_filings": r.new_filings, "gaps": r.gaps,
                     "error": r.error, "forms": r.forms}
                    for r in report.results],
    }
    if args.json:
        _emit_json(payload, args.json)
    print(f"完成：更新 {report.updated} 家、跳過 {report.skipped} 家、"
          f"失敗 {report.failed} 家，耗時 {_fmt_hms(elapsed)}")
    if report.gap_tickers:
        print(f"⚠ 有抓取缺漏，建議單獨重跑：{', '.join(report.gap_tickers)}")
    if report.failed:
        for r in report.results:
            if r.status == "failed":
                print(f"  失敗 {r.ticker}：{r.error}", file=sys.stderr)
    return 1 if report.failed else 0


def _warn_if_offline(report: dict) -> None:
    """走過離線退路就大聲講（TODO J9）。

    ⚠ 這段**不可以省**。離線退路拿到的是「上次連得上 SEC 時的資料」，
    很可能漏掉最新一季——使用者不知道的話，會拿舊數字當最新的做判斷。
    """
    if not report:
        return
    print("\n⚠ 連不上 SEC，以下公司改用**本地既有資料**：", file=sys.stderr)
    for ticker, forms in sorted(report.items()):
        print(f"    {ticker}（{'／'.join(forms)}）", file=sys.stderr)
    print("  → 這批資料**可能漏掉最新一季**。網路恢復後重跑一次會自動補上。",
          file=sys.stderr)


# ── compare：跨公司比較（給 AI／skill 用的 JSON 出口）───────────────────────
#
# **為什麼要這條**：`comparison.py` 已經把跨公司比較最難的部分做完了——各家
# 財年不同，同一個「Q2」不是同一段時間（ARLO 的 Q2 結束在 06-28、NVDA 在
# 07-27），所以欄位是**對齊過的日曆季**。但那套邏輯原本只有 GUI 叫得到、
# 而且只輸出 Excel。AI 想比較 A/B/C/D 四家就得跑四次 `gaap` 再自己對齊，
# 等於重造一次已經做好的輪子，還容易對錯。
#
# 這條直接把 `build_comparison()` 接出來，AI 一次呼叫拿到對齊好的資料。

def _metric_choices() -> tuple[list[str], list[str]]:
    """(三張報表的科目, 比率名稱)。都是**英文機器鍵**，跟 Excel A 欄同一套。"""
    from fetcher_gaap import BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE
    from ratios import RATIO_DEFS
    statement = [row[0] for tpl in (IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE)
                 for row in tpl]
    seen, ordered = set(), []
    for name in statement:                 # Net Income 在 IS 與 CF 各有一列
        if name not in seen:
            seen.add(name)
            ordered.append(name)
    return ordered, [name for name, _, _, _ in RATIO_DEFS]


def cmd_compare(args: argparse.Namespace) -> int:
    from comparison import build_comparison
    from fetcher_gaap import offline_report, reset_offline_tracking

    statement_metrics, ratio_metrics = _metric_choices()
    if args.list_metrics:
        print(f"報表科目（{len(statement_metrics)} 個）：")
        print("  " + ", ".join(statement_metrics))
        print(f"\n財務比率（{len(ratio_metrics)} 個）：")
        print("  " + ", ".join(ratio_metrics))
        return 0

    tickers = [t.strip().upper() for t in args.tickers if t.strip()]
    if len(tickers) < 2:
        raise CliError("至少要兩家公司才比得起來")

    metrics = [m.strip() for m in (args.metrics or []) if m.strip()]
    if not metrics:
        # 預設給最常用的四個。刻意不預設「全部」——指標愈多抓得愈久，
        # 而且 AI 要的通常就是營收獲利那幾條。
        metrics = ["Revenue", "Gross Profit", "Operating Income", "Net Income"]
    known = set(statement_metrics) | set(ratio_metrics)
    unknown = [m for m in metrics if m not in known]
    if unknown:
        raise CliError(f"不認得的指標：{', '.join(unknown)}"
                       f"（用 --list-metrics 看有哪些）")

    identity = resolve_identity(args.identity)
    start_year, end_year = parse_years(args.years)
    reset_offline_tracking()

    def on_company_start(ticker: str, current: int, total: int) -> None:
        # 走 stderr：stdout 是 --json 的資料流，混進去會壞掉 pipeline
        print(f"[{current}/{total}] {ticker} …", file=sys.stderr, flush=True)

    result = build_comparison(
        tickers, identity, metrics,
        frequency="annual" if args.annual else "quarterly",
        start_year=start_year, end_year=end_year,
        on_company_start=on_company_start,
    )

    # 欄位鍵是**對齊過的日曆季**，各家在同一欄的財季不同——所以
    # `fiscal_labels` 一定要一起給，不然看數字的人無從對帳。
    #
    # ⚠ 過濾「每一家都是 None」的期間。`_aligned_labels()` 換不成日曆季的
    # 財季標籤（期末日缺漏時的退路）會原樣留著，實測 FORM 帶出
    # `FY2011Q4`／`FY2013Q1`／`FY2013Q4` 三個全空的鍵，混在 2025Q1、2025Q2
    # 中間既難看又會讓下游以為那是一個真的期間。
    all_periods = {p for per in result.metrics.values()
                   for company in per.values() for p in company}
    periods = sorted(
        p for p in all_periods
        if any(company.get(p) is not None
               for per in result.metrics.values()
               for company in per.values())
    )
    # 殘留鍵也要從各個 dict 清掉，不是只清 `periods`——契約要單純：
    # **`periods` 就是全部期間，每個 dict 的鍵剛好是這些**。留著
    # `FY2018Q4: null` 會讓讀 JSON 的人（或 AI）以為那是一個真的期間。
    keep = set(periods)

    def _trim(by_ticker: dict) -> dict:
        return {t: {p: v for p, v in d.items() if p in keep}
                for t, d in (by_ticker or {}).items()}

    payload = {
        "tickers": tickers,
        "metrics": metrics,
        "frequency": "annual" if args.annual else "quarterly",
        "periods": periods,
        # {指標: {ticker: {日曆季: 值}}}
        "data": {m: _trim(per) for m, per in result.metrics.items()},
        # {ticker: {日曆季: 該公司自己的財季標籤}}——各家財年不同時的對帳依據
        "fiscal_labels": _trim(result.fiscal_labels),
        "period_ends": _trim(result.period_ends),
        # 季報表出現的 Q4 一定是推算的（SEC 沒有 Q4 的 10-Q），標出來
        "synthetic_q4": {t: sorted(p for p in v if p in keep)
                         for t, v in result.synthetic_q4.items()},
        "failures": [{"ticker": f.ticker, "error_type": f.error_type}
                     for f in result.failures],
        # TODO J9：哪幾家是靠本地既有資料撐的（連不上 SEC）。非空代表那幾家
        # **可能漏掉最新一季**——讀這份 JSON 的人／AI 必須看得到。
        "offline": offline_report(),
    }

    if args.json:
        _emit_json(payload, args.json)
    else:
        ratio_set = set(ratio_metrics)
        print(f"比較 {', '.join(tickers)}｜{len(periods)} 個期間")
        for metric in metrics:
            # ⚠ 比率不可以除以 1e6。`Gross Margin (%)` 的值是 44.3（百分比），
            # 當成金額格式化會印出 `0.0M`——數字還在，但看起來像沒資料。
            is_ratio = metric in ratio_set
            print(f"\n{metric}")
            print("  " + _pad("期間", 12)
                  + "".join(_pad(t, 16, ">") for t in tickers))
            for period in periods:
                cells = ""
                for ticker in tickers:
                    v = (result.metrics.get(metric, {}).get(ticker, {})
                         .get(period))
                    if v is None:
                        text = "—"
                    elif is_ratio:
                        text = f"{v:,.2f}"
                    else:
                        text = f"{v / 1e6:,.1f}M"
                    cells += _pad(text, 16, ">")
                print("  " + _pad(period, 12) + cells)
    _warn_if_offline(offline_report())
    if result.failures:
        for f in result.failures:
            print(f"失敗 {f.ticker}：{f.error_type}", file=sys.stderr)
    return 1 if result.failures else 0


# ── db-status（TODO J7）─────────────────────────────────────────────────────
#
# **完全不連網**，只掃本地資料夾——這是「資料庫裡有哪些公司」的查詢入口，
# 給人看也給 AI 看。刻意跟兩個既有的東西分開：
#
#   - `update-db --list` 印的是**更新名單**（config 裡「要抓誰」），不是實際有誰
#   - `scripts/audit_local_db.py` 會對每家打兩次 SEC 列清單（「完整嗎、該補嗎」），
#     要 identity、要跑一分鐘，不適合拿來瞄一眼
#
# 這支只回答「現在本地有什麼」，244 家實測 0.5 秒。

def _pad(text: str, width: int, align: str = "<") -> str:
    """補到「終端機看起來」的寬度。中日韓字元佔兩格，`f"{s:<8}"` 是按**字數**
    補的，中文標題混英文資料一定歪。`unicodedata.east_asian_width` 的 W／F
    就是雙寬那兩類。"""
    import unicodedata
    shown = sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)
    pad = " " * max(0, width - shown)
    return pad + text if align == ">" else text + pad


def cmd_db_status(args: argparse.Namespace) -> int:
    import filing_cache
    import local_db
    from config import load_config
    from filing_cache import cache_root

    if args.rebuild:
        # 唯一會寫檔的路徑，而且只在明確要求時走。**很慢**：`rebuild_meta()`
        # 要走 `scan_filings()`，那個為了讀 4 個小欄位（form／filing_date／
        # cached_at／版本）把每一份 70KB 的 filing JSON 整個 `json.load()`
        # 進來——2026-09-18 實測 2.75 秒/家（244 家 × 75 份、冷快取），
        # 244 家約 11 分鐘。所以絕不當預設，而且要印進度。
        targets = filing_cache.list_cached_tickers()
        if args.ticker:
            # 只想看一家卻重算 244 家＝白等 11 分鐘（2026-09-18 code review）
            wanted = set(local_db.normalize_tickers(args.ticker))
            targets = [r for r in targets if r["ticker"] in wanted]
        started = time.monotonic()
        done = 0
        for i, row in enumerate(targets):
            before = local_db.read_meta(row["ticker"])
            local_db.load_meta(row["ticker"])
            if before is None or before.get("file_count") != row["count"]:
                done += 1
            print(f"\r重算 meta {i + 1}/{len(targets)}（實際重算 {done} 家，"
                  f"已花 {_fmt_hms(time.monotonic() - started)}）… ",
                  end="", file=sys.stderr, flush=True)
        print(f"\n重算完成：{done} 家，耗時 {_fmt_hms(time.monotonic() - started)}",
              file=sys.stderr)

    rows = local_db.overview_rows(load_config())
    if args.ticker:
        rows = [r for r in rows
                if r["ticker"] in local_db.normalize_tickers(args.ticker)]
    summary = local_db.overview_summary(rows)

    if args.json:
        _emit_json({"root": str(cache_root()), **summary, "companies_detail": rows},
                   args.json)
        return 0

    print(f"本地財報資料庫：{cache_root()}")
    print(f"{summary['companies']} 家 / {summary['filings']} 份 / "
          f"{summary['size_bytes'] / 1e9:.2f} GB")
    if not rows:
        print("（空的——還沒抓過任何公司，或資料庫路徑不對）")
        return 0
    # 表格顯示的是**財報期間**（period_*）。SEC 收件日（filed_*）仍在
    # --json 輸出裡，需要時查得到——兩者差一整個期間，不要混用。
    print("\n" + _pad("代號", 8) + _pad("份數", 6, ">") + "  "
          + _pad("財報期間", 20) + _pad("年數", 6, ">") + "  "
          + _pad("到底", 6) + _pad("上次查", 8, ">")
          + _pad("大小", 10, ">") + "  註記")
    for r in rows:
        span = f"{(r['period_from'] or '—')[:7]}~{(r['period_to'] or '—')[:7]}"
        years = f"{r['years']:.1f}" if r["years"] is not None else "—"
        bottom = {local_db.BOTTOM_YES: "是", local_db.BOTTOM_NO: "否"}.get(
            r["bottom"], "—") + ("?" if r["bottom_stale"] else "")
        notes = []
        if r["in_list"]:
            notes.append("★名單")
        if not r["meta_ok"]:
            notes.append("需重算")
        checked = "—" if r["stale_days"] is None else (
            "今天" if r["stale_days"] == 0 else f"{r['stale_days']}天前")
        print(_pad(r["ticker"], 8) + _pad(str(r["filings"]), 6, ">") + "  "
              + _pad(span, 20) + _pad(years, 6, ">") + "  "
              + _pad(bottom, 6) + _pad(checked, 8, ">")
              + _pad(f"{r['size_bytes'] / 1e6:.1f}MB", 10, ">")
              + "  " + " ".join(notes))
    if any(not r["meta_ok"] for r in rows):
        print("\n⚠ 有公司標示「需重算」——meta 快照跟資料夾對不上，財報期間欄"
              "顯示不出來。**資料本身沒問題**，只是快照過期。\n"
              "  修法一：等下一次「更新本地庫」，它會順便修好（不必特地跑）\n"
              "  修法二：`db-status --rebuild`，但很慢——要讀每一份 filing，"
              "實測 2.75 秒/家、244 家約 11 分鐘", file=sys.stderr)
    return 0


# ── 共用 ────────────────────────────────────────────────────────────────────

def _emit_json(payload: dict[str, Any], target: str) -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if target == "-":
        print(text)
    else:
        Path(target).parent.mkdir(parents=True, exist_ok=True)
        Path(target).write_text(text, encoding="utf-8")
        print(f"寫入 {target}", file=sys.stderr)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="src/cli.py",
        description="SEC Financial Fetcher 指令列介面（不呼叫任何 AI API）",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("ticker", help="股票代號，如 AAPL")
    common.add_argument(
        "--years",
        help="年份範圍，如 2023-2026 或 2024。"
             "注意：press-release 篩的是**發布日**換算的年份，不是財期："
             "篩選發生在下載之前，那時還讀不到期末日。非 12 月結算的公司"
             "在年份邊界可能差到 3 季（NVDA／CRM 最嚴重），要精確就把範圍"
             "放寬一年，再自己用 fiscal_label 篩",
    )
    common.add_argument("--identity", help="SEC EDGAR Identity（預設讀 config.json）")
    common.add_argument(
        "--lang", metavar="CODE",
        help="輸出 Excel 的顯示語言（B 欄譯文、Index 版面）。代號見 i18n.LANGUAGES：zh_tw／zh_cn／en／ja。預設讀 config.json，"
             "沒設定就 zh_tw。A 欄英文機器鍵與 C 欄公司原文不受影響")
    common.add_argument("--max-filings", type=int, default=80,
                        help="最多處理幾份申報（預設 80，約 20 年）")
    common.add_argument("--json", nargs="?", const="-", metavar="PATH",
                        help="輸出 JSON；不給路徑或給 - 就印到 stdout")

    g = sub.add_parser("gaap", parents=[common], help="抓 GAAP 三表 + 比率 + segment")
    g.add_argument("--xlsx", metavar="PATH", help="輸出 Excel 路徑")
    scope = g.add_mutually_exclusive_group()
    scope.add_argument("--quarterly-only", action="store_true", help="只抓 10-Q")
    scope.add_argument("--annual-only", action="store_true", help="只抓 10-K")
    g.set_defaults(func=cmd_gaap)

    p = sub.add_parser("press-release", parents=[common],
                       help="抓 Item 2.02 8-K 新聞稿的 Non-GAAP 調節表")
    # --tables 是預設行為。留著這個旗標是因為對外文件寫的就是這個介面，
    # 而且明講一次比讓人猜預設值好。
    p.add_argument("--tables", action="store_true",
                   help="輸出已解析並篩過的表格（預設行為）")
    p.add_argument("--raw", action="store_true",
                   help="改輸出新聞稿全文（除錯用，一季約 450K 字元）")
    p.set_defaults(func=cmd_press_release)

    # update-db 不吃 `common`：那組參數（ticker 單數、--years、--max-filings）
    # 對「整批拓到底」完全不適用——深度是固定的（設計書第六節：一律拓到底）。
    u = sub.add_parser("update-db",
                       help="更新本地財報資料庫（走更新名單、拓到底、只暖快取）")
    u.add_argument("tickers", nargs="*", metavar="TICKER",
                   help="只更新這幾家；不給就走 config 的更新名單")
    u.add_argument("--identity", help="SEC EDGAR Identity（預設讀 config.json）")
    u.add_argument("--json", nargs="?", const="-", metavar="PATH",
                   help="輸出這一輪的結果 JSON；不給路徑或給 - 就印到 stdout")
    u.add_argument("--list", action="store_true", help="只印出更新名單，不抓取")
    u.add_argument("--add", nargs="+", metavar="TICKER", help="加進更新名單")
    u.add_argument("--remove", nargs="+", metavar="TICKER", help="從更新名單移除")
    u.add_argument("--import-watchlist", action="store_true",
                   help="把 watchlist 全部加進更新名單")
    u.add_argument("--import-cached", action="store_true",
                   help="把快取裡已有的公司全部加進更新名單")
    u.add_argument("--config-path", help="改寫哪一份 config.json（測試用）")
    u.set_defaults(func=cmd_update_db)

    # db-status 刻意不叫 list-db：跟上面的 `update-db --list`（印更新名單）
    # 只差一個字，兩個都叫 list 保證有人搞混「名單」跟「實際有什麼」。
    d = sub.add_parser("db-status",
                       help="列出本地財報資料庫裡實際有哪些公司（不連網）")
    d.add_argument("ticker", nargs="*", metavar="TICKER", help="只看這幾家")
    d.add_argument("--json", nargs="?", const="-", metavar="PATH",
                   help="輸出 JSON；不給路徑或給 - 就印到 stdout")
    d.add_argument("--rebuild", action="store_true",
                   help="先重算對不上的 meta 快照。**很慢**：要讀每一份 filing，"
                        "實測 2.75 秒/家、244 家約 11 分鐘。多數情況不必跑——"
                        "下一次 update-db 會順便修好")
    d.set_defaults(func=cmd_db_status)

    # compare 不吃 `common`：那組的 `ticker` 是單數，這裡要多家
    c = sub.add_parser("compare",
                       help="跨公司比較（欄位是對齊過的日曆季）→ JSON 給 AI 用")
    c.add_argument("tickers", nargs="*", metavar="TICKER", help="要比的公司")
    c.add_argument("--metrics", nargs="+", metavar="NAME",
                   help="要比哪些指標（英文機器鍵）。不給就用 Revenue／"
                        "Gross Profit／Operating Income／Net Income")
    c.add_argument("--list-metrics", action="store_true",
                   help="列出所有可用的指標名稱後結束")
    c.add_argument("--annual", action="store_true", help="比年報，預設比季報")
    c.add_argument("--years", help="期間範圍，如 2023-2026。篩的是期末日")
    c.add_argument("--identity", help="SEC EDGAR Identity（預設讀 config.json）")
    c.add_argument("--json", nargs="?", const="-", metavar="PATH",
                   help="輸出 JSON；不給路徑或給 - 就印到 stdout")
    c.set_defaults(func=cmd_compare)
    return parser


def _force_utf8_io() -> None:
    """把 stdout／stderr 轉成 UTF-8。Windows 主控台預設 cp950，編不出去就整個掛。

    實測：`src/cli.py press-release ARLO`（不給 --json）在 cp950 主控台只印得出
    `失敗 -> UnicodeEncodeError`——`⚠` 編不進 cp950。新聞稿內文更危險，數字
    旁邊的 `—`、`™`、重音字母都可能出現，逐字元挑符號是治不完的。

    `errors="replace"` 是保險：真的遇到編不出去的字元印成 `?`，不要讓一個
    符號炸掉整趟輸出。測試環境的 capsys 沒有 reconfigure，getattr 擋掉。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass


def main(argv: list[str] | None = None) -> int:
    _force_utf8_io()
    parser = build_parser()
    args = parser.parse_args(argv)

    # 語言要在任何抓取／寫檔之前設好——Excel 的 B 欄是寫入當下查表的。
    # 沒給 --lang 就沿用 config.json，跟 GUI 產出的檔案一致。
    lang = getattr(args, "lang", None) or load_config().get("language")
    if getattr(args, "lang", None) and not i18n.is_supported(args.lang):
        parser.error(
            f"--lang 不認得 {args.lang!r}，可用："
            + "／".join(c for c, _ in i18n.available_languages()))
    i18n.set_lang(lang)

    if args.command == "gaap" and not args.xlsx and not args.json:
        parser.error("gaap 至少要給 --xlsx 或 --json，否則抓完沒有任何產出")

    try:
        return args.func(args)
    except CliError as exc:
        print(f"錯誤：{exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        # 例外訊息可能挾帶 URL 或金鑰，只印類型 + status code。
        print(f"失敗 -> {type(exc).__name__}{_exc_status(exc)}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
