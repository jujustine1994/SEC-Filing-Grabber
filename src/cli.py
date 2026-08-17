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
import unicodedata
from pathlib import Path
from typing import Any

from config import load_config
import i18n
from errsafe import _exc_status
from excel_writer import check_output_writable, write_statements
from fiscal_input import fiscal_quarter_of, fy_start_month
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
                      end_year: int | None, max_filings: int) -> list[tuple[str, Any]]:
    from edgar import Company, set_identity
    from fetcher_nongaap import _list_earnings_filings

    set_identity(identity)
    return _list_earnings_filings(Company(ticker), start_year=start_year,
                                  end_year=end_year, max_filings=max_filings)


def _fy_end_month(ticker: str, identity: str) -> int | None:
    """公司財年結束月（1-12）。查不到回 None。

    走 EDGAR submissions 的 `fiscalYearEnd`（`"0926"` = 9 月 26 日），
    **一個 ticker 一次請求**，不必為了問財年而下載 10-K。
    `fetcher_gaap._detect_fy_end_month()` 是另一條路，但它要 `filing.obj()`
    最多三份 10-K，慢得多。

    查不到就回 None 而不是預設 12：非 12 月結算的公司會被整批標錯一到三季，
    而那正是這次要修的錯誤。
    """
    from edgar import Company, set_identity

    set_identity(identity)
    raw = str(getattr(Company(ticker), "fiscal_year_end", "") or "").strip()
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

    from fetcher_gaap import collect_gaps

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

# 每一季都帶著這句。季度標籤是用 Item 2.02 8-K 的 period_of_report 換算的，
# 而 EDGAR 那欄放的是**發布日**不是財期結束日，實測 12 家有 11 家晚一季。
# 調查報告：docs/8k-period-off-by-one.md。修法會動到快取 key，尚未修。
_LABEL_WARNING = (
    # 全形／特殊符號要避開：Windows 主控台是 cp950，U+2212（真減號）與 ⚠ 都
    # 編不進去，`--json -` 印到 stdout 會整個 UnicodeEncodeError 掛掉。
    "label 由 period_of_report（發布日）換算，已知有系統性 off-by-one："
    "偏 -3 到 +1 季，偏多少由財年結束月決定。**改用同一季的 fiscal_label**"
    "（由 period_end + 財年結束月算出，與 Data_Q 的財季同一套慣例）；"
    "fiscal_label 是空的才退回看表格裡的期間表頭。"
)


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
                     fy_end_month: int | None) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "label": label,
        "label_source": "period_of_report",
        "label_warning": _LABEL_WARNING,
        "period_of_report": str(getattr(filing, "period_of_report", "")),
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

    filings = _earnings_filings(ticker=args.ticker, identity=identity,
                                start_year=start_year, end_year=end_year,
                                max_filings=args.max_filings)

    # 一個 ticker 問一次，不是每份申報問一次。查不到就是 None，`fiscal_label`
    # 留空——寧可沒有標籤也不要一個錯的。
    fy_end_month = _fy_end_month(args.ticker, identity)
    if fy_end_month is None:
        print(f"[{args.ticker}] 查不到財年結束月，fiscal_label 留空",
              file=sys.stderr)

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
        quarters.append(_quarter_payload(label, filing, html, args.raw, fy_end_month))

    payload = {"ticker": args.ticker, "fy_end_month": fy_end_month,
               "quarters": quarters, "skipped": skipped}

    if args.json:
        _emit_json(payload, args.json)
    else:
        print("\n\n".join(_render_quarter(q) for q in quarters))
        if skipped:
            print(f"\n跳過：{', '.join(s['label'] for s in skipped)}", file=sys.stderr)
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
