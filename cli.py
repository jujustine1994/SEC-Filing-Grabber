"""cli.py — 給外部 skill 用的指令列介面（TODO B1）。

    ./venv/Scripts/python.exe cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx
    ./venv/Scripts/python.exe cli.py press-release ARLO --years 2025-2026 --tables --json

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
from errsafe import _exc_status
from excel_writer import check_output_writable, write_statements
from output_tables import append_ratio_table
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

    tables = _gaap_tables(
        ticker=args.ticker,
        identity=identity,
        max_filings=args.max_filings,
        start_year=start_year,
        end_year=end_year,
        fetch_quarterly=not args.annual_only,
        fetch_annual=not args.quarterly_only,
    )
    if not tables:
        print(f"[{args.ticker}] 沒有抓到任何資料", file=sys.stderr)
        return 1

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
    "label 由 period_of_report（發布日）換算，已知有系統性 off-by-one："
    "多數情況比實際財期晚一季。實際財期請看表格裡的期間表頭。"
)


def _quarter_payload(label: str, filing, html: str, raw: bool) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "label": label,
        "label_source": "period_of_report",
        "label_warning": _LABEL_WARNING,
        "period_of_report": str(getattr(filing, "period_of_report", "")),
        "filing_date": str(getattr(filing, "filing_date", "")),
        "accession": str(getattr(filing, "accession_no", "")),
    }
    if raw:
        entry["text"] = html
        entry["chars"] = len(html)
        return entry

    tables = parse_tables(html)
    kept = filter_nongaap(tables)
    entry["n_tables_total"] = len(tables)
    entry["n_tables_kept"] = len(kept)
    entry["tables"] = [t.to_dict() for t in kept]
    entry["chars"] = sum(len(t.text()) for t in kept)
    return entry


def _render_quarter(entry: dict[str, Any]) -> str:
    head = (f"### {entry['label']}  (period_of_report={entry['period_of_report']}, "
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
        quarters.append(_quarter_payload(label, filing, html, args.raw))

    payload = {"ticker": args.ticker, "quarters": quarters, "skipped": skipped}

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
        prog="cli.py",
        description="SEC Financial Fetcher 指令列介面（不呼叫任何 AI API）",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("ticker", help="股票代號，如 AAPL")
    common.add_argument("--years", help="年份範圍，如 2023-2026 或 2024")
    common.add_argument("--identity", help="SEC EDGAR Identity（預設讀 config.json）")
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


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

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
