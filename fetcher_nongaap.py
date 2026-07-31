"""
fetcher_nongaap.py — Non-GAAP data extraction from 8-K press releases via AI.

Flow:
  8-K (Item 2.02) → edgartools eps_reconciliation + AI on EX-99.1 press release
  → nongaap_cache.json (per-ticker, incremental) → StatementTable list
"""

import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any

from edgar import Company, set_identity

from errsafe import _exc_status
from fetcher_gaap import StatementTable

CACHE_FILENAME = "nongaap_cache.json"


# ── Cache I/O ───────────────────────────────────────────────────────────────
#
# Cache format (new):  {ticker: {quarter_label: {filing_date, eps_recon, metrics}}}
# Cache format (old):  {quarter_label: {...}}  — single-ticker, no isolation
#
# Old-format detection: top-level key starts with "FY" (quarter labels look like FY2024Q1).
# On first load of an old file, the data is treated as belonging to the current ticker
# and will be written in new format on the next save.

def _load_cache(cache_path: Path, ticker: str) -> dict:
    """Return ticker's cached quarters from nongaap_cache.json.

    Returns {} if file is missing, malformed, or ticker has no cached data.
    Transparently handles old single-ticker format (no ticker key).
    """
    if not cache_path.exists():
        return {}
    try:
        with open(cache_path, encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}

    if not isinstance(data, dict) or not data:
        return {}

    # Old format: keys are quarter labels like "FY2024Q1"
    if next(iter(data)).startswith("FY"):
        return data  # treat as this ticker's data; migrated to new format on next save

    return data.get(ticker, {})


def _save_cache(cache_path: Path, ticker: str, ticker_data: dict) -> None:
    """Write ticker's data into nongaap_cache.json (multi-ticker format).

    Reads the existing file, updates this ticker's slice, and writes back.
    Old-format files are silently replaced with new format on first write
    (ticker_data already contains all previously loaded quarters).
    """
    if cache_path.exists():
        try:
            with open(cache_path, encoding="utf-8") as f:
                all_data = json.load(f)
            if not isinstance(all_data, dict):
                all_data = {}
            # Old-format file: discard raw dict — ticker_data already has the migrated content
            if all_data and next(iter(all_data)).startswith("FY"):
                all_data = {}
        except (json.JSONDecodeError, OSError):
            all_data = {}
    else:
        all_data = {}

    all_data[ticker] = ticker_data
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)


# ── Period helpers ───────────────────────────────────────────────────────────

def _period_to_quarter_label(period_of_report: str) -> str:
    """Convert '20240331' or '2024-03-31' to 'FY2024Q1'."""
    period = period_of_report.replace("-", "")
    year = period[:4]
    month = int(period[4:6])
    if month <= 3:
        suffix = "Q1"
    elif month <= 6:
        suffix = "Q2"
    elif month <= 9:
        suffix = "Q3"
    else:
        suffix = "Q4"
    return f"FY{year}{suffix}"


# ── StatementTable builders ──────────────────────────────────────────────────

def _build_eps_recon_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_EPS_Recon StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # Collect all EPS recon keys (union across quarters)
    all_keys: list[str] = []
    seen: set[str] = set()
    for q in sorted_qs:
        for key in cache[q].get("eps_recon", {}):
            if key not in seen:
                all_keys.append(key)
                seen.add(key)

    if not all_keys:
        return None

    values: list[list[Any]] = []
    for key in all_keys:
        values.append([cache[q].get("eps_recon", {}).get(key) for q in sorted_qs])

    return StatementTable(
        sheet_name="Data_EPS_Recon",
        quarter_labels=sorted_qs,
        filing_dates=filing_dates,
        concepts=all_keys,
        values=values,
        ticker=ticker,
        labels=[""] * len(all_keys),
    )


def _build_nongaap_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_NonGAAP StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # Union of all metric names
    all_metrics: list[str] = []
    seen: set[str] = set()
    for q in sorted_qs:
        for key in cache[q].get("metrics", {}):
            if key not in seen:
                all_metrics.append(key)
                seen.add(key)

    if not all_metrics:
        return None

    values: list[list[Any]] = []
    for metric in all_metrics:
        values.append([cache[q].get("metrics", {}).get(metric) for q in sorted_qs])

    return StatementTable(
        sheet_name="Data_NonGAAP",
        quarter_labels=sorted_qs,
        filing_dates=filing_dates,
        concepts=all_metrics,
        values=values,
        ticker=ticker,
        labels=[""] * len(all_metrics),
    )


# ── Extraction functions ─────────────────────────────────────────────────────

def _extract_eps_recon(eight_k) -> dict[str, float]:
    """Extract EPS reconciliation using edgartools native support.

    Returns dict like {"GAAP EPS": 0.53, "SBC": -0.12, "Non-GAAP EPS": 0.65}.
    Returns {} if not available.
    """
    try:
        earnings = getattr(eight_k, "earnings", None)
        if earnings is None:
            return {}
        recon = getattr(earnings, "eps_reconciliation", None)
        if recon is None:
            return {}
        df = recon.dataframe
        if df is None or df.empty:
            return {}

        result: dict[str, float] = {}
        value_cols = [c for c in df.columns if c not in {"label", "concept", "description"}]
        if not value_cols:
            return {}
        val_col = value_cols[0]

        label_col = "label" if "label" in df.columns else df.columns[0]
        for _, row in df.iterrows():
            label = str(row.get(label_col, "") or "").strip()
            val = row.get(val_col)
            if label and val is not None:
                try:
                    result[label] = float(val)
                except (ValueError, TypeError):
                    pass
        return result
    except Exception as exc:
        print(f"[fetcher_nongaap] eps_recon warning: {exc!r}", file=sys.stderr)
        return {}


# ── Metric name normalisation ────────────────────────────────────────────────
#
# Press releases (especially NVDA) embed period tokens and comparison-period
# rows in their Non-GAAP tables, causing the AI to return names like:
#
#   "Non-GAAP Gross margin (Q4 FY26)"   ← trailing suffix — strip
#   "Non-GAAP Q4 FY26 Gross margin"     ← period in middle — strip
#   "Q2 FY26 Non-GAAP Revenue"          ← period as prefix — strip
#   "Non-GAAP Q3 FY26 Gross margin"     ← comparison quarter — dedup, discard
#   "Non-GAAP FY2026 Gross margin"      ← annual duplicate — discard if Q exists
#   "Expected Non-GAAP Gross margin (Q1 FY27)"  ← guidance — always discard
#
# Strategy:
#   1. Drop guidance / outlook entries.
#   2. Strip ALL period tokens (Q\d FY\d+ or FY\d+) and trailing noise labels.
#   3. Two-pass dedup: quarterly-tagged entries win over FY-only entries;
#      first occurrence wins within each bucket (press release shows current
#      quarter first, so comparison-period rows are skipped automatically).

_PERIOD_TOKEN_RE = re.compile(r"\b(?:Q[1-4]\s+)?FY\d{2,4}\b", re.IGNORECASE)
_QTR_TOKEN_RE    = re.compile(r"\bQ[1-4]\s+FY\d{2,4}\b", re.IGNORECASE)
_TRAILING_LABEL_RE = re.compile(
    r"\s*\(\s*(?:Table|Summary|Chart|Note|Detail|Details)\s*\)\s*$", re.IGNORECASE
)
_GUIDANCE_PREFIXES = ("expected", "outlook", "guidance", "anticipated", "projected")


def _clean_metric_name(name: str) -> str:
    """Strip period tokens and trailing noise labels, normalise whitespace."""
    clean = _PERIOD_TOKEN_RE.sub("", name)
    clean = _TRAILING_LABEL_RE.sub("", clean)
    clean = re.sub(r"\(\s*\)", "", clean)   # remove empty parens left by period removal
    return re.sub(r"\s+", " ", clean).strip()


def _normalize_nongaap_metrics(raw: dict[str, Any]) -> dict[str, Any]:
    """Strip period tokens, drop guidance rows, deduplicate (quarterly > FY).

    Two-pass approach:
      Pass 1 — entries that had a Q# token (quarterly) or no period token at all.
      Pass 2 — entries that had only an FY token (annual totals).
    First occurrence of each clean name wins within each pass.
    """
    quarterly: dict[str, Any] = {}
    fy_only: dict[str, Any] = {}

    for name, val in raw.items():
        name_lower = name.lower().strip()
        # Drop guidance / outlook / forward-looking entries
        if any(name_lower.startswith(p) for p in _GUIDANCE_PREFIXES):
            continue
        if "outlook" in name_lower:
            continue

        has_qtr  = bool(_QTR_TOKEN_RE.search(name))
        has_period = bool(_PERIOD_TOKEN_RE.search(name))
        is_fy_only = has_period and not has_qtr  # FY token present but no Q# token

        clean = _clean_metric_name(name)
        if not clean:
            continue

        if is_fy_only:
            if clean not in fy_only:
                fy_only[clean] = val
        else:
            if clean not in quarterly:
                quarterly[clean] = val

    # Quarterly (incl. no-token) wins; FY fills gaps only
    result = dict(quarterly)
    for clean, val in fy_only.items():
        if clean not in result:
            result[clean] = val
    return result


# ── AI extraction ────────────────────────────────────────────────────────────

_NONGAAP_PROMPT = """\
你是財務分析師。以下是一份公司季度財報新聞稿（Markdown 格式）。
請提取所有 Non-GAAP 財務指標，回傳 JSON 格式：

{{"指標名稱": 數值（純數字，不含貨幣符號或逗號）}}

規則：
- 只取 Non-GAAP / Adjusted / Excluding 相關指標
- 金額若以百萬為單位則乘以 1000000，以十億為單位則乘以 1000000000
- 百分比直接回傳數字（如 17.6%→17.6）
- 若找不到任何 Non-GAAP 指標，回傳空 JSON {{}}
- 只回傳 JSON，不要說明文字

新聞稿內容：
{press_release_text}
"""


def _call_ai(text: str, ai_config: dict) -> dict[str, Any]:
    """Call configured AI provider with press release text. Returns parsed JSON dict."""
    provider = ai_config.get("provider", "google")
    model    = ai_config.get("model", "")
    api_key  = ai_config.get("api_key", "")
    prompt   = _NONGAAP_PROMPT.format(press_release_text=text[:12000])  # token guard

    try:
        if provider == "google":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            response = genai.GenerativeModel(model).generate_content(prompt)
            raw = response.text
        elif provider == "openai":
            from openai import OpenAI
            response = OpenAI(api_key=api_key).chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=1024,
            )
            raw = response.choices[0].message.content
        elif provider == "anthropic":
            import anthropic
            response = anthropic.Anthropic(api_key=api_key).messages.create(
                model=model, max_tokens=1024,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text
        else:
            return {}

        # Strip markdown code fences if present
        raw = raw.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:])
        if raw.endswith("```"):
            raw = raw.rsplit("```", 1)[0]

        parsed = json.loads(raw.strip())
        result = {}
        for k, v in parsed.items():
            if v is None:
                continue
            try:
                result[k] = float(v)
            except (ValueError, TypeError):
                pass
        return _normalize_nongaap_metrics(result)

    except Exception as exc:
        # 只印類型 + status code。不可用 {exc!r} / {exc}：三家 LLM SDK 的例外訊息
        # 天生挾帶 URL（google-generativeai 走 REST 時帶 ?key=），而 launcher.ps1
        # 刻意留著主控台視窗，等於把金鑰印在畫面上。
        print(
            f"[fetcher_nongaap] AI call failed: {type(exc).__name__}{_exc_status(exc)}",
            file=sys.stderr,
        )
        return {}


def _extract_nongaap_metrics(eight_k, ai_config: dict) -> dict[str, Any]:
    """Get press release text and call AI to extract Non-GAAP metrics.

    Returns dict of {metric_name: value}. Returns {} on any failure.
    """
    try:
        press_releases = getattr(eight_k, "press_releases", None)
        text = None

        if press_releases:
            for pr in press_releases:
                try:
                    text = pr.markdown() if hasattr(pr, "markdown") else pr.text()
                    if text:
                        break
                except Exception:
                    continue

        # Fallback: search attachments for EX-99
        if not text:
            try:
                attachments = getattr(eight_k, "_filing", None)
                if attachments:
                    attachments = attachments.attachments
                    for att in attachments:
                        doc_type = str(getattr(att, "document_type", "") or "")
                        if "EX-99" in doc_type.upper():
                            text = att.markdown() if hasattr(att, "markdown") else att.text()
                            if text:
                                break
            except Exception:
                pass

        if not text:
            return {}

        text = unicodedata.normalize("NFKC", text)
        return _call_ai(text, ai_config)

    except Exception as exc:
        # 在 AI 呼叫鏈上（下方 _call_ai）。目前 _call_ai 自己吞例外不會冒上來，
        # 但仍只印類型 + status：哪天內層改成 re-raise，這裡就會變成金鑰洩漏點。
        print(
            f"[fetcher_nongaap] metrics extraction failed: "
            f"{type(exc).__name__}{_exc_status(exc)}",
            file=sys.stderr,
        )
        return {}


# ── 8-K discovery ────────────────────────────────────────────────────────────

def _filter_nongaap_by_year(
    filings: list[tuple],
    start_year: int | None,
    end_year: int | None,
) -> list[tuple]:
    """Filter (label, filing, eight_k) tuples by year extracted from label (e.g. 'FY2021Q2' → 2021)."""
    if start_year is None and end_year is None:
        return filings
    result = []
    for item in filings:
        label = item[0]
        m = re.search(r'(\d{4})', label)
        if m is None:
            result.append(item)
            continue
        year = int(m.group(1))
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(item)
    return result


def _list_earnings_filings(
    company,
    start_year: int | None = None,
    end_year: int | None = None,
    max_filings: int = 80,
) -> list[tuple[str, Any]]:
    """Return [(quarter_label, filing)] for earnings 8-Ks, newest first.

    Filters entirely on listing metadata (``items`` and ``period_of_report``),
    which EDGAR supplies with the filing index — no document is downloaded here.
    Callers download only the filings they actually need.

    Item 2.02 is "Results of Operations and Financial Condition", i.e. the
    earnings release. SEC adopted that numbering on 2004-08-23; earlier filings
    used Item 12 or Item 5 and are not matched. See the design doc for why that
    is acceptable (max_filings defaults to 80 quarters ≈ 20 years).
    """
    candidates: list[tuple[str, Any]] = []
    for filing in company.get_filings(form="8-K", amendments=False):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" not in items:
            continue
        period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
        if len(period) < 8:
            continue
        candidates.append((_period_to_quarter_label(period), filing))

    # Dedupe by quarter, keeping the oldest filing for each — matches prior behaviour
    # where a corrected re-filing does not displace the original release.
    seen: set[str] = set()
    deduped: list[tuple[str, Any]] = []
    for label, filing in reversed(candidates):      # oldest → newest
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing))
    deduped.reverse()                               # back to newest → oldest

    deduped = _filter_nongaap_by_year(deduped, start_year, end_year)
    return deduped[:max_filings]


def _get_earnings_filings(company) -> list[tuple[str, Any, Any]]:
    """Return list of (quarter_label, filing, eight_k) for 8-K filings with Item 2.02.

    Sorted oldest → newest. Deduplicated by quarter_label (keeps oldest filing per quarter).
    eight_k is the already-parsed filing object — callers should use it directly to avoid
    a redundant filing.obj() call.
    """
    # edgartools returns newest-first; we reverse to get oldest-first for deduplication
    results = []
    for filing in company.get_filings(form="8-K", amendments=False):
        try:
            eight_k = filing.obj()
            items = getattr(eight_k, "items", []) or []
            has_202 = any("2.02" in str(item) for item in items)
            if not has_202:
                if not getattr(eight_k, "has_earnings", False):
                    continue
            period = str(filing.period_of_report or "").replace("-", "")
            if len(period) < 8:
                continue
            label = _period_to_quarter_label(period)
            results.append((label, filing, eight_k))
        except Exception as exc:
            print(f"[fetcher_nongaap] 8-K scan warning: {exc!r}", file=sys.stderr)
            continue

    # Sort oldest first, deduplicate by quarter_label (keep first = oldest filing for that period)
    seen: set[str] = set()
    deduped = []
    for label, filing, eight_k in reversed(results):
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing, eight_k))
    return list(reversed(deduped))


# ── Public API ───────────────────────────────────────────────────────────────

def fetch_nongaap_statements(
    ticker: str,
    identity: str,
    ai_config: dict,
    output_dir: Path,
    progress_cb=None,
    max_filings: int = 80,
    start_year: int | None = None,
    end_year: int | None = None,
) -> list[StatementTable]:
    """Fetch Non-GAAP statements from 8-K filings for a ticker.

    Args:
        ticker:      Stock ticker, e.g. "AAPL"
        identity:    SEC EDGAR identity string
        ai_config:   {"provider": ..., "model": ..., "api_key": ...}
        output_dir:  Directory where nongaap_cache.json will be stored
        progress_cb: Optional callable(current, total, label) for progress updates
        max_filings: Max number of earnings quarters to process (newest first, default 80)

    Returns:
        List of StatementTable: [Data_EPS_Recon, Data_NonGAAP] (omits None tables)
    """
    set_identity(identity)
    company = Company(ticker)
    cache_path = Path(output_dir) / CACHE_FILENAME

    cache = _load_cache(cache_path, ticker)
    filings = _get_earnings_filings(company)[:max_filings]  # newest max_filings quarters only
    filings = _filter_nongaap_by_year(filings, start_year, end_year)

    new_filings = [(lbl, f, ek) for lbl, f, ek in filings if lbl not in cache]
    total = len(new_filings)

    for i, (quarter_label, filing, eight_k) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eps_recon = _extract_eps_recon(eight_k)
            metrics   = _extract_nongaap_metrics(eight_k, ai_config)
            cache[quarter_label] = {
                "filing_date": str(filing.filing_date),
                "eps_recon":   eps_recon,
                "metrics":     metrics,
            }
            # Save after each quarter (crash-safe incremental)
            _save_cache(cache_path, ticker, cache)
        except Exception as exc:
            # 同上：這層包住 _extract_nongaap_metrics -> _call_ai，屬 AI 呼叫鏈。
            print(
                f"[fetcher_nongaap] {quarter_label} failed: "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )

    tables: list[StatementTable] = []
    eps_tbl = _build_eps_recon_table(ticker, cache)
    ng_tbl  = _build_nongaap_table(ticker, cache)
    if eps_tbl:
        tables.append(eps_tbl)
    if ng_tbl:
        tables.append(ng_tbl)
    return tables
