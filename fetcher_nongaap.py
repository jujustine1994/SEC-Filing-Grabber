"""
fetcher_nongaap.py — Non-GAAP data extraction from 8-K press releases via AI.

Flow:
  8-K (Item 2.02) → edgartools eps_reconciliation + AI on EX-99.1 press release
  → nongaap_cache.json (per-ticker, incremental) → StatementTable list
"""

import json
import sys
import unicodedata
from pathlib import Path
from typing import Any

from edgar import Company, set_identity

from fetcher_gaap import StatementTable

CACHE_FILENAME = "nongaap_cache.json"


# ── Cache I/O ───────────────────────────────────────────────────────────────

def _load_cache(cache_path: Path) -> dict:
    """Load nongaap_cache.json. Returns {} if missing or malformed."""
    if not cache_path.exists():
        return {}
    try:
        with open(cache_path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_cache(cache_path: Path, data: dict) -> None:
    """Save cache dict to JSON. Creates parent dirs if needed."""
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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
        return {k: float(v) for k, v in parsed.items() if v is not None}

    except Exception as exc:
        print(f"[fetcher_nongaap] AI call failed: {exc!r}", file=sys.stderr)
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
        print(f"[fetcher_nongaap] metrics extraction failed: {exc!r}", file=sys.stderr)
        return {}


# ── 8-K discovery ────────────────────────────────────────────────────────────

def _get_earnings_filings(company) -> list[tuple[str, Any]]:
    """Return list of (quarter_label, filing) for 8-K filings with Item 2.02.

    Sorted oldest → newest. Deduplicated by quarter_label.
    """
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
            results.append((label, filing))
        except Exception as exc:
            print(f"[fetcher_nongaap] 8-K scan warning: {exc!r}", file=sys.stderr)
            continue

    # Sort oldest first, deduplicate by quarter_label (keep first = oldest filing for that period)
    seen: set[str] = set()
    deduped = []
    for label, filing in reversed(results):
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing))
    return list(reversed(deduped))


# ── Public API ───────────────────────────────────────────────────────────────

def fetch_nongaap_statements(
    ticker: str,
    identity: str,
    ai_config: dict,
    output_dir: Path,
    progress_cb=None,
) -> list[StatementTable]:
    """Fetch Non-GAAP statements from 8-K filings for a ticker.

    Args:
        ticker:      Stock ticker, e.g. "AAPL"
        identity:    SEC EDGAR identity string
        ai_config:   {"provider": ..., "model": ..., "api_key": ...}
        output_dir:  Directory where nongaap_cache.json will be stored
        progress_cb: Optional callable(current, total, label) for progress updates

    Returns:
        List of StatementTable: [Data_EPS_Recon, Data_NonGAAP] (omits None tables)
    """
    set_identity(identity)
    company = Company(ticker)
    cache_path = Path(output_dir) / CACHE_FILENAME

    cache = _load_cache(cache_path)
    filings = _get_earnings_filings(company)

    new_filings = [(lbl, f) for lbl, f in filings if lbl not in cache]
    total = len(new_filings)

    for i, (quarter_label, filing) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eight_k = filing.obj()
            eps_recon = _extract_eps_recon(eight_k)
            metrics   = _extract_nongaap_metrics(eight_k, ai_config)
            cache[quarter_label] = {
                "filing_date": str(filing.filing_date),
                "eps_recon":   eps_recon,
                "metrics":     metrics,
            }
            # Save after each quarter (crash-safe incremental)
            _save_cache(cache_path, cache)
        except Exception as exc:
            print(f"[fetcher_nongaap] {quarter_label} failed: {exc!r}", file=sys.stderr)

    tables: list[StatementTable] = []
    eps_tbl = _build_eps_recon_table(ticker, cache)
    ng_tbl  = _build_nongaap_table(ticker, cache)
    if eps_tbl:
        tables.append(eps_tbl)
    if ng_tbl:
        tables.append(ng_tbl)
    return tables
