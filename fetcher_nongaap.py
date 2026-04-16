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
