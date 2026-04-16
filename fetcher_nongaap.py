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
