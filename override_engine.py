"""
override_engine.py — Auto-repair for missing financial rows.

When a newly fetched ticker has key rows all-None across recent quarters,
this module diagnoses the cause and writes a permanent per-ticker override
so subsequent fetches apply the fix automatically.

Diagnosis pipeline:
    E1: rule-based fuzzy match (no API cost)
    E2: LLM call via existing ai config (only when E1 fails and api_key set)

Override storage: %APPDATA%/SEC_Financial_Tools/ticker_overrides.json
"""

from __future__ import annotations

import json
import os
import re
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd


# ── Storage path ──────────────────────────────────────────────────────────

def _default_override_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "SEC_Financial_Tools" / "ticker_overrides.json"
    return Path.home() / ".sec_financial_tools" / "ticker_overrides.json"

DEFAULT_OVERRIDE_PATH = _default_override_path()


# ── Key rows to monitor (per statement) ──────────────────────────────────

KEY_ROWS: dict[str, list[str]] = {
    "IS": ["Revenue", "Operating Income", "Net Income", "Diluted EPS"],
    "BS": ["Total Assets", "Total Liabilities", "Total Equity — Parent"],
    "CF": ["Operating Cash Flow", "Capex"],
}

# Synonyms for E1 fuzzy match: target_std_name → list of substrings to look for
# in the EDGAR DataFrame's standard_concept or label columns.
SYNONYM_MAP: dict[str, list[str]] = {
    "Revenue": [
        "Revenue", "Revenues", "SalesRevenue", "NetRevenue", "TotalRevenue",
        "RevenueFromContract", "SalesAndRevenue",
    ],
    "Operating Income": [
        "OperatingIncome", "IncomeLossFromOperations", "OperatingProfit",
        "OperatingIncomeLoss",
    ],
    "Net Income": [
        "NetIncome", "ProfitLoss", "NetEarnings", "NetIncomeLoss",
        "NetIncomeLossAvailableToCommonStockholders",
    ],
    "Diluted EPS": [
        "EarningsPerShareDiluted", "DilutedEPS", "EPSDiluted",
        "IncomeLossPerDilutedShare",
    ],
    "Total Assets": ["Assets", "TotalAssets"],
    "Total Liabilities": ["Liabilities", "TotalLiabilities", "LiabilitiesAndStockholdersEquity"],
    "Total Equity — Parent": [
        "Equity", "StockholdersEquity", "ShareholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ],
    "Operating Cash Flow": [
        "NetCashFromOperatingActivities", "NetCashProvidedByOperatingActivities",
        "OperatingCashFlow", "CashGeneratedFromOperations",
    ],
    "Capex": [
        "PaymentsToAcquirePropertyPlantAndEquipment", "Capex",
        "CapitalExpenditures", "PurchaseOfPropertyPlantAndEquipment",
    ],
}


# ── load / save ───────────────────────────────────────────────────────────

def load_overrides(ticker: str, path: Path | None = None) -> dict:
    """Return override dict for `ticker`, or {} if none recorded."""
    path = Path(path) if path else DEFAULT_OVERRIDE_PATH
    if not path.exists():
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}
    return data.get(ticker, {})


def save_overrides(ticker: str, overrides: dict, path: Path | None = None) -> None:
    """Merge `overrides` for `ticker` into the override file (other tickers untouched)."""
    path = Path(path) if path else DEFAULT_OVERRIDE_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    existing: dict = {}
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                existing = json.load(f)
        except (json.JSONDecodeError, OSError):
            existing = {}
    existing[ticker] = overrides
    with open(path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)


# ── check_key_rows ────────────────────────────────────────────────────────

def check_key_rows(
    concepts: list[str],
    values: list[list[Any]],
    statement: str,
) -> list[str]:
    """
    Return list of key row std_names that are all-None in the most recent 4 quarters.

    Only checks rows listed in KEY_ROWS[statement]; ignores other rows entirely.
    A row is flagged only when ALL values in its value list are None (or empty).
    """
    target_names = KEY_ROWS.get(statement, [])
    missing: list[str] = []
    for name in target_names:
        if name not in concepts:
            continue
        idx = concepts.index(name)
        row_vals = values[idx]
        recent = row_vals[-4:] if len(row_vals) >= 4 else row_vals
        if all(v is None for v in recent):
            missing.append(name)
    return missing


# ── E1: fuzzy match ───────────────────────────────────────────────────────

def e1_fuzzy_match(df: pd.DataFrame, target_std_name: str) -> str | None:
    """
    Search EDGAR DataFrame for a std_concept matching `target_std_name` by synonym.

    Checks both standard_concept and label columns (case-insensitive substring).
    Returns the matching standard_concept string, or None if not found.
    """
    synonyms = SYNONYM_MAP.get(target_std_name, [])
    if not synonyms:
        return None
    for _, row in df.iterrows():
        sc = str(row.get("standard_concept") or "")
        lb = str(row.get("label") or "")
        for syn in synonyms:
            s = syn.lower()
            if s in sc.lower() or s in lb.lower():
                # Return the actual standard_concept from the DataFrame row
                return sc if sc else lb
    return None


# ── E2: LLM diagnosis ─────────────────────────────────────────────────────

def _llm_call(prompt: str, ai_config: dict) -> str:
    """
    Call the configured LLM and return the raw text response.

    Uses the same provider/model/api_key as the rest of the app.
    Raises on API error (caller should catch).
    """
    provider = ai_config.get("provider", "google")
    model    = ai_config.get("model", "")
    api_key  = ai_config.get("api_key", "")

    if provider == "google":
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        m = genai.GenerativeModel(model)
        resp = m.generate_content(prompt)
        return resp.text.strip()

    if provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=64,
        )
        return resp.choices[0].message.content.strip()

    if provider == "anthropic":
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model=model,
            max_tokens=64,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text.strip()

    raise ValueError(f"Unknown AI provider: {provider}")


def _build_e2_prompt(df: pd.DataFrame, target_std_name: str, ticker: str) -> str:
    # Build concept list: "standard_concept | label" per row, skip abstract rows
    lines = []
    for _, row in df.iterrows():
        if row.get("abstract"):
            continue
        sc = str(row.get("standard_concept") or "").strip()
        lb = str(row.get("label") or "").strip()
        if sc or lb:
            lines.append(f"{sc} | {lb}")
    concept_block = "\n".join(lines[:80])  # cap at 80 rows to keep prompt short

    return (
        f"You are a financial data engineer.\n"
        f"Below is the EDGAR XBRL concept list for {ticker}:\n\n"
        f"{concept_block}\n\n"
        f"Which standard_concept best represents \"{target_std_name}\"?\n"
        f"Reply with ONLY the standard_concept string, or ABSENT if none matches.\n"
        f"One line, no explanation."
    )


def e2_llm_diagnose(
    df: pd.DataFrame,
    target_std_name: str,
    ticker: str,
    ai_config: dict,
) -> dict | None:
    """
    Call LLM to identify the correct std_concept for `target_std_name`.

    Returns:
        {"fix_type": "concept_override", "std_concept": "...", "source": "E2"}
        {"fix_type": "structural_absence", "confirmed_absent": True, "source": "E2"}
        None  — if api_key is empty or LLM call fails
    """
    if not ai_config.get("api_key", ""):
        return None

    prompt = _build_e2_prompt(df, target_std_name, ticker)
    try:
        response = _llm_call(prompt, ai_config).strip()
    except Exception:
        return None

    # ABSENT anywhere in response (case-insensitive word boundary) → structural_absence
    if re.search(r'\bABSENT\b', response, re.IGNORECASE):
        return {"fix_type": "structural_absence", "confirmed_absent": True, "source": "E2"}
    # Reject garbage: a valid std_concept is CamelCase with no spaces, ≤100 chars
    if ' ' in response or len(response) > 100:
        return None
    return {"fix_type": "concept_override", "std_concept": response, "source": "E2"}


# ── run_diagnosis ─────────────────────────────────────────────────────────

def run_diagnosis(
    ticker: str,
    statement: str,
    df: pd.DataFrame,
    missing_rows: list[str],
    ai_config: dict,
    override_path: Path | None = None,
) -> dict:
    """
    Diagnose and record overrides for `missing_rows` in `statement` for `ticker`.

    Returns dict of {std_name: override_entry} for rows successfully diagnosed.
    Saves results to override file (merging with existing overrides for ticker).
    """
    if not missing_rows:
        return {}

    today = date.today().isoformat()
    new_fixes: dict[str, dict] = {}

    for row_name in missing_rows:
        # E1: rule-based fuzzy match
        sc = e1_fuzzy_match(df, row_name)
        if sc:
            new_fixes[row_name] = {
                "fix_type": "concept_override",
                "std_concept": sc,
                "diagnosed_at": today,
                "source": "E1",
            }
            continue

        # E2: LLM (only if api_key set)
        result = e2_llm_diagnose(df, row_name, ticker, ai_config)
        if result:
            result["diagnosed_at"] = today
            new_fixes[row_name] = result

    if not new_fixes:
        return {}

    # Merge into existing overrides for this ticker and persist
    existing = load_overrides(ticker, path=override_path)
    stmt_overrides = existing.get(statement, {})
    stmt_overrides.update(new_fixes)
    existing[statement] = stmt_overrides
    save_overrides(ticker, existing, path=override_path)

    return new_fixes
