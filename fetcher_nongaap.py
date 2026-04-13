"""
fetcher_nongaap.py — Non-GAAP data extraction from 8-K press releases via AI.

Phase 2 — Not yet implemented.
"""

from fetcher_gaap import StatementTable


def fetch_nongaap_statements(
    ticker: str, identity: str, ai_config: dict
) -> list[StatementTable]:
    """Fetch Non-GAAP metrics from 8-K Exhibit 99.1 for all available quarters.

    Phase 2 placeholder — raises NotImplementedError.

    Args:
        ticker:     Stock ticker, e.g. "AAPL"
        identity:   SEC EDGAR identity string
        ai_config:  {"provider": ..., "model": ..., "api_key": ...}
    """
    raise NotImplementedError(
        "Non-GAAP extraction is not yet implemented (Phase 2). "
        "Please uncheck Non-GAAP in the settings."
    )
