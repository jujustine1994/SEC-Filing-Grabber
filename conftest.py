import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).parent
sys.path.insert(0, str(_ROOT / "src"))


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "slow: live integration tests hitting real EDGAR API (excluded from default CI)",
    )
    config.addinivalue_line(
        "markers",
        "b1: B1 overflow-row tests (subset of slow, run with: pytest -m 'slow and b1')",
    )
    config.addinivalue_line(
        "markers",
        "cf_overflow: CF YTD overflow correctness tests (subset of slow, run with: pytest -m 'slow and cf_overflow')",
    )
