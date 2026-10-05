import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).parent
sys.path.insert(0, str(_ROOT / "src"))


@pytest.fixture(autouse=True)
def isolated_database(tmp_path, monkeypatch):
    """No offline test may fall through to the user's permanent database."""
    from database import create_database
    root = tmp_path.parent / (tmp_path.name + "-database")
    create_database(root, test_mode=True)
    monkeypatch.setenv("SEC_LOCAL_DB_ROOT", str(root))
    monkeypatch.setenv("SEC_CONFIG_PATH", str(tmp_path.parent / (tmp_path.name + "-config.json")))
    if "main" in sys.modules:
        monkeypatch.setattr(sys.modules["main"], "CONFIG_PATH",
                            tmp_path.parent / (tmp_path.name + "-config.json"))
    yield root


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
