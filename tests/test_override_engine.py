"""Tests for override_engine.py — auto-repair for missing financial rows."""
import json
import pytest
import pandas as pd
from pathlib import Path
from unittest.mock import patch, MagicMock

from override_engine import (
    load_overrides,
    save_overrides,
    check_key_rows,
    e1_fuzzy_match,
    e2_llm_diagnose,
    run_diagnosis,
)


# ── helpers ───────────────────────────────────────────────────────────────

def _make_edgar_df(concepts=None, labels=None, std_concepts=None):
    """Minimal EDGAR DataFrame (like edgartools output)."""
    concepts     = concepts     or ["us-gaap_Revenues", "us-gaap_OperatingIncomeLoss", "us-gaap_NetIncomeLoss"]
    labels       = labels       or ["Total revenues", "Operating income", "Net income"]
    std_concepts = std_concepts or ["Revenues", "OperatingIncomeLoss", "NetIncome"]
    return pd.DataFrame({
        "concept":          concepts,
        "label":            labels,
        "standard_concept": std_concepts,
        "abstract":         [False] * len(concepts),
        "is_breakdown":     [False] * len(concepts),
        "2025-03-31 (Q1)":  [100.0, 20.0, 15.0],
    })


def _make_values(n_rows=9, n_quarters=4, none_rows=None):
    """Build values matrix; none_rows is list of row indices that should be all None."""
    none_rows = none_rows or []
    vals = []
    for i in range(n_rows):
        if i in none_rows:
            vals.append([None] * n_quarters)
        else:
            vals.append([float(i + 1) * 10] * n_quarters)
    return vals


# ── load_overrides / save_overrides ───────────────────────────────────────

def test_load_overrides_returns_empty_when_no_file(tmp_path):
    result = load_overrides("AAPL", path=tmp_path / "overrides.json")
    assert result == {}


def test_load_overrides_returns_ticker_data(tmp_path):
    data = {"AAPL": {"IS": {"Revenue": {"fix_type": "concept_override", "std_concept": "Revenues"}}}}
    f = tmp_path / "overrides.json"
    f.write_text(json.dumps(data), encoding="utf-8")
    result = load_overrides("AAPL", path=f)
    assert result == data["AAPL"]


def test_load_overrides_returns_empty_for_unknown_ticker(tmp_path):
    data = {"AAPL": {"IS": {}}}
    f = tmp_path / "overrides.json"
    f.write_text(json.dumps(data), encoding="utf-8")
    result = load_overrides("MSFT", path=f)
    assert result == {}


def test_load_overrides_handles_malformed_json(tmp_path):
    f = tmp_path / "overrides.json"
    f.write_text("{not valid json", encoding="utf-8")
    result = load_overrides("AAPL", path=f)
    assert result == {}


def test_save_overrides_creates_file(tmp_path):
    p = tmp_path / "overrides.json"
    save_overrides("AAPL", {"IS": {"Revenue": {"fix_type": "concept_override"}}}, path=p)
    assert p.exists()


def test_save_overrides_merges_with_existing(tmp_path):
    p = tmp_path / "overrides.json"
    p.write_text(json.dumps({"TSLA": {"IS": {}}}), encoding="utf-8")
    save_overrides("AAPL", {"IS": {"Revenue": {"fix_type": "concept_override"}}}, path=p)
    data = json.loads(p.read_text(encoding="utf-8"))
    assert "TSLA" in data
    assert "AAPL" in data


def test_save_overrides_overwrites_same_ticker(tmp_path):
    p = tmp_path / "overrides.json"
    p.write_text(json.dumps({"AAPL": {"IS": {"old": {}}}}), encoding="utf-8")
    save_overrides("AAPL", {"IS": {"Revenue": {"fix_type": "derived"}}}, path=p)
    data = json.loads(p.read_text(encoding="utf-8"))
    assert "old" not in data["AAPL"]["IS"]
    assert "Revenue" in data["AAPL"]["IS"]


# ── check_key_rows ────────────────────────────────────────────────────────

IS_CONCEPTS = [
    "Revenue", "COGS", "Gross Profit", "R&D", "SG&A", "Other Op Expense",
    "Operating Income", "Interest Expense", "Other Non-op", "Pre-tax Income",
    "Tax", "Net Income", "EPS Basic", "Diluted EPS", "Shares Basic",
    "Shares Diluted", "D&A", "Stock Comp", "EBITDA", "Op Margin",
    "Net Margin", "ROE",
]

def test_check_key_rows_returns_empty_when_all_present():
    # All key rows have values
    concepts = IS_CONCEPTS
    vals = _make_values(n_rows=len(concepts), n_quarters=4)
    result = check_key_rows(concepts, vals, "IS")
    assert result == []


def test_check_key_rows_detects_revenue_none():
    concepts = IS_CONCEPTS
    revenue_idx = concepts.index("Revenue")
    vals = _make_values(n_rows=len(concepts), n_quarters=4, none_rows=[revenue_idx])
    result = check_key_rows(concepts, vals, "IS")
    assert "Revenue" in result


def test_check_key_rows_detects_operating_income_none():
    concepts = IS_CONCEPTS
    oi_idx = concepts.index("Operating Income")
    vals = _make_values(n_rows=len(concepts), n_quarters=4, none_rows=[oi_idx])
    result = check_key_rows(concepts, vals, "IS")
    assert "Operating Income" in result


def test_check_key_rows_ignores_non_key_rows():
    concepts = IS_CONCEPTS
    # Make a non-key row None (e.g., "Op Margin" at index 19)
    vals = _make_values(n_rows=len(concepts), n_quarters=4, none_rows=[19])
    result = check_key_rows(concepts, vals, "IS")
    assert result == []


def test_check_key_rows_requires_all_recent_quarters_none():
    """A row with at least one non-None value in last 4 quarters is not flagged."""
    concepts = IS_CONCEPTS
    revenue_idx = concepts.index("Revenue")
    vals = _make_values(n_rows=len(concepts), n_quarters=4)
    # Only the last quarter is None — should NOT be flagged
    vals[revenue_idx] = [100.0, 200.0, 150.0, None]
    result = check_key_rows(concepts, vals, "IS")
    assert "Revenue" not in result


def test_check_key_rows_cf_operating_cash_flow():
    cf_concepts = [
        "Net Income", "D&A", "Stock Comp", "Change in Receivables",
        "Change in Payables", "Other Operating CF", "Operating Cash Flow",
        "Capex", "Acquisitions", "Investment Proceeds", "Other Investing CF",
        "Investing Cash Flow", "Debt Issuance", "Debt Repayment",
        "Dividends Paid", "Share Buybacks", "Other Financing CF",
        "Financing Cash Flow", "FX Effect", "Net Change in Cash",
        "Beginning Cash", "Ending Cash", "Capital Expenditures",
        "Free Cash Flow", "Other CF",
    ]
    ocf_idx = cf_concepts.index("Operating Cash Flow")
    vals = _make_values(n_rows=len(cf_concepts), n_quarters=4, none_rows=[ocf_idx])
    result = check_key_rows(cf_concepts, vals, "CF")
    assert "Operating Cash Flow" in result


# ── e1_fuzzy_match ────────────────────────────────────────────────────────

def test_e1_fuzzy_match_finds_by_std_concept():
    df = _make_edgar_df(
        std_concepts=["Revenues", "OperatingIncomeLoss", "NetIncome"]
    )
    result = e1_fuzzy_match(df, "Revenue")
    assert result == "Revenues"


def test_e1_fuzzy_match_finds_by_label():
    df = _make_edgar_df(
        std_concepts=["CustomConcept", "OperatingIncomeLoss", "NetIncome"],
        labels=["Total revenues and other income", "Operating income", "Net income"],
    )
    result = e1_fuzzy_match(df, "Revenue")
    assert result == "CustomConcept"


def test_e1_fuzzy_match_returns_none_when_no_match():
    df = _make_edgar_df(
        std_concepts=["Revenues", "OperatingIncomeLoss", "NetIncome"],
        labels=["Revenue", "Operating income", "Net income"],
    )
    result = e1_fuzzy_match(df, "Diluted EPS")
    assert result is None


def test_e1_fuzzy_match_net_income_profit_loss():
    """ProfitLoss is a valid synonym for Net Income (TSLA/BA pattern)."""
    df = _make_edgar_df(
        std_concepts=["Revenues", "OperatingIncomeLoss", "ProfitLoss"],
        labels=["Revenue", "Operating income", "Net income attributable"],
    )
    result = e1_fuzzy_match(df, "Net Income")
    assert result == "ProfitLoss"


# ── _llm_call google-genai 遷移（2026-08-21，TODO D7）────────────────────

def test_llm_call_google_uses_genai_client():
    """provider=google 要用新版 google-genai SDK（Client().models.generate_content），
    不是舊版已終止支援的 google.generativeai（GenerativeModel）。"""
    import override_engine as oe

    fake_response = MagicMock()
    fake_response.text = "  Revenues  \n"
    fake_client = MagicMock()
    fake_client.models.generate_content.return_value = fake_response

    with patch("google.genai.Client", return_value=fake_client) as mock_client_cls:
        result = oe._llm_call("prompt text", {
            "provider": "google", "model": "gemini-flash-latest", "api_key": "fake-key",
        })

    mock_client_cls.assert_called_once_with(api_key="fake-key")
    fake_client.models.generate_content.assert_called_once_with(
        model="gemini-flash-latest", contents="prompt text",
    )
    assert result == "Revenues"


# ── e2_llm_diagnose ───────────────────────────────────────────────────────

def test_e2_llm_diagnose_returns_none_when_no_api_key():
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": ""}
    result = e2_llm_diagnose(df, "Revenue", "AAPL", ai_cfg)
    assert result is None


def test_e2_llm_diagnose_returns_concept_override_on_match(monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call", lambda prompt, cfg: "Revenues")
    result = e2_llm_diagnose(df, "Revenue", "AAPL", ai_cfg)
    assert result == {"fix_type": "concept_override", "std_concept": "Revenues", "source": "E2"}


def test_e2_llm_diagnose_returns_structural_absence_on_absent(monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call", lambda prompt, cfg: "ABSENT")
    result = e2_llm_diagnose(df, "Operating Income", "XOM", ai_cfg)
    assert result == {"fix_type": "structural_absence", "confirmed_absent": True, "source": "E2"}


def test_e2_llm_diagnose_handles_llm_whitespace(monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call", lambda prompt, cfg: "  OperatingIncomeLoss  \n")
    result = e2_llm_diagnose(df, "Operating Income", "COHR", ai_cfg)
    assert result["std_concept"] == "OperatingIncomeLoss"


def test_e2_llm_diagnose_absent_in_sentence_returns_structural_absence(monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    """LLM returns 'ABSENT' embedded in a sentence — should still be structural_absence."""
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call",
                        lambda prompt, cfg: "I cannot find a match, ABSENT")
    result = e2_llm_diagnose(df, "Operating Income", "XOM", ai_cfg)
    assert result is not None
    assert result["fix_type"] == "structural_absence"


def test_e2_llm_diagnose_garbage_response_returns_none(monkeypatch):
    """LLM returns a full sentence (not a concept name) — should return None, not store garbage."""
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call",
                        lambda prompt, cfg: "The best match is OperatingIncomeLoss based on the data.")
    result = e2_llm_diagnose(df, "Operating Income", "COHR", ai_cfg)
    assert result is None


def test_e2_llm_diagnose_absent_case_insensitive(monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    """ABSENT check must be case-insensitive."""
    df = _make_edgar_df()
    ai_cfg = {"provider": "google", "model": "gemini", "api_key": "test-key"}
    monkeypatch.setattr("override_engine._llm_call", lambda prompt, cfg: "absent")
    result = e2_llm_diagnose(df, "Operating Income", "XOM", ai_cfg)
    assert result["fix_type"] == "structural_absence"


# ── run_diagnosis ─────────────────────────────────────────────────────────

def test_run_diagnosis_returns_empty_when_no_missing(tmp_path):
    df = _make_edgar_df()
    result = run_diagnosis(
        ticker="AAPL", statement="IS", df=df,
        missing_rows=[], ai_config={"api_key": ""},
        override_path=tmp_path / "overrides.json",
    )
    assert result == {}


def test_run_diagnosis_e1_path_writes_override(tmp_path):
    df = _make_edgar_df(std_concepts=["Revenues", "OperatingIncomeLoss", "NetIncome"])
    p = tmp_path / "overrides.json"
    result = run_diagnosis(
        ticker="NEWCO", statement="IS", df=df,
        missing_rows=["Revenue"], ai_config={"api_key": ""},
        override_path=p,
    )
    assert "Revenue" in result
    assert result["Revenue"]["fix_type"] == "concept_override"
    # Override saved to file
    saved = json.loads(p.read_text(encoding="utf-8"))
    assert "NEWCO" in saved
    assert "Revenue" in saved["NEWCO"]["IS"]


def test_run_diagnosis_e2_path_when_e1_fails(tmp_path, monkeypatch):
    # E2 預設關閉（見 override_engine.E2_LLM_ENABLED），測 E2 行為要先打開
    monkeypatch.setattr("override_engine.E2_LLM_ENABLED", True)
    df = _make_edgar_df(
        std_concepts=["SomeOddConcept", "AnotherConcept", "ThirdConcept"],
        labels=["Something", "Another thing", "Third thing"],
    )
    monkeypatch.setattr("override_engine._llm_call", lambda prompt, cfg: "SomeOddConcept")
    p = tmp_path / "overrides.json"
    result = run_diagnosis(
        ticker="NEWCO", statement="IS", df=df,
        missing_rows=["Revenue"],
        ai_config={"api_key": "test-key", "provider": "google", "model": "gemini"},
        override_path=p,
    )
    assert result["Revenue"]["source"] == "E2"


def test_run_diagnosis_skips_e2_when_no_api_key(tmp_path):
    df = _make_edgar_df(
        std_concepts=["SomeOddConcept", "AnotherConcept", "ThirdConcept"],
        labels=["Something", "Another thing", "Third thing"],
    )
    p = tmp_path / "overrides.json"
    result = run_diagnosis(
        ticker="NEWCO", statement="IS", df=df,
        missing_rows=["Revenue"],
        ai_config={"api_key": ""},
        override_path=p,
    )
    # E1 failed, E2 skipped → no override for Revenue
    assert "Revenue" not in result


# ═════════════════════════════════════════════════════════════════════════════
# E2 LLM 診斷預設關閉（2026-08-03）
#
# 專案定位確立為「只抓 SEC 原始資料」，GAAP 這條路徑不該碰 AI。
# E1 模糊比對是純程式的，保留；E2 會呼叫 LLM，預設關掉。
# 關掉的位置在 override_engine 而不是呼叫端——即使有人不小心把 ai_config
# 傳進來，也不會真的打 API。
# ═════════════════════════════════════════════════════════════════════════════

def test_e2_disabled_by_default():
    import override_engine
    assert override_engine.E2_LLM_ENABLED is False


def test_e2_returns_none_when_disabled(monkeypatch):
    """關閉時直接回 None，且**完全不呼叫** _llm_call。"""
    import override_engine, pandas as pd
    called = []
    monkeypatch.setattr(override_engine, "_llm_call",
                        lambda *a, **k: called.append(1) or "{}")
    monkeypatch.setattr(override_engine, "E2_LLM_ENABLED", False)

    out = override_engine.e2_llm_diagnose(
        pd.DataFrame({"concept": ["us-gaap_Revenues"], "label": ["Revenue"]}),
        "Revenue", "TEST", {"api_key": "k", "provider": "google"},
    )
    assert out is None
    assert called == []


def test_e2_still_works_when_explicitly_enabled(monkeypatch):
    """留一條開關給日後需要時用，不是把功能刪掉。"""
    import override_engine, pandas as pd
    monkeypatch.setattr(override_engine, "E2_LLM_ENABLED", True)
    monkeypatch.setattr(override_engine, "_llm_call", lambda *a, **k: "Revenues")
    out = override_engine.e2_llm_diagnose(
        pd.DataFrame({"concept": ["us-gaap_Revenues"], "label": ["Revenue"]}),
        "Revenue", "TEST", {"api_key": "k", "provider": "google"},
    )
    assert out is not None
