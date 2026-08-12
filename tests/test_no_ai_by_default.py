"""釘住「這個專案預設不打 AI API」。

CTH 的 Gemini 額度按**次數**計且已見底，所以 GAAP 路徑刻意做到零 API 依賴。
這件事目前只靠兩個模組層級的旗標維持，沒有測試釘著——有人為了除錯把它改回
True 忘了改回來，下次抓一批公司就會靜默燒掉整天的額度，而且看不出來。

失敗時**不要直接把旗標改回去**：先確認為什麼有人打開它。
"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def test_e2_llm_diagnosis_is_off():
    import override_engine
    assert override_engine.E2_LLM_ENABLED is False


def test_nongaap_fetching_is_off():
    """停用要停在**源頭**。只在輸出端過濾掉 `Data_NonGAAP` 的話，AI 會照常
    抓完 6 季才被丟掉，等於白燒額度（2026-08-03 差點犯的錯）。"""
    import main
    assert main.NONGAAP_ENABLED is False


def test_e2_diagnosis_does_not_reach_the_llm(monkeypatch):
    """把 `_llm_call` 換成會爆的東西，跑 E2 診斷，不該爆。

    只檢查旗標是 False 太弱——旗標對了但有人在別處繞過去也一樣會燒額度。
    這個測試釘的是**行為**：診斷路徑走完，`_llm_call` 一次都沒被碰到。
    刻意帶著有效的 api_key，證明擋下來的是旗標不是「沒有金鑰」。
    """
    import override_engine

    calls: list = []

    def _record(*args, **kwargs):
        # ⚠ 這裡不可以 raise。`e2_llm_diagnose` 把 `_llm_call` 包在
        # `except Exception: return None` 裡，拋什麼都會被吞掉，測試就變成
        # 「旗標開著也會過」的空測試（第一版就是這樣寫的，正控制才抓到）。
        calls.append(args)
        return "ABSENT"

    monkeypatch.setattr(override_engine, "_llm_call", _record)
    result = override_engine.e2_llm_diagnose(
        pd.DataFrame({"concept": ["us-gaap:Revenues"], "label": ["Net sales"]}),
        target_std_name="Revenue",
        ticker="TEST",
        ai_config={"provider": "google", "model": "x", "api_key": "looks-real"},
    )
    assert calls == [], "E2 打了 AI API——額度按次數計，這條路徑必須是零呼叫"
    assert result is None


def test_nongaap_guard_sits_before_the_fetch_call():
    """守衛要在**呼叫 fetch 之前**，不是在輸出端過濾。

    差別很重要：擋在輸出端的話，AI 會照常抓完 6 季**才被丟掉**，等於白燒額度
    （2026-08-03 差點犯的錯）。
    """
    source = (ROOT / "src" / "main.py").read_text(encoding="utf-8")
    guard = source.index("fetch_nongaap and NONGAAP_ENABLED")
    call = source.index("ng_tables = fetch_nongaap_statements(")
    assert guard < call
