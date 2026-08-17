"""按鈕標示不得互相混淆（CTH 2026-08-17 回報）。

症狀：Ticker 旁的「快速掃描 ▶」與底下的「▶ 執行」都掛著同一個 ▶，
兩顆看起來都像「開始跑」，使用者分不出哪顆是查期間、哪顆是真的抓資料，
於是同一顆按了兩次還以為要按兩次才會動。

這裡釘的是**視覺上分得開**這個性質，不是釘死某個字。譯文可以改，
但四種語言都不准讓這兩顆共用同一個動作符號。
"""

import re

import pytest

import i18n

LANGS = [code for code, _, _ in i18n.LANGUAGES]

# 會被使用者讀成「按下去會開始做事」的符號。新增按鈕若要用新符號，
# 一併加進來，否則這條測試會漏掉它。
ACTION_GLYPHS = "▶⬇🔍⏵►"


def _strings(lang: str) -> dict[str, str]:
    return i18n._strings(lang)


def _glyphs(text: str) -> set[str]:
    return {ch for ch in text if ch in ACTION_GLYPHS}


# ── 掃描鍵 vs 執行鍵 ──────────────────────────────────────────────────────

@pytest.mark.parametrize("lang", LANGS)
def test_scan_and_run_buttons_do_not_share_an_action_glyph(lang):
    """『查可用期間』與『開始抓取』是完全不同的兩件事，符號不可以一樣。"""
    s = _strings(lang)
    scan, run = s["gui.btn.scan"], s["gui.btn.run"]
    shared = _glyphs(scan) & _glyphs(run)
    assert not shared, (
        f"{lang}: 掃描鍵 {scan!r} 與執行鍵 {run!r} 共用符號 {sorted(shared)}——"
        "使用者會分不出哪顆才是真的開始抓資料"
    )


@pytest.mark.parametrize("lang", LANGS)
def test_scan_and_run_buttons_are_not_the_same_text(lang):
    s = _strings(lang)
    assert s["gui.btn.scan"] != s["gui.btn.run"]


@pytest.mark.parametrize("lang", LANGS)
def test_both_action_buttons_carry_an_action_glyph(lang):
    """光是文字不同還不夠——沒有符號的按鈕在一排 ttk 按鈕裡不顯眼。
    兩顆都要有符號，只是不能是同一個。"""
    s = _strings(lang)
    for key in ("gui.btn.scan", "gui.btn.run"):
        assert _glyphs(s[key]), f"{lang}: {key} = {s[key]!r} 沒有任何動作符號"


# ── 掃描進行中的分段提示 ──────────────────────────────────────────────────

@pytest.mark.parametrize("lang", LANGS)
def test_scan_hint_mentions_how_long_it_takes(lang):
    """使用者按了掃描以為沒反應而重按，是因為看不到「還在跑、要等多久」。
    這條提示必須帶上秒數，光寫『掃描中』治不好。"""
    s = _strings(lang)
    hint = s["gui.status.scan_hint"]
    assert re.search(r"\d", hint), f"{lang}: scan_hint = {hint!r} 沒有給出預估時間"
