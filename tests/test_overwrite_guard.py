"""覆蓋既有輸出檔前先提示（CTH 2026-08-17 要求）。

現況是不問就覆蓋。使用者事前完全不知道自己要蓋掉誰——檔名是程式組的
（`AAPL Apple Inc data.xlsx`），未必記得上次存到哪。

做法刻意保持簡單（CTH 指定）：跳個提醒，按確定就覆蓋掉。沒有「不再提醒」
旗標，沒有另存新檔，備份維持原本的單一滾動 `.bak.xlsx`。

跟 test_first_run_language.py 一樣不開視窗：真正會出錯的是「什麼時候該問」
這個判斷，不是對話框版面。
"""

import pytest

import i18n
import main

LANGS = [code for code, _, _ in i18n.LANGUAGES]


@pytest.fixture
def existing(tmp_path):
    p = tmp_path / "AAPL Apple Inc data.xlsx"
    p.write_bytes(b"not really xlsx")
    return p


@pytest.fixture
def absent(tmp_path):
    return tmp_path / "NVDA.xlsx"


# ── 批量更新：一次問完，不逐檔問 ──────────────────────────────────────────

def test_lists_only_the_files_that_exist(tmp_path, existing, absent):
    other = tmp_path / "MSFT.xlsx"
    other.write_bytes(b"x")
    assert main.existing_outputs([existing, absent, other]) == [existing, other]


def test_returns_empty_when_nothing_would_be_overwritten(absent):
    """全部都是新檔就不該跳視窗，不多一次點擊。"""
    assert main.existing_outputs([absent]) == []


def test_preserves_input_order(tmp_path):
    """訊息裡要列出檔名，順序跟使用者的 watchlist 一致才好對。"""
    paths = []
    for name in ("C.xlsx", "A.xlsx", "B.xlsx"):
        p = tmp_path / name
        p.write_bytes(b"x")
        paths.append(p)
    assert main.existing_outputs(paths) == paths


# ── 訊息文字 ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("lang", LANGS)
def test_overwrite_messages_exist_in_every_language(lang):
    s = i18n._strings(lang)
    for key in ("gui.dlg.overwrite_title", "gui.msg.overwrite_single",
                "gui.msg.overwrite_batch"):
        assert key in s, f"{lang} 缺 {key}"


@pytest.mark.parametrize("lang", LANGS)
def test_single_message_names_the_file(lang):
    """使用者要能從訊息本身看出蓋掉哪個檔，不是只說「檔案已存在」。"""
    assert "{name}" in i18n._strings(lang)["gui.msg.overwrite_single"]


@pytest.mark.parametrize("lang", LANGS)
def test_batch_message_reports_both_counts(lang):
    """「12 家裡有 8 家會被蓋」比「有些檔案會被蓋」有用得多。"""
    msg = i18n._strings(lang)["gui.msg.overwrite_batch"]
    assert "{total}" in msg and "{n}" in msg and "{names}" in msg
