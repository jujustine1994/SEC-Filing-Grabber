"""下載中途斷網的處理（CTH 2026-08-17）。

修之前：`fetcher_gaap.py` 的 `filing.obj()` 迴圈全是 `except Exception:
continue`，網路在第 5 份 filing 掛掉，那一季就默默跳過，程式照樣顯示
「完成」，Excel 少一季使用者不會知道。錯誤只 print 到 stderr，GUI 看不到。

定案的規則（CTH 2026-08-17 拍板）：

  網路類例外  → 退避重試，救得回閃斷；救不回就**中止整趟、不寫檔**
  解析類例外  → 維持現況跳過該期，不中止（現在用起來沒問題，不要動）

兩者一定要分得開，這個模組就是那條界線。
"""

import socket

import pytest

from net_retry import NetworkDownError, is_network_error, with_retry


# 冒充 httpx 的例外——靠類別名稱辨識，不 import httpx（它是 edgartools
# 的相依，不是我們直接宣告的）
class ConnectTimeout(Exception):
    pass


class ReadTimeout(Exception):
    pass


class RemoteProtocolError(Exception):
    pass


# ── 什麼算網路問題 ────────────────────────────────────────────────────────

@pytest.mark.parametrize("exc", [
    ConnectTimeout("timed out"),
    ReadTimeout("timed out"),
    RemoteProtocolError("server disconnected"),
    ConnectionError("connection reset"),          # builtin
    ConnectionResetError("reset by peer"),
    ConnectionAbortedError("aborted"),
    TimeoutError("timed out"),
    socket.gaierror("name resolution failed"),
])
def test_network_failures_are_recognised(exc):
    assert is_network_error(exc)


@pytest.mark.parametrize("exc", [
    ValueError("bad XBRL"),
    KeyError("Revenue"),
    TypeError("NoneType is not subscriptable"),
    AttributeError("'NoneType' object has no attribute 'income_statement'"),
    IndexError("list index out of range"),
])
def test_parse_failures_are_not_network_failures(exc):
    """這幾種正是舊申報沒有 XBRL、報表結構對不上會拋的東西。
    它們必須留在「跳過該期」那條路，不可以升級成中止整趟。"""
    assert not is_network_error(exc)


def test_disk_errors_are_not_treated_as_network():
    """OSError 太廣——磁碟滿、權限不足都是 OSError，重試沒有意義。
    socket.gaierror 雖然也是 OSError 子類，但它有自己的名字所以認得出來。"""
    assert not is_network_error(PermissionError("denied"))
    assert not is_network_error(FileNotFoundError("missing"))
    assert not is_network_error(IsADirectoryError("nope"))


def test_wrapped_network_error_is_seen_through_the_cause_chain():
    """edgartools 會把底層例外包起來再拋，只看最外層會漏判成解析錯誤，
    於是斷網被當成「這一季解不開」默默跳過——正是要修的那個 bug。"""
    outer = RuntimeError("failed to load filing")
    outer.__cause__ = ConnectTimeout("timed out")
    assert is_network_error(outer)


def test_cause_chain_walk_survives_a_cycle():
    """__cause__ 兜成環時不可以無限迴圈——那才是真的「跳不出來」。"""
    a = RuntimeError("a")
    b = RuntimeError("b")
    a.__cause__ = b
    b.__cause__ = a
    assert is_network_error(a) is False


@pytest.mark.parametrize("status,expected", [
    (503, True), (502, True), (504, True), (429, True), (500, True),
    (404, False), (403, False), (400, False),
])
def test_http_status_decides_for_status_errors(status, expected):
    """SEC 偶爾回 503 或 429（打太快）。那是暫時的、值得重試；
    404 重試一百次還是 404。"""
    class Resp:
        status_code = status

    class HTTPStatusError(Exception):
        response = Resp()

    assert is_network_error(HTTPStatusError("boom")) is expected


# ── 重試 ──────────────────────────────────────────────────────────────────

def test_succeeds_without_retrying_when_nothing_is_wrong():
    slept = []
    calls = []
    assert with_retry(lambda: (calls.append(1), "ok")[1], sleep=slept.append) == "ok"
    assert len(calls) == 1
    assert slept == []


def test_retries_a_network_error_then_succeeds():
    """閃斷就是要救回來，不該讓使用者少一季。"""
    slept, calls = [], []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise ConnectTimeout("timed out")
        return "ok"

    assert with_retry(flaky, attempts=3, sleep=slept.append) == "ok"
    assert len(calls) == 3


def test_gives_up_after_the_attempt_limit():
    slept, calls = [], []

    def always_down():
        calls.append(1)
        raise ConnectTimeout("timed out")

    with pytest.raises(NetworkDownError):
        with_retry(always_down, attempts=3, sleep=slept.append)
    assert len(calls) == 3, "attempts=3 是總共試 3 次，不是 3 次之外再加一次"


def test_exhausted_retries_raise_network_down_not_the_raw_error():
    """呼叫端要能用一個型別分辨「網路斷了該中止」與「這期解不開該跳過」。
    原始例外掛在 __cause__ 上，除錯時查得到。"""
    def always_down():
        raise ConnectTimeout("timed out")

    with pytest.raises(NetworkDownError) as caught:
        with_retry(always_down, attempts=2, sleep=lambda s: None)
    assert isinstance(caught.value.__cause__, ConnectTimeout)


def test_backoff_grows_between_attempts():
    """固定間隔在對方限流時只會一直撞牆。要退避。"""
    slept = []

    def always_down():
        raise ConnectTimeout("timed out")

    with pytest.raises(NetworkDownError):
        with_retry(always_down, attempts=4, base_delay=2.0, sleep=slept.append)
    assert slept == [2.0, 4.0, 8.0], "最後一次失敗後不該再等"


def test_parse_errors_are_not_retried_and_pass_straight_through():
    """重試解析錯誤只是讓使用者多等 14 秒看到同一個錯。
    而且必須原樣拋出、不可包成 NetworkDownError，否則呼叫端會誤以為斷網。"""
    slept, calls = [], []

    def bad_data():
        calls.append(1)
        raise ValueError("bad XBRL")

    with pytest.raises(ValueError):
        with_retry(bad_data, attempts=3, sleep=slept.append)
    assert len(calls) == 1
    assert slept == []


def test_network_down_error_is_not_mistaken_for_a_network_error():
    """NetworkDownError 本身是我們的訊號，不是要再重試一輪的原因。"""
    assert not is_network_error(NetworkDownError("已重試三次"))
