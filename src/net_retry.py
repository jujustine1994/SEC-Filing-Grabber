"""net_retry.py — 分辨「網路斷了」與「這一期解不開」，並對前者退避重試。

CTH 2026-08-17 拍板的規則：

    網路類例外  → 退避重試；救不回就中止整趟、不寫檔
    解析類例外  → 維持現況跳過該期，不中止

為什麼一定要分開：`fetcher_gaap.py` 的四個 `filing.obj()` 迴圈原本都是
`except Exception: continue`。網路在第 5 份 filing 掛掉時，那一季被當成
「這期沒資料」默默跳過，程式照樣顯示「完成」，使用者拿到一份少一季的
Excel 卻不知道。比直接失敗更糟。

本模組刻意不 import httpx。httpx 是 edgartools 的相依不是我們宣告的，
版本一換例外類別的位置就可能變；改用類別名稱比對，多一層保險。
"""

from __future__ import annotations

import time
from typing import Callable, TypeVar

T = TypeVar("T")


class NetworkDownError(Exception):
    """重試用盡仍連不上。呼叫端看到這個就中止整趟，不要寫出殘缺的檔案。

    原始例外掛在 `__cause__` 上，除錯時查得到。
    """


# 以類別名稱辨識網路問題。涵蓋 httpx（edgartools 用的）、requests、
# urllib、socket 四套的命名。
#
# 只列名稱不列型別，是因為同名的類別在不同套件裡都代表同一件事
# （requests.ConnectionError 與內建 ConnectionError 都是連不上）。
_NETWORK_EXC_NAMES = frozenset({
    # httpx
    "ConnectError", "ConnectTimeout", "ReadTimeout", "WriteTimeout",
    "PoolTimeout", "ReadError", "WriteError", "CloseError",
    "RemoteProtocolError", "LocalProtocolError", "ProxyError",
    "NetworkError", "TimeoutException", "TransportError",
    # requests
    "ConnectionError", "Timeout", "ChunkedEncodingError",
    "ContentDecodingError", "SSLError", "ProxyError",
    # urllib / http.client
    "URLError", "IncompleteRead", "RemoteDisconnected",
    # 內建與 socket
    "ConnectionResetError", "ConnectionAbortedError", "ConnectionRefusedError",
    "BrokenPipeError", "TimeoutError", "gaierror", "herror",
})

# 值得重試的 HTTP 狀態。SEC 會在打太快時回 429，維護時回 5xx。
# 404 / 403 重試一百次還是一樣，不列入。
_RETRYABLE_STATUS = frozenset({429, 500, 502, 503, 504})


def _status_verdict(exc: BaseException) -> bool | None:
    """例外帶 HTTP 狀態碼時由狀態碼決定。沒有就回 None 交給名稱比對。"""
    response = getattr(exc, "response", None)
    status = getattr(response, "status_code", None)
    if isinstance(status, int):
        return status in _RETRYABLE_STATUS
    return None


def is_network_error(exc: BaseException) -> bool:
    """這個例外是「連不上」還是「資料解不開」？

    會沿著 `__cause__` / `__context__` 往下找——edgartools 會把底層例外
    包成自己的型別再拋，只看最外層會把斷網誤判成解析錯誤，於是被默默
    跳過，正是要修的那個 bug。

    走訪時記已看過的 id，`__cause__` 兜成環時不會無限迴圈。
    """
    if isinstance(exc, NetworkDownError):
        # 我們自己的訊號，不是「再重試一輪」的理由
        return False

    seen: set[int] = set()
    current: BaseException | None = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))

        verdict = _status_verdict(current)
        if verdict is not None:
            return verdict
        if type(current).__name__ in _NETWORK_EXC_NAMES:
            return True

        current = current.__cause__ or current.__context__
    return False


def with_retry(fn: Callable[[], T], attempts: int = 3,
               base_delay: float = 2.0,
               sleep: Callable[[float], None] = time.sleep) -> T:
    """跑 `fn()`，只在網路問題時退避重試。

    Args:
        attempts:   總共試幾次（不是「重試幾次」）。3 = 原本一次 + 兩次重試。
        base_delay: 第一次退避的秒數，之後每次加倍（2 → 4 → 8）。
        sleep:      注入點，測試用；正式跑就是 time.sleep。

    Raises:
        NetworkDownError: 網路問題重試用盡。原始例外在 `__cause__`。
        其他例外: 解析錯誤等原樣往外拋，一次都不重試——重試只會讓使用者
                  多等十幾秒看到同一個錯。
    """
    last: BaseException | None = None
    for attempt in range(1, attempts + 1):
        try:
            return fn()
        except Exception as exc:
            if not is_network_error(exc):
                raise
            last = exc
            if attempt < attempts:
                sleep(base_delay * (2 ** (attempt - 1)))

    raise NetworkDownError(
        f"{type(last).__name__} after {attempts} attempts"
    ) from last


def sec_reachable(timeout: float = 5.0) -> bool:
    """現在連得上 SEC 嗎？

    抓某一期失敗時用來回答「是網路斷了還是這期資料有問題」。比對照例外
    類別名單可靠得多——名單要跟著 httpx / requests 的版本走，漏一個就
    誤判；直接戳一次是當下的事實。

    刻意用 urllib 不用 httpx：這是在錯誤處理路徑上跑的，不該再依賴
    可能正是出問題那一層的東西。連不上一律回 False，這個函式自己
    不會拋例外。
    """
    import urllib.error
    import urllib.request

    req = urllib.request.Request(
        "https://www.sec.gov/",
        method="HEAD",
        # SEC 擋沒有 User-Agent 的請求，會回 403——那樣會被誤判成連不上
        headers={"User-Agent": "SEC Financial Tools connectivity-check"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout):
            return True
    except urllib.error.HTTPError:
        # 有回應就代表網路是通的，即使狀態碼不是 200
        return True
    except Exception:
        return False


def classify_failure(exc: BaseException,
                     probe: Callable[[], bool] = sec_reachable) -> str:
    """把一次失敗歸類成 `"network"` 或 `"data"`。

    先看例外類型（快，不用等網路）；看不出來就實際戳一次 SEC。
    戳得通代表伺服器有回應、問題出在這份資料本身；戳不通就是網路。

    這個分類只影響「要跟使用者怎麼講」——網路的重抓有救，資料的重抓
    一樣沒救。兩種都不會中止抓取。
    """
    if isinstance(exc, NetworkDownError) or is_network_error(exc):
        return "network"
    return "data" if probe() else "network"


def configure_timeouts(seconds: float = 30.0) -> None:
    """給 edgartools 的 HTTP client 一個明確的逾時。

    edgartools 預設不設 `timeout`（`get_http_config()` 回 None），於是落到
    httpx 自己的 5 秒。抓大份 filing 時 5 秒太短會誤判成逾時，然後被上面
    的重試白白重跑三遍。給 30 秒（連線階段 10 秒）比較貼近實際。

    拿不到 configure_http 就靜靜略過——舊版 edgartools 沒有這個 API，
    那時仍是 httpx 預設值，會慢但不會壞。
    """
    try:
        from edgar import configure_http
    except ImportError:
        return
    try:
        configure_http(timeout=seconds)
    except Exception:
        pass
