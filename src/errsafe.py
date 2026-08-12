"""例外物件的安全萃取——共用模組。

為什麼要獨立一支檔案：main.py 會 import fetcher_*，所以 fetcher_* 不能反過來
import main。這個函式原本只住在 main.py 裡，fetcher_nongaap 拿不到，只好自己
print(f"...{exc!r}")，把完整例外（含 google-generativeai URL 上的 ?key=）印到
stderr。兩份拷貝必然漂移，這裡收成一份讓雙方都 import。
"""


def _exc_status(e: BaseException) -> str:
    """從例外物件安全萃取 HTTP status code，絕不觸碰訊息全文（避免挾帶 URL/response/key）。

    回傳如 ' | HTTP 503'，取不到則回空字串。三家 LLM SDK 與 requests/urllib 的 status
    分別掛在 status_code / code / status / response.status_code 上，逐一探測且只收 int。
    """
    for attr in ("status_code", "code", "status"):
        v = getattr(e, attr, None)
        if isinstance(v, int):
            return f" | HTTP {v}"
    resp = getattr(e, "response", None)
    if resp is not None:
        for attr in ("status_code", "status", "code"):
            v = getattr(resp, attr, None)
            if isinstance(v, int):
                return f" | HTTP {v}"
    return ""
