"""fetch_ledger.py — 記下這一趟抓取有哪幾期沒拿到，以及為什麼。

CTH 2026-08-17 定案的行為：

    照抓、抓不到就留空，**但把缺了哪幾期主動講出來**。

演進過程（免得日後有人改回去）：

  第一版  抓不到就 `except Exception: continue`，靜默。使用者拿到少一季
          的 Excel 不會發現——這是原本要修的 bug。
  第二版  網路問題一律中止整趟、不寫檔。CTH 否決：「不希望抓得太嚴格讓
          資料永遠抓不出來」。
  定案    本模組。缺漏是程式主動報告，不是使用者自己去發現。

「是不是網路問題」不靠猜例外類別名稱（那份名單得跟著 httpx / requests
的版本走，漏一個就誤判），而是失敗當下實際戳一次 SEC：戳得通代表伺服器
有回應、問題在這份資料；戳不通就是網路。每一趟最多戳一次。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from i18n import t
from net_retry import NetworkDownError, is_network_error, sec_reachable

# 摘要裡最多列幾期，其餘用數字帶過。抓 20 年的範圍整個斷線會有上百期，
# 全列出來會擠爆 GUI 與 Index 那一列。
_MAX_LISTED = 6


@dataclass(frozen=True)
class Gap:
    """一期沒拿到的紀錄。

    只存期間標籤與分類——**不存例外訊息**。requests / httpx 的例外訊息
    天生挾帶完整 URL 與 response 片段，落進畫面或 log 就是把敏感資料
    寫上磁碟（見 windows-tool.md「錯誤行怎麼寫」）。
    """
    where: str
    kind: str          # "network" | "data"
    exc_name: str


@dataclass
class FetchLedger:
    """一趟抓取的缺漏帳本。

    Args:
        probe:       判斷「現在連得上 SEC 嗎」，注入點讓測試不必真的連網。
        brake_after: 連續幾期都是網路問題就不再重試（見 give_up_retrying）。
    """
    probe: Callable[[], bool] = sec_reachable
    brake_after: int = 3
    gaps: list[Gap] = field(default_factory=list)

    _consecutive_network: int = 0
    _probe_result: bool | None = None      # None = 這趟還沒戳過

    # ── 記錄 ──────────────────────────────────────────────────────────

    def succeeded(self) -> None:
        """某一期抓到了。網路顯然還在，把連續計數歸零。"""
        self._consecutive_network = 0

    def record(self, where: str, exc: BaseException) -> None:
        """某一期沒抓到。`where` 是期間標籤或申報日，給人看的。"""
        kind = self._classify(exc)
        self.gaps.append(Gap(where=where, kind=kind, exc_name=type(exc).__name__))
        if kind == "network":
            self._consecutive_network += 1
        else:
            # 解析失敗代表伺服器有回應、網路是通的，不該累計成「斷線」
            self._consecutive_network = 0

    def _classify(self, exc: BaseException) -> str:
        # NetworkDownError = 已經退避重試三次都連不上，那就是網路問題，
        # 不必也不該再戳一次 SEC。實測踩過：探測在事後跑，網路可能已經
        # 恢復，於是斷網被報成「SEC 連得上，是資料問題」，方向完全相反。
        # （is_network_error 刻意把 NetworkDownError 排除在外——那個函式
        #  回答的是「要不要再重試一輪」，跟這裡問的不是同一件事。）
        if isinstance(exc, NetworkDownError) or is_network_error(exc):
            return "network"        # 例外自己就說了，不必再戳
        if self._probe_result is None:
            self._probe_result = self._safe_probe()
        return "data" if self._probe_result else "network"

    def _safe_probe(self) -> bool:
        """戳失敗一律當「連得上」——寧可把斷線誤報成資料問題，也不要因為
        探測本身出錯就對使用者宣告網路斷了。"""
        try:
            return bool(self.probe())
        except Exception:
            return True

    # ── 查詢 ──────────────────────────────────────────────────────────

    @property
    def has_gaps(self) -> bool:
        return bool(self.gaps)

    @property
    def network_blamed(self) -> bool:
        """有任何一期是網路造成的。這種缺漏重抓有救，要跟使用者講。"""
        return any(g.kind == "network" for g in self.gaps)

    @property
    def give_up_retrying(self) -> bool:
        """連續多期都是網路問題 = 網路真的斷了，別再退避重試。

        不是中止抓取——剩下的照跑，只是每期失敗得快一點。整個網路斷掉時
        40 份財報各重試 2+4 秒等於乾等 4 分鐘才拿到一份空檔。
        """
        return self._consecutive_network >= self.brake_after

    def summary(self) -> str:
        """一行給人看的摘要，已翻譯。沒缺漏回空字串。

        GUI 與 Excel 的 Index 頁共用同一句——兩邊講的是同一件事，
        分開寫遲早會不一致。
        """
        if not self.gaps:
            return ""
        names = [g.where for g in self.gaps]
        shown = t("xls.meta.sep").join(names[:_MAX_LISTED])
        if len(names) > _MAX_LISTED:
            shown += t("fetch.gaps_ellipsis")
        key = "fetch.gaps_network" if self.network_blamed else "fetch.gaps_data"
        return t(key, n=len(names), periods=shown)
