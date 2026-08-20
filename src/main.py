"""
main.py — SEC Financial Fetcher GUI entry point.

Two-tab Tkinter app:
  Tab 1 (單一公司): Single ticker GAAP fetch
  Tab 2 (批量更新): Batch watchlist update

Persistent buttons: 管理 Watchlist, 進階設定
"""

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.request
from datetime import date
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import i18n
from i18n import t
from config import load_config, save_config, CONFIG_PATH
from errsafe import _exc_status
from excel_writer import write_statements, check_output_writable
from fetcher_gaap import collect_gaps, fetch_gaap_statements, report_progress
from net_retry import configure_timeouts
from output_tables import append_ratio_table, has_any_data

def _build_fixed_height_scrollable(parent, height=110):
    """固定高度的可捲動容器。回傳 (container, inner_frame)——動態內容（如掃描後的

    sheet 勾選清單）塞進 inner_frame，多了就捲動，不會撐大 parent。
    見 project-rules windows-tool-tkinter-ui pattern_fixed_window.py。
    """
    container = tk.Frame(parent, height=height)
    container.pack(fill="x", pady=(4, 8))
    container.pack_propagate(False)

    canvas = tk.Canvas(container, highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    inner_frame = tk.Frame(canvas)
    window_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    inner_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.bind(
        "<Configure>",
        lambda e: canvas.itemconfigure(window_id, width=e.width),
    )

    # Windows 滑鼠滾輪：event.delta 是 120 的倍數，除以 120 換成 yview_scroll 的
    # 「幾格」單位，取負號是因為滾輪往上（delta > 0）要往上捲（scroll 負方向）。
    # 只在滑鼠停在這個容器上時綁定/解除，避免搶走其他 widget（如外層 Notebook）
    # 的滾輪事件。
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_wheel(_event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind_wheel(_event):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _bind_wheel)
    canvas.bind("<Leave>", _unbind_wheel)

    return container, inner_frame

_FINANCIAL_SECTOR_TICKERS = frozenset({"GS", "JPM", "BAC", "C", "WFC", "MS", "BLK", "BX", "KKR"})


# ---- 執行紀錄（log）基礎設施 ----
# 規則見 windows-tool.md「執行紀錄」。核心限制：每次開檔→寫→關檔，不持有 handle（地雷十），
# 不寫 BOM，logs/ 一律在專案根目錄（launcher.ps1 旁），不可寫死 ".."（會污染專案外）。

def _find_project_root() -> str:
    """往上找 launcher.ps1 所在目錄＝專案根目錄。

    不可寫死 os.path.join(SCRIPT_DIR, "..", "logs")：主程式在根目錄的專案會算到專案外層
    （Documents\\Code\\logs），污染其他專案。用這個函式，主程式在根目錄或 src/ 都對，
    日後把 .py 搬進 src/ 也不會壞。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    d = here
    while True:
        if os.path.exists(os.path.join(d, "launcher.ps1")):
            return d
        parent = os.path.dirname(d)
        if parent == d:      # 找到磁碟根目錄仍沒找到，退回自己所在目錄，至少不寫到專案外
            return here
        d = parent


PROJECT_ROOT = Path(_find_project_root())
CACHE_PATH = PROJECT_ROOT / "company_cache.json"

LOG_DIR = os.path.join(_find_project_root(), "logs")
LOG_FILE = os.path.join(LOG_DIR, "app.log")


# =========================================================
# 視窗擺放
# =========================================================
#
# CTH 2026-08-17 回報「視窗很高，起始位置往下切到最下面」。成因是 __init__
# 只給了 geometry("700x650") 沒給座標——位置全交給 Windows，它會用階梯式
# 落點（每開一個新視窗往右下挪一點），開久了就掉到工作列底下。
#
# 解法是自己算座標，而且要對「工作區」算不是對「螢幕」算：螢幕高 1080 但
# 工作列吃掉 40，對螢幕置中的視窗下緣就會被蓋住。

def work_area() -> tuple[int, int, int, int]:
    """回傳 (left, top, right, bottom)——螢幕扣掉工作列後的可用矩形。

    走 Win32 的 SPI_GETWORKAREA。本專案是 Windows 專用工具（見
    README「系統需求」），可以直接用 ctypes，不必為跨平台繞路。
    拿不到就退回整個螢幕——那時視窗可能被工作列切到一點，但不會比
    現在（完全不算座標）更糟。
    """
    try:
        import ctypes
        from ctypes import wintypes

        SPI_GETWORKAREA = 0x0030
        rect = wintypes.RECT()
        ok = ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
        )
        if ok and rect.right > rect.left and rect.bottom > rect.top:
            return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass

    # 退路：問 tkinter 要整個螢幕大小。需要一個 root 才問得到，所以只在
    # 真的走到這裡時才建暫時視窗。
    try:
        probe = tk.Tk()
        probe.withdraw()
        w, h = probe.winfo_screenwidth(), probe.winfo_screenheight()
        probe.destroy()
        return (0, 0, w, h)
    except Exception:
        return (0, 0, 1280, 800)


# 正中央看起來偏低（人眼的視覺重心比幾何中心高）。往上帶一點比較自然。
_VERTICAL_BIAS = 0.35


def _center_on_parent(child, parent) -> None:
    """把子視窗擺在父視窗上方偏中，不要跑到螢幕角落。

    y 取 1/3 而不是 1/2：對話框通常比父視窗矮很多，正中央看起來偏低。
    """
    child.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - child.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - child.winfo_height()) // 3
    child.geometry(f"+{max(x, 0)}+{max(y, 0)}")


# =========================================================
# 覆蓋既有輸出檔
# =========================================================

def existing_outputs(paths) -> list:
    """挑出真的存在的輸出檔，維持原本順序。

    批量更新用：一次列出「這批有哪幾個會被覆蓋」問一次，
    不逐檔跳視窗——12 檔就要點 12 次沒人受得了。
    """
    return [p for p in paths if Path(p).exists()]


def fit_geometry(win_w: int, win_h: int,
                 area: tuple[int, int, int, int]) -> str:
    """算出保證落在工作區內的 tkinter geometry 字串。

    視窗比工作區大就縮到剛好塞得下——寧可矮一點也不要下緣看不到，
    因為看不到的那塊往往正是「執行」按鈕。

    多螢幕時 left/top 可能是負的（副螢幕在主螢幕左邊或上面），
    所有座標都要相對 area 算，不可以假設從 0 開始。
    """
    left, top, right, bottom = area
    avail_w, avail_h = right - left, bottom - top

    w = min(win_w, avail_w)
    h = min(win_h, avail_h)
    x = left + (avail_w - w) // 2
    y = top + int((avail_h - h) * _VERTICAL_BIAS)
    return f"{w}x{h}+{x}+{y}"


def _write_log(msg: str, level: str = "INFO"):
    """寫一行到 logs/app.log。每次開檔→寫→關檔，不持有 handle（地雷十）。"""
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%H:%M:%S')}] [{level:<5}] {msg}\n")
    except OSError:
        pass   # log 掛掉不能拖垮主程式；也涵蓋兩個實例同時跑撞在一起


def _write_log_header(msg: str):
    """任務起始行，唯一有完整日期的行。"""
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"=== {time.strftime('%Y-%m-%d %H:%M:%S')} {msg} ===\n")
    except OSError:
        pass


# _exc_status 已移至 errsafe.py（從這裡 import）。原因：main 會 import fetcher_*，
# fetcher_* 無法反向 import main，函式只住在這裡的話 fetcher 側拿不到、只能自己
# 複製一份，必然漂移——而且已經發生過：fetcher_nongaap 一度直接 print {exc!r}，
# 把 google-generativeai URL 上的 ?key= 印上主控台。


def _migrate_config_if_needed():
    """If old config.json exists in project dir, move it to APPDATA."""
    old_path = PROJECT_ROOT / "config.json"
    if old_path.exists() and not CONFIG_PATH.exists():
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copy2(old_path, CONFIG_PATH)
        old_path.unlink()

PROVIDER_DEFAULTS = {
    "google":    "gemini-flash-latest",
    "openai":    "gpt-4o-mini",
    "anthropic": "claude-haiku-4-5-20251001",
}


# ---- CTH Banner ----

def show_cth_banner():
    b = "\033[90m"; c = "\033[96m"; y = "\033[93m"; r = "\033[0m"
    print(f"{b}/*  ================================  *\\{r}")
    print(f"{b} *                                    *{r}")
    print(f"{b} *    {c}██████╗████████╗██╗  ██╗{b}        *{r}")
    print(f"{b} *   {c}██╔════╝   ██║   ██║  ██║{b}        *{r}")
    print(f"{b} *   {c}██║        ██║   ███████║{b}        *{r}")
    print(f"{b} *   {c}██║        ██║   ██╔══██║{b}        *{r}")
    print(f"{b} *   {c}╚██████╗   ██║   ██║  ██║{b}        *{r}")
    print(f"{b} *    {c}╚═════╝   ╚═╝   ╚═╝  ╚═╝{b}        *{r}")
    print(f"{b} *                                    *{r}")
    print(f"{b} *          {y}created by CTH{b}            *{r}")
    print(f"{b}\\*  ================================  */{r}")
    print()


# ---- App ----


# ── Non-GAAP 暫停開關（2026-08-03）───────────────────────────────────────
#
# Non-GAAP 改走 skill 抽取（TODO B），本工具暫不產出 Data_NonGAAP。
# 兩個 checkbox 因此停用——不停用的話會照常呼叫 AI 抓完 6 季，**抓完才被
# 過濾掉**，等於白燒你的 API 額度。
#
# 相關程式碼（nongaap_layout / metric_rules / 快取）全部保留，改 True 就回來。
NONGAAP_ENABLED = False


# ── Watchlist 的預設群組名稱（2026-08-14）─────────────────────────────────
#
# 這是**存進 config.json 的資料，不是介面文字**，所以不進 i18n。程式到處在
# `g["name"] == UNCATEGORIZED` 比對群組；跟著語言翻譯的話，使用者切成英文後
# 既有的「未分類」就對不上，會再長出一個 "Uncategorized" 空群組，股票留在
# 舊的那個裡看起來像消失了。
#
# 顯示給人看時走 _group_display()，那層才換成當前語言。
UNCATEGORIZED = "未分類"


def _group_display(name: str) -> str:
    """群組名稱的顯示文字。預設群組翻譯，使用者自訂的群組原樣顯示。"""
    return t("gui.wl.uncategorized") if name == UNCATEGORIZED else name


def _group_stored(display: str) -> str:
    """顯示名稱換回儲存名稱，_group_display 的反向。

    下拉選單顯示的是譯名，但寫進 config.json 的必須永遠是 UNCATEGORIZED——
    少了這層，日文使用者新增公司會建出一個叫「未分類」的日文群組，跟既有的
    那個並存，股票分散在兩邊。
    """
    return UNCATEGORIZED if display == t("gui.wl.uncategorized") else display


# 抽到 output_tables.py（2026-08-07）讓 cli.py 共用同一份組裝邏輯。
# 這裡保留舊名稱的別名，呼叫端與既有測試都不必改。
_append_ratio_table = append_ratio_table



class SECFetcherApp:
    """Two-tab SEC financial fetcher UI.

    All background work runs in daemon threads; results are pushed to msg_queue
    and applied in _poll_queue every 100ms because Tkinter is not thread-safe.
    """

    @property
    def TICKER_PH(self) -> str:
        """Ticker 輸入框的 placeholder。

        必須是 property 不能是類別屬性——類別屬性在 import 時就求值，那時
        set_lang() 還沒跑（語言是 __init__ 讀 config 之後才設的），值會凍結
        在預設語言。所有呼叫端都是 self.TICKER_PH，改成 property 不必動。
        """
        return t("gui.lbl.ticker_placeholder")

    def __init__(self, root: tk.Tk):
        """Load config, initialise state, build UI, start the 100ms queue poll."""
        self.root = root
        self.root.title("SEC Financial Fetcher")
        # 明確呼叫一次 geometry() 就會永久關閉 tkinter 的自動撐高（geometry
        # propagation）——快速掃描跳出的可選 Sheet 清單原本會撐高視窗，改用
        # _build_sheet_panel 裡的固定高度可捲動容器承接，不會再撐開這裡鎖定的尺寸。
        # resizable(True, True) 保留、不衝突：那只管使用者能不能手動拖邊框。
        #
        # 座標一定要自己算（見模組頂端 fit_geometry 的註解）：只給大小不給座標時
        # Windows 會用階梯式落點，視窗越開越低，最後下緣掉到工作列底下。
        #
        # 寬高改成算工作區的比例，不再是寫死的 900x720（2026-08-18，TODO E4/E16
        # CTH 回報「視窗變寬了，內容卻沒跟著撐開」）。固定像素在不同螢幕上感受
        # 差很多：CTH 的螢幕上 900px 只佔畫面一小塊，內容再怎麼撐都撐不滿；比例
        # 才能在任何螢幕上維持「約 2/3 寬度」的觀感。真正讓內容用滿寬度的是
        # `tab.columnconfigure(0, weight=1)`（見 `_build_tab1`／`_build_settings_panel`）
        # ——只放大視窗不夠，欄位本身要有 weight 才會跟著撐開，這是原本沒撐開的
        # 根因。高度一併拉高（TODO E6 決定的方向：拉高視窗＋捲動），給「處理
        # 進度」log 區更多空間，見 `_TAB3_HEIGHT` 旁的重新量測記錄。
        _area = work_area()
        _avail_w = _area[2] - _area[0]
        _avail_h = _area[3] - _area[1]
        self._WIN_W = min(max(int(_avail_w * self._WIN_W_RATIO), self._WIN_W_MIN), self._WIN_W_MAX)
        self._WIN_H = min(max(int(_avail_h * self._WIN_H_RATIO), self._WIN_H_MIN), self._WIN_H_MAX)
        _geom = fit_geometry(self._WIN_W, self._WIN_H, _area)
        self.root.geometry(_geom)
        # minsize 不可以大於實際擺出來的尺寸，否則小螢幕上 fit_geometry 縮好的
        # 視窗會被 minsize 又頂回去、重新出界。
        _fitted_w, _fitted_h = (int(v) for v in _geom.split("+")[0].split("x"))
        self.root.minsize(min(760, _fitted_w), min(560, _fitted_h))
        self.root.resizable(True, True)

        _migrate_config_if_needed()
        self.cfg = load_config(CONFIG_PATH)
        # 語言必須在建任何 widget 之前設好——t() 是在建置時查一次表，
        # 設晚了介面會停在預設語言。不認得的代號 set_lang 會自己退回繁中。
        i18n.set_lang(self.cfg.get("language"))
        self.msg_queue: queue.Queue = queue.Queue()
        self.is_running = False
        # Runtime state for popups
        self._wl_found_name = ""
        self._wl_list_container = None
        self.wl_lookup_label = None
        self.wl_add_btn = None
        self.wl_cache_label = None
        self.wl_add_var = None
        self.wl_group_var: tk.StringVar | None = None
        self.wl_group_cb = None
        self._wl_draft: dict = {}
        self._wl_group_collapsed: dict[str, bool] = {}
        self._last_output_folder: Path | None = None
        self.settings_identity_var = None
        self.settings_provider_var = None
        self.settings_model_var = None
        self.settings_key_var = None
        self.settings_key_entry = None
        self.settings_key_toggle_btn = None
        self.settings_outdir_var = None
        self.settings_test_label = None
        self.settings_fmt_var = None
        self.settings_max_filings_var = None
        self.settings_saved_label = None
        self.nongaap_warn_label = None
        self.btn_confirm_company = None
        self.tab1_name_label = None
        self._scan_hint_label = None
        self.tab1_outdir_var = None
        self.tab1_fmt_var = None
        self.tab1_custom_var = None
        self.tab1_custom_entry = None
        self.tab1_preview_label = None
        self.tab1_fetch_q_var: tk.BooleanVar | None = None
        self.tab1_fetch_k_var: tk.BooleanVar | None = None
        self.tab1_start_year_var: tk.StringVar | None = None
        self.tab1_end_year_var: tk.StringVar | None = None
        self.batch_fetch_q_var: tk.BooleanVar | None = None
        self.batch_fetch_k_var: tk.BooleanVar | None = None
        self.batch_start_year_var: tk.StringVar | None = None
        self.batch_end_year_var: tk.StringVar | None = None
        self._sheet_check_vars: dict[str, tk.BooleanVar] = {}
        self._sheet_panel_frame: tk.Frame | None = None
        self._scan_btn: ttk.Button | None = None
        self._tab1_adv_collapsed: bool = True
        self._tab1_adv_frame = None
        self._tab1_adv_toggle_btn = None
        self._tab2_adv_collapsed: bool = True
        self._tab2_adv_frame = None
        self._tab2_adv_toggle_btn = None

        self._build_ui()
        self._poll_queue()

    # =========================================================
    # UI Construction
    # =========================================================

    def _build_ui(self):
        """Build full window layout: notebook (row 0), persistent buttons (row 1), log (row 2), open-folder (row 3)."""
        pad = {"padx": 14, "pady": 6}

        # Global font — 11pt for all ttk widgets
        style = ttk.Style()
        style.configure(".", font=("", 11))
        style.configure("TNotebook.Tab", font=("", 11))

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="ew", **pad)
        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()
        self._update_identity_warnings()

        # 「管理 Watchlist」原本放在這裡（root 層級、row=1，跨頁籤都看得到），
        # TODO E19：這顆只跟批量更新有關，Tab1 單一公司用不到，搬進
        # `_build_tab2` 最上方了。row=1 空出來不影響版面（沒設 weight，高度
        # 是 0），`frame_log` 維持在 row=2 不用跟著搬。

        # Progress log
        frame_log = ttk.LabelFrame(self.root, text=t("gui.frame.progress"), padding=8)
        frame_log.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 4))
        self.progress_label = ttk.Label(frame_log, text=t("gui.status.idle"))
        self.progress_label.pack(anchor="w")
        self.progress_bar = ttk.Progressbar(frame_log, mode="determinate", length=440)
        self.progress_bar.pack(fill="x", pady=(4, 8))
        self.log_text = scrolledtext.ScrolledText(
            frame_log, width=60, height=10, state="disabled", font=("Consolas", 10)
        )
        self.log_text.pack(fill="both", expand=True)

        # Open folder button (shown after completion)
        frame_output = tk.Frame(self.root)
        frame_output.grid(row=3, column=0, pady=(0, 12))
        self.btn_open_folder = ttk.Button(
            frame_output, text=t("gui.btn.open_output_folder"), command=self._open_output_folder
        )

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

    def _build_tab1(self):
        """Build Tab 1 (單一公司): ticker input, GAAP/Non-GAAP toggles, output settings, run button."""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=t("gui.tab.single"))
        # 沒有這行，視窗變寬時這個欄位不會跟著撐開——sticky="ew" 只管子元件
        # 貼齊欄位邊界，欄位本身要有 weight 才會分到多出來的寬度（TODO E16）。
        tab.columnconfigure(0, weight=1)

        # Row 0: Ticker + inline company name
        row_ticker = ttk.Frame(tab)
        row_ticker.grid(row=0, column=0, sticky="ew", pady=4)
        ttk.Label(row_ticker, text="Ticker:").pack(side="left", padx=(0, 8))
        self.ticker_var = tk.StringVar()
        self.ticker_entry = ttk.Entry(row_ticker, textvariable=self.ticker_var, width=18, foreground="grey")
        self.ticker_entry.pack(side="left")
        self.ticker_var.set(self.TICKER_PH)
        self.ticker_entry.bind("<FocusIn>",  lambda e: self._ph_in(self.ticker_entry, self.ticker_var, self.TICKER_PH))
        self.ticker_entry.bind("<FocusOut>", lambda e: self._on_ticker_focusout(e))
        self.ticker_entry.bind("<Return>",   lambda e: self._confirm_company())
        self.btn_confirm_company = None
        self.tab1_name_label = ttk.Label(row_ticker, text="", foreground="#555555")
        self.tab1_name_label.pack(side="left", padx=(10, 0))
        self._scan_btn = ttk.Button(row_ticker, text=t("gui.btn.scan"), command=self._run_preview_scan, width=16)
        self._scan_btn.pack(side="left", padx=(12, 0))
        _scan_help_lbl = tk.Label(row_ticker, text="？", foreground="#0078D4", cursor="hand2",
                                   font=("Microsoft JhengHei", 11, "bold"))
        _scan_help_lbl.pack(side="left", padx=(4, 0))
        _scan_help_lbl.bind("<Button-1>", lambda e: self._show_scan_help())
        # 掃描要打 EDGAR 抓最新一份 10-Q，5~15 秒。公司名稱卻走本機快取幾乎瞬間
        # 回來——CTH 就是看到名稱跳出來、以為這次點擊只做了名稱，等不到期間便
        # 再按一次（2026-08-17 回報）。用一個獨立的提示標籤講明「還在跑、要多久」，
        # 不塞進按鈕文字：按鈕寬度會跟著字數變，一長一短版面會左右跳。
        self._scan_hint_label = ttk.Label(row_ticker, text="", foreground="#0078D4")
        self._scan_hint_label.pack(side="left", padx=(8, 0))

        # Row 1: SEC Identity warning (hidden unless cfg["identity"] is empty)
        self.identity_warn_label = tk.Label(
            tab, text=t("gui.lbl.identity_missing"),
            foreground="orange", cursor="hand2", font=("", 10)
        )
        self.identity_warn_label.grid(row=1, column=0, sticky="w", padx=2)
        self.identity_warn_label.bind("<Button-1>", lambda e: self._goto_settings_tab())
        self.identity_warn_label.grid_remove()

        # Row 2: Checkboxes
        row_type = ttk.Frame(tab)
        row_type.grid(row=2, column=0, sticky="ew", pady=4)
        self.fetch_gaap_var    = tk.BooleanVar(value=True)
        self.fetch_nongaap_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row_type, text=t("gui.chk.gaap"),               variable=self.fetch_gaap_var).pack(side="left", padx=(0, 16))
        _ng_text = (t("gui.chk.nongaap") if NONGAAP_ENABLED
                    else t("gui.chk.nongaap_paused"))
        _ng_cb = ttk.Checkbutton(row_type, text=_ng_text, variable=self.fetch_nongaap_var)
        if not NONGAAP_ENABLED:
            _ng_cb.state(["disabled"])
        _ng_cb.pack(side="left")
        self.fetch_nongaap_var.trace_add("write", self._on_nongaap_toggle)

        # Row 3: 報表類型 toggle
        adv_toggle_row1 = ttk.Frame(tab)
        adv_toggle_row1.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self._tab1_adv_toggle_btn = ttk.Button(adv_toggle_row1, text=t("gui.btn.report_type_collapsed"),
                                                command=self._toggle_tab1_adv, width=12)
        self._tab1_adv_toggle_btn.pack(side="left")

        # Row 4: 報表類型 content (hidden by default)
        self._tab1_adv_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=(8, 4))
        self._tab1_adv_frame.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        self.tab1_fetch_q_var = tk.BooleanVar(value=True)
        self.tab1_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self._tab1_adv_frame, text=t("gui.chk.quarterly"), variable=self.tab1_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(self._tab1_adv_frame, text=t("gui.chk.annual"), variable=self.tab1_fetch_k_var).pack(side="left")
        self._tab1_adv_frame.grid_remove()

        # Row 5: Date range
        row_date = ttk.Frame(tab)
        row_date.grid(row=5, column=0, sticky="ew", pady=(2, 4))
        ttk.Label(row_date, text=t("gui.lbl.year_from")).pack(side="left", padx=(0, 4))
        self.tab1_start_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_start_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text=t("gui.lbl.year_to")).pack(side="left", padx=(8, 4))
        self.tab1_end_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_end_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text=t("gui.lbl.year_hint"), foreground="#555555").pack(side="left", padx=(4, 0))

        # Row 6: Sheet selection panel (hidden until scan completes)
        # 最新季度／送件日不能另開一行 Label——這個視窗鎖死 650px 高、不會自動撐大
        # （見 __init__ 的 geometry() 註解），下面「處理進度」的 log 區已經很緊繃，
        # 實測顯示 sheet 面板一展開，log 可視高度就只剩個位數 px；多加一行 23px
        # 會直接把 log 擠到全隱形。改寫進 LabelFrame 自己的標題列，不佔新的一行，
        # 高度成本是 0
        self._SHEET_PANEL_TITLE_BASE = t("gui.frame.optional_sheets")
        self._sheet_panel_frame = ttk.LabelFrame(tab, text=self._SHEET_PANEL_TITLE_BASE, padding=6)
        self._sheet_panel_frame.grid(row=6, column=0, sticky="ew", pady=(0, 4))
        _, self._sheet_panel_inner = _build_fixed_height_scrollable(self._sheet_panel_frame, height=60)
        self._sheet_panel_frame.grid_remove()

        # Row 7: Non-GAAP warning (hidden by default)
        self.nongaap_warn_label = ttk.Label(
            tab, text=t("gui.lbl.nongaap_need_key"),
            foreground="orange", font=("", 10)
        )
        self.nongaap_warn_label.grid(row=7, column=0, sticky="w", padx=2)
        self.nongaap_warn_label.grid_remove()

        # Row 8: Output settings toggle
        self._out_collapsed = False
        out_toggle_row = ttk.Frame(tab)
        out_toggle_row.grid(row=8, column=0, sticky="ew", pady=(8, 0))
        self._out_toggle_btn = ttk.Button(out_toggle_row, text=t("gui.btn.output_expanded"),
                                           command=self._toggle_out_settings, width=12)
        self._out_toggle_btn.pack(side="left")

        # Row 9: Output settings content (collapsible)
        out_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=8)
        out_frame.grid(row=9, column=0, sticky="ew", pady=(0, 4))
        out_frame.columnconfigure(0, weight=1)
        self._out_settings_frame = out_frame

        # Storage location row — 路徑欄用 fill/expand 撐滿剩餘寬度，不再是固定
        # 26 字元寬。路徑常常比 26 字元長（截圖裡 CTH 的路徑就被切掉看不到尾巴），
        # 視窗變寬時這欄本來就該優先分到多出來的空間
        loc_row = ttk.Frame(out_frame)
        loc_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(loc_row, text=t("gui.lbl.save_location")).pack(side="left")
        self.tab1_outdir_var = tk.StringVar(value=self.cfg.get("output_dir", "output"))
        ttk.Entry(loc_row, textvariable=self.tab1_outdir_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(loc_row, text=t("gui.btn.browse"), width=5, command=self._browse_output_dir).pack(side="left")

        # Filename format radios
        ttk.Label(out_frame, text=t("gui.lbl.filename_format")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.tab1_fmt_var = tk.StringVar(value=self.cfg.get("filename_format", "ticker_name"))

        ttk.Radiobutton(out_frame, text=t("gui.radio.name_ticker_company"),
                        variable=self.tab1_fmt_var, value="ticker_name",
                        command=self._on_tab1_fmt_change).grid(row=2, column=0, sticky="w", padx=(16, 0))
        ttk.Radiobutton(out_frame, text=t("gui.radio.name_ticker_only"),
                        variable=self.tab1_fmt_var, value="ticker_only",
                        command=self._on_tab1_fmt_change).grid(row=3, column=0, sticky="w", padx=(16, 0))

        custom_row = ttk.Frame(out_frame)
        custom_row.grid(row=4, column=0, sticky="w", padx=(16, 0))
        ttk.Radiobutton(custom_row, text=t("gui.radio.name_custom"),
                        variable=self.tab1_fmt_var, value="custom",
                        command=self._on_tab1_fmt_change).pack(side="left")
        self.tab1_custom_var = tk.StringVar(value=self.cfg.get("filename_custom", ""))
        is_custom = self.tab1_fmt_var.get() == "custom"
        self.tab1_custom_entry = ttk.Entry(custom_row, textvariable=self.tab1_custom_var, width=22,
                                           state="normal" if is_custom else "disabled")
        self.tab1_custom_entry.pack(side="left", padx=(4, 4))
        ttk.Label(custom_row, text=".xlsx", foreground="gray").pack(side="left")
        self.tab1_custom_var.trace_add("write", lambda *_: self._update_tab1_preview())

        # Preview label
        self.tab1_preview_label = ttk.Label(out_frame, text="", foreground="#555555", font=("", 10))
        self.tab1_preview_label.grid(row=5, column=0, sticky="w", pady=(6, 0))
        self._update_tab1_preview()

        # 設為預設：這裡改的資料夾/檔名格式只影響這次執行，不再像過去那樣
        # 一改就悄悄寫回全域 config（Tab2 批次完全看不到那個變化卻共用它）。
        # 要讓這次的選擇變成下次開程式的預設值，得按這顆按鈕明確存檔。
        default_row = ttk.Frame(out_frame)
        default_row.grid(row=6, column=0, sticky="w", pady=(4, 0))
        ttk.Button(default_row, text=t("gui.btn.set_as_default"), width=10,
                   command=self._set_tab1_output_as_default).pack(side="left")
        self.tab1_default_saved_label = ttk.Label(default_row, text="", foreground="#1a7a34")
        self.tab1_default_saved_label.pack(side="left", padx=8)
        # 這顆按鈕的意義不會從按鈕文字本身看出來，一律顯示這行小字說明（不用點才看得到）
        ttk.Label(out_frame, text=t("gui.lbl.set_as_default_hint"),
                  foreground="#888888", font=("", 9)).grid(row=7, column=0, sticky="w", pady=(0, 2))

        # Row 10: Execute button
        self.btn_run_single = ttk.Button(tab, text=t("gui.btn.run"), command=self._run_single, width=16)
        self.btn_run_single.grid(row=10, column=0, pady=(8, 4))

    def _toggle_tab1_adv(self):
        self._tab1_adv_collapsed = not self._tab1_adv_collapsed
        if self._tab1_adv_collapsed:
            self._tab1_adv_frame.grid_remove()
            self._tab1_adv_toggle_btn.config(text=t("gui.btn.report_type_collapsed"))
        else:
            self._tab1_adv_frame.grid()
            self._tab1_adv_toggle_btn.config(text=t("gui.btn.report_type_expanded"))

    def _toggle_tab2_adv(self):
        self._tab2_adv_collapsed = not self._tab2_adv_collapsed
        if self._tab2_adv_collapsed:
            self._tab2_adv_frame.grid_remove()
            self._tab2_adv_toggle_btn.config(text=t("gui.btn.report_type_collapsed"))
        else:
            self._tab2_adv_frame.grid()
            self._tab2_adv_toggle_btn.config(text=t("gui.btn.report_type_expanded"))

    def _toggle_out_settings(self):
        """Collapse or expand the output-settings section, updating the toggle button arrow."""
        self._out_collapsed = not self._out_collapsed
        if self._out_collapsed:
            self._out_settings_frame.grid_remove()
            self._out_toggle_btn.config(text=t("gui.btn.output_collapsed"))
        else:
            self._out_settings_frame.grid()
            self._out_toggle_btn.config(text=t("gui.btn.output_expanded"))

    def _on_nongaap_toggle(self, *_args):
        """Show the API Key warning when Non-GAAP is enabled but api_key is not yet configured."""
        if self.fetch_nongaap_var.get() and not self.cfg["ai"].get("api_key"):
            self.nongaap_warn_label.grid()
        else:
            self.nongaap_warn_label.grid_remove()

    def _on_batch_nongaap_toggle(self, *_args):
        if self.batch_nongaap_var.get() and not self.cfg["ai"].get("api_key"):
            self.batch_nongaap_warn.pack(side="left", padx=8)
        else:
            self.batch_nongaap_warn.pack_forget()

    def _build_tab2(self):
        """Build Tab 2 (批量更新): scrollable group-organised watchlist + batch run button."""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=t("gui.tab.batch"))
        tab.columnconfigure(0, weight=1)

        # Row 0: 管理 Watchlist（TODO E19：從跨頁籤都看得到的持久按鈕搬進來，
        # 這顆只跟批量更新有關，Tab1 單一公司用不到，搬到這裡最上方最直覺）
        ttk.Button(tab, text=t("gui.btn.manage_watchlist"),
                   command=self._open_watchlist_popup, width=18).grid(
            row=0, column=0, sticky="w", pady=(0, 4))

        self.tab2_list_frame = ttk.LabelFrame(tab, text=" Watchlist ", padding=6)
        self.tab2_list_frame.grid(row=1, column=0, sticky="ew", pady=4)
        self.tab2_list_frame.columnconfigure(0, weight=1)

        tab2_canvas = tk.Canvas(self.tab2_list_frame, height=150, highlightthickness=0)
        tab2_sb = ttk.Scrollbar(self.tab2_list_frame, orient="vertical", command=tab2_canvas.yview)
        tab2_canvas.configure(yscrollcommand=tab2_sb.set)
        tab2_canvas.grid(row=0, column=0, sticky="ew")
        tab2_sb.grid(row=0, column=1, sticky="ns")
        tab2_inner = ttk.Frame(tab2_canvas)
        tab2_win = tab2_canvas.create_window((0, 0), window=tab2_inner, anchor="nw")
        tab2_inner.bind("<Configure>", lambda e: (
            tab2_canvas.configure(scrollregion=tab2_canvas.bbox("all")),
            tab2_canvas.itemconfig(tab2_win, width=tab2_canvas.winfo_width()),
        ))
        tab2_canvas.bind("<Configure>", lambda e: tab2_canvas.itemconfig(tab2_win, width=e.width))
        self._tab2_canvas = tab2_canvas
        self._tab2_inner = tab2_inner

        self.tab2_check_vars: dict[str, tk.BooleanVar] = {}
        self._refresh_tab2_list()

        # Row 2: SEC Identity warning (hidden unless cfg["identity"] is empty)
        self.tab2_identity_warn_label = tk.Label(
            tab, text=t("gui.lbl.identity_missing"),
            foreground="orange", cursor="hand2", font=("", 10)
        )
        self.tab2_identity_warn_label.grid(row=2, column=0, sticky="w", padx=2, pady=(0, 2))
        self.tab2_identity_warn_label.bind("<Button-1>", lambda e: self._goto_settings_tab())
        self.tab2_identity_warn_label.grid_remove()

        # Row 3: 唯讀顯示目前全域輸出預設——批次抓出來的檔案就是照這個值落地，
        # Tab1 過去可以悄悄改掉這個全域值卻讓 Tab2 完全看不到，這裡至少讓
        # 批次使用者知道檔案會存去哪
        self.tab2_output_default_label = ttk.Label(tab, text="", foreground="#555555", font=("", 10))
        self.tab2_output_default_label.grid(row=3, column=0, sticky="w", pady=(0, 4))
        self._refresh_output_default_display()

        row_sel = ttk.Frame(tab)
        row_sel.grid(row=4, column=0, sticky="w", pady=4)
        ttk.Button(row_sel, text=t("gui.btn.select_all"),   command=self._select_all,   width=8).pack(side="left", padx=(0, 8))
        ttk.Button(row_sel, text=t("gui.btn.select_none"), command=self._deselect_all, width=8).pack(side="left")

        row_opts = ttk.Frame(tab)
        row_opts.grid(row=5, column=0, sticky="w", pady=(4, 0))
        self.batch_nongaap_var = tk.BooleanVar(value=False)
        _bng_text = (t("gui.chk.batch_nongaap") if NONGAAP_ENABLED
                     else t("gui.chk.batch_nongaap_paused"))
        _bng_cb = ttk.Checkbutton(row_opts, text=_bng_text, variable=self.batch_nongaap_var)
        if not NONGAAP_ENABLED:
            _bng_cb.state(["disabled"])
        _bng_cb.pack(side="left")
        self.batch_nongaap_warn = ttk.Label(
            row_opts, text=t("gui.lbl.need_api_key"), foreground="#cc8800"
        )
        self.batch_nongaap_warn.pack(side="left", padx=8)
        self.batch_nongaap_warn.pack_forget()
        self.batch_nongaap_var.trace_add("write", self._on_batch_nongaap_toggle)

        # Row 6: 報表類型 toggle
        adv_toggle_row2 = ttk.Frame(tab)
        adv_toggle_row2.grid(row=6, column=0, sticky="ew", pady=(4, 0))
        self._tab2_adv_toggle_btn = ttk.Button(adv_toggle_row2, text=t("gui.btn.report_type_collapsed"),
                                                command=self._toggle_tab2_adv, width=12)
        self._tab2_adv_toggle_btn.pack(side="left")

        # Row 7: 報表類型 content (hidden by default)
        self._tab2_adv_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=(8, 4))
        self._tab2_adv_frame.grid(row=7, column=0, sticky="ew", pady=(0, 4))
        self.batch_fetch_q_var = tk.BooleanVar(value=True)
        self.batch_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self._tab2_adv_frame, text=t("gui.chk.quarterly"), variable=self.batch_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(self._tab2_adv_frame, text=t("gui.chk.annual"), variable=self.batch_fetch_k_var).pack(side="left")
        self._tab2_adv_frame.grid_remove()

        # Row 8: Date range
        row_date2 = ttk.Frame(tab)
        row_date2.grid(row=8, column=0, sticky="ew", pady=(2, 0))
        ttk.Label(row_date2, text=t("gui.lbl.year_from")).pack(side="left", padx=(0, 4))
        self.batch_start_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date2, from_=1993, to=2099, textvariable=self.batch_start_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date2, text=t("gui.lbl.year_to")).pack(side="left", padx=(8, 4))
        self.batch_end_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date2, from_=1993, to=2099, textvariable=self.batch_end_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date2, text=t("gui.lbl.year_hint"), foreground="#555555").pack(side="left", padx=(4, 0))

        self.btn_run_batch = ttk.Button(tab, text=t("gui.btn.run_batch"), command=self._run_batch, width=20)
        self.btn_run_batch.grid(row=9, column=0, pady=(8, 4))

    # =========================================================
    # Placeholder helpers
    # =========================================================

    def _ph_in(self, entry, var, placeholder):
        """Clear placeholder text when the user focuses into an entry field."""
        if var.get() == placeholder:
            var.set("")
            entry.configure(foreground="black")
            if entry is self.ticker_entry and self.tab1_name_label:
                self.tab1_name_label.config(text="")

    def _ph_out(self, entry, var, placeholder):
        """Restore placeholder text (grey) when user leaves an empty entry field."""
        if not var.get().strip():
            var.set(placeholder)
            entry.configure(foreground="grey")

    def _on_ticker_focusout(self, event):
        """Restore placeholder and spawn a background company-name lookup when user leaves the ticker field."""
        self._ph_out(self.ticker_entry, self.ticker_var, self.TICKER_PH)
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            if self.tab1_name_label:
                self.tab1_name_label.config(text="")
            self._update_tab1_preview()
            return
        if self.tab1_name_label:
            self.tab1_name_label.config(text=t("gui.status.looking_up"), foreground="#555555")
        if self.btn_confirm_company:
            self.btn_confirm_company.config(state="disabled")
        self._update_tab1_preview()
        threading.Thread(target=lambda: self._tab1_lookup_worker(ticker), daemon=True).start()

    def _confirm_company(self):
        """Trigger company-name lookup manually (bound to Enter key in ticker entry)."""
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            return
        if self.tab1_name_label:
            self.tab1_name_label.config(text=t("gui.status.looking_up"), foreground="#555555")
        if self.btn_confirm_company:
            self.btn_confirm_company.config(state="disabled")
        threading.Thread(target=lambda: self._tab1_lookup_worker(ticker), daemon=True).start()

    def _tab1_lookup_worker(self, ticker: str):
        """Background thread: resolve company name for Tab 1 inline display (local cache first, then live EDGAR)."""
        # Check local cache first
        if CACHE_PATH.exists():
            try:
                with open(CACHE_PATH, encoding="utf-8") as f:
                    companies = json.load(f).get("companies", {})
                if ticker in companies:
                    self.msg_queue.put(("tab1_name_result", ("ok", ticker, companies[ticker])))
                    return
            except (json.JSONDecodeError, OSError):
                pass
        # Fallback: live EDGAR query
        try:
            from edgar import Company, set_identity
            identity = self.cfg.get("identity") or "SEC Tool sec@example.com"
            set_identity(identity)
            c = Company(ticker)
            name = c.name or ""
            if name:
                self.msg_queue.put(("tab1_name_result", ("ok", ticker, name)))
            else:
                self.msg_queue.put(("tab1_name_result", ("notfound", ticker, "")))
        except Exception:
            self.msg_queue.put(("tab1_name_result", ("notfound", ticker, "")))

    def _get_ph_value(self, var, placeholder) -> str:
        v = var.get().strip()
        return "" if v == placeholder else v

    def _on_tab1_fmt_change(self):
        """Enable/disable the custom filename entry and refresh the preview when format radio changes.

        只影響這次執行要用的值，不再悄悄寫回全域 config——要存成下次預設得按
        「設為預設」（`_set_tab1_output_as_default`）。
        """
        is_custom = self.tab1_fmt_var.get() == "custom"
        if self.tab1_custom_entry:
            self.tab1_custom_entry.config(state="normal" if is_custom else "disabled")
        self._update_tab1_preview()

    def _update_tab1_preview(self):
        """Refresh the filename preview label based on current format setting and ticker."""
        if not self.tab1_preview_label:
            return
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper() if self.ticker_var else ""
        fmt = self.tab1_fmt_var.get() if self.tab1_fmt_var else "ticker_name"
        if fmt == "ticker_name":
            if ticker:
                name = self._lookup_company_name(ticker)
                safe_name = re.sub(r'[\\/:*?"<>|]', "", name).strip()
                preview = f"{ticker} {safe_name} data.xlsx"
            else:
                preview = t("gui.lbl.filename_sample")
        elif fmt == "ticker_only":
            preview = f"{ticker}.xlsx" if ticker else "TICKER.xlsx"
        else:  # custom
            custom = self.tab1_custom_var.get().strip() if self.tab1_custom_var else ""
            preview = f"{custom}.xlsx" if custom else t("gui.lbl.filename_empty")
        self.tab1_preview_label.config(text=t("gui.lbl.preview", name=preview))

    def _browse_output_dir(self):
        """Open folder picker; only remembers the folder per-ticker, not as the global default.

        全域預設值只在按「設為預設」（`_set_tab1_output_as_default`）時才會動，
        這裡挑的資料夾只影響這次執行，跟過去「挑了就悄悄變全域預設」的行為不同。
        """
        from tkinter import filedialog
        current = self.tab1_outdir_var.get().strip() if self.tab1_outdir_var else "output"
        initial = str(PROJECT_ROOT / current) if not os.path.isabs(current) else current
        folder = filedialog.askdirectory(title=t("gui.dlg.choose_output_dir"), initialdir=initial)
        if folder:
            self.tab1_outdir_var.set(folder)
            # 記住這個 ticker 的路徑（跟全域預設是兩件事）
            ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
            if ticker:
                if "ticker_paths" not in self.cfg:
                    self.cfg["ticker_paths"] = {}
                self.cfg["ticker_paths"][ticker] = folder
                save_config(self.cfg, CONFIG_PATH)

    def _save_tab1_output_settings(self):
        """Persist Tab 1 output settings (dir, filename format, custom name) to config.json as the new global default."""
        if self.tab1_outdir_var:
            self.cfg["output_dir"] = self.tab1_outdir_var.get().strip() or "output"
        if self.tab1_fmt_var:
            self.cfg["filename_format"] = self.tab1_fmt_var.get()
        if self.tab1_custom_var:
            self.cfg["filename_custom"] = self.tab1_custom_var.get().strip()
        save_config(self.cfg, CONFIG_PATH)

    def _set_tab1_output_as_default(self):
        """「設為預設」按鈕：把目前輸出資料夾/檔名格式明確存成下次開程式的全域預設值。"""
        self._save_tab1_output_settings()
        self._refresh_output_default_display()
        if getattr(self, "tab1_default_saved_label", None) is not None:
            self.tab1_default_saved_label.config(text=t("gui.lbl.output_default_saved"))
            self.root.after(3000, lambda: self.tab1_default_saved_label.config(text=""))

    def _refresh_output_default_display(self):
        """更新 Tab2 那行唯讀文字，顯示目前全域輸出預設資料夾（Tab1「設為預設」改了要跟著更新）。"""
        if getattr(self, "tab2_output_default_label", None) is not None:
            self.tab2_output_default_label.config(
                text=t("gui.lbl.current_default_output", dir=self.cfg.get("output_dir", "output"))
            )

    def _goto_settings_tab(self):
        """SEC Identity 警告文字被點擊時，切到 Tab3（進階設定）。"""
        self.notebook.select(2)

    def _update_identity_warnings(self):
        """依 cfg["identity"] 是否已填，切換 Tab1/Tab2 的 SEC Identity 提示顯示。"""
        missing = not self.cfg.get("identity")
        for label in (getattr(self, "identity_warn_label", None),
                      getattr(self, "tab2_identity_warn_label", None)):
            if label is None:
                continue
            if missing:
                label.grid()
            else:
                label.grid_remove()

    # =========================================================
    # Tab 2 watchlist list
    # =========================================================

    def _refresh_tab2_list(self):
        """Rebuild Tab 2 watchlist display with group headers and per-group select/deselect buttons."""
        for w in self._tab2_inner.winfo_children():
            w.destroy()
        self.tab2_check_vars = {}
        watchlist = self.cfg.get("watchlist", [])
        if not watchlist:
            ttk.Label(self._tab2_inner, text=t("gui.lbl.watchlist_empty"),
                      foreground="gray").grid(row=0, column=0, columnspan=4, sticky="w")
            self._tab2_inner.update_idletasks()
            self._tab2_canvas.configure(scrollregion=self._tab2_canvas.bbox("all"))
            return

        self._ensure_groups(self.cfg)
        groups = self._get_groups_sorted(self.cfg)
        wl_set = {w["ticker"] for w in watchlist}
        cols = 3
        grid_row = 0

        for group in groups:
            gname = group["name"]
            tickers = sorted(t for t in group["tickers"] if t in wl_set)
            if not tickers:
                continue
            # Group header
            hdr = ttk.Frame(self._tab2_inner)
            hdr.grid(row=grid_row, column=0, columnspan=cols + 1, sticky="ew", pady=(6, 2))
            ttk.Label(hdr, text=_group_display(gname), font=("", 11, "bold"),
                      foreground="#333").pack(side="left")
            ttk.Button(hdr, text=t("gui.btn.select_all"), width=5,
                       command=lambda ts=tickers: self._select_group(ts, True)).pack(side="left", padx=(8, 2))
            ttk.Button(hdr, text=t("gui.btn.select_none"), width=6,
                       command=lambda ts=tickers: self._select_group(ts, False)).pack(side="left")
            grid_row += 1
            # Ticker checkboxes
            for i, ticker in enumerate(tickers):
                var = tk.BooleanVar(value=True)
                self.tab2_check_vars[ticker] = var
                r, c = divmod(i, cols)
                ttk.Checkbutton(self._tab2_inner, text=ticker, variable=var).grid(
                    row=grid_row + r, column=c, sticky="w", padx=8, pady=2)
            grid_row += (len(tickers) + cols - 1) // cols

        self._tab2_inner.update_idletasks()
        self._tab2_canvas.configure(scrollregion=self._tab2_canvas.bbox("all"))

    def _select_all(self):
        for v in self.tab2_check_vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self.tab2_check_vars.values():
            v.set(False)

    def _select_group(self, tickers: list[str], value: bool):
        for t in tickers:
            if t in self.tab2_check_vars:
                self.tab2_check_vars[t].set(value)

    # =========================================================
    # Watchlist popup
    # =========================================================

    def _open_watchlist_popup(self):
        """Open watchlist manager. All edits are staged in _wl_draft and committed only on '儲存關閉'."""
        import copy
        self._ensure_groups(self.cfg)
        self._wl_draft = copy.deepcopy({
            "watchlist": self.cfg.get("watchlist", []),
            "groups":    self.cfg.get("groups",    []),
        })
        self._wl_group_collapsed = {}
        popup = tk.Toplevel(self.root)
        popup.title(t("gui.btn.manage_watchlist"))
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)
        popup.update()
        popup.attributes("-topmost", False)
        popup.bind("<Escape>", lambda e: popup.destroy())
        self._build_watchlist_popup(popup)

    def _build_watchlist_popup(self, popup: tk.Toplevel):
        """Build watchlist popup: scrollable list, add-company section, cache status, save/discard buttons."""
        pad = {"padx": 12, "pady": 4}
        popup.columnconfigure(0, weight=1)

        # ── Watchlist (scrollable, groups) ──────────────────────────
        list_frame = ttk.LabelFrame(popup, text=t("gui.frame.watchlist_current"), padding=6)
        list_frame.grid(row=0, column=0, sticky="ew", **pad)
        list_frame.columnconfigure(0, weight=1)

        # 圖示按鈕的說明寫在清單上方講一次，不要每個公司列都重複一遍（會很擠）——
        # TODO E13：📁／[x] 原本沒有任何文字說明，看不出是幹嘛用的
        ttk.Label(list_frame, text=t("gui.wl.icon_legend"),
                  foreground="#888888", font=("", 9)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        wl_canvas = tk.Canvas(list_frame, height=200, highlightthickness=0)
        wl_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=wl_canvas.yview)
        wl_canvas.configure(yscrollcommand=wl_scrollbar.set)
        wl_canvas.grid(row=1, column=0, sticky="ew")
        wl_scrollbar.grid(row=1, column=1, sticky="ns")
        wl_inner = ttk.Frame(wl_canvas)
        wl_win = wl_canvas.create_window((0, 0), window=wl_inner, anchor="nw")
        wl_inner.bind("<Configure>", lambda e: (
            wl_canvas.configure(scrollregion=wl_canvas.bbox("all")),
            wl_canvas.itemconfig(wl_win, width=wl_canvas.winfo_width()),
        ))
        wl_canvas.bind("<Configure>", lambda e: wl_canvas.itemconfig(wl_win, width=e.width))
        self._wl_list_canvas = wl_canvas
        self._wl_list_container = wl_inner
        self._refresh_wl_popup_list(wl_inner)

        ttk.Button(popup, text=t("gui.btn.add_group"),
                   command=lambda: self._wl_add_group(wl_inner)).grid(
            row=1, column=0, sticky="w", padx=12, pady=(0, 4))

        # ── Add company ────────────────────────────────────────────
        add_frame = ttk.LabelFrame(popup, text=t("gui.frame.add_company"), padding=6)
        add_frame.grid(row=2, column=0, sticky="ew", **pad)
        row_add = ttk.Frame(add_frame)
        row_add.grid(row=0, column=0, sticky="ew")
        ttk.Label(row_add, text="Ticker:").pack(side="left", padx=(0, 6))
        self.wl_add_var = tk.StringVar()
        wl_entry = ttk.Entry(row_add, textvariable=self.wl_add_var, width=10)
        wl_entry.pack(side="left", padx=(0, 8))
        wl_entry.bind("<Return>", lambda e: self._wl_lookup())
        ttk.Label(row_add, text=t("gui.lbl.group")).pack(side="left", padx=(8, 4))
        group_names = [_group_display(g["name"])
                       for g in self._get_groups_sorted(self._wl_draft)] or [_group_display(UNCATEGORIZED)]
        self.wl_group_var = tk.StringVar(value=group_names[0])
        self.wl_group_cb = ttk.Combobox(row_add, textvariable=self.wl_group_var,
                                         values=group_names, width=12, state="readonly")
        self.wl_group_cb.pack(side="left", padx=(0, 8))
        ttk.Button(row_add, text=t("gui.btn.lookup"), command=lambda: self._wl_lookup()).pack(side="left")
        self.wl_lookup_label = ttk.Label(add_frame, text="", foreground="gray")
        self.wl_lookup_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.wl_add_btn = ttk.Button(add_frame, text=t("gui.btn.add_to_watchlist"), command=self._wl_add, state="disabled")
        self.wl_add_btn.grid(row=2, column=0, sticky="w", pady=4)
        self._wl_found_name = ""

        # ── Cache status ───────────────────────────────────────────
        cache_frame = ttk.Frame(popup)
        cache_frame.grid(row=3, column=0, sticky="ew", **pad)
        self.wl_cache_label = ttk.Label(cache_frame, text=self._wl_cache_status(), foreground="#555555")
        self.wl_cache_label.pack(side="left")
        ttk.Button(cache_frame, text=t("gui.btn.update_name_cache"),
                   command=self._wl_update_cache).pack(side="left", padx=10)

        # ── Save / discard ─────────────────────────────────────────
        btn_row = ttk.Frame(popup)
        btn_row.grid(row=4, column=0, pady=8)
        ttk.Button(btn_row, text=t("gui.btn.save_close"), width=12,
                   command=lambda: self._wl_save_close(popup)).pack(side="left", padx=6)
        ttk.Button(btn_row, text=t("gui.btn.discard_close"), width=12,
                   command=popup.destroy).pack(side="left", padx=6)

    def _refresh_wl_popup_list(self, container):
        """Redraw the watchlist inside the popup from _wl_draft, with collapsible group headers."""
        for w in container.winfo_children():
            w.destroy()
        watchlist = self._wl_draft.get("watchlist", [])
        wl_map = {w["ticker"]: w for w in watchlist}
        groups = self._get_groups_sorted(self._wl_draft)

        if not watchlist and not any(g["tickers"] for g in groups):
            ttk.Label(container, text=t("gui.wl.empty"), foreground="gray").pack(anchor="w")
            container.update_idletasks()
            if hasattr(self, "_wl_list_canvas"):
                self._wl_list_canvas.configure(scrollregion=self._wl_list_canvas.bbox("all"))
            return

        for group in groups:
            gname = group["name"]
            tickers = sorted(t for t in group["tickers"] if t in wl_map)
            is_collapsed = self._wl_group_collapsed.get(gname, False)

            # Group header — 群組名在左，操作按鈕統一靠右（TODO E13，CTH 建議照
            # Windows 慣例）。原本重新命名／刪除緊貼在群組名右邊，群組名一長就會
            # 把按鈕往右推、視覺上像要溢出；靠右對齊後兩者互不影響對方的位置
            hdr = ttk.Frame(container)
            hdr.pack(fill="x", pady=(6, 0))
            arrow = "▶" if is_collapsed else "▼"
            ttk.Button(hdr, text=f"{arrow} {_group_display(gname)}", width=16,
                       command=lambda g=gname, c=container: self._wl_toggle_group(g, c)).pack(side="left")
            if gname != UNCATEGORIZED:
                ttk.Button(hdr, text=t("gui.btn.delete_group"), width=8,
                           command=lambda g=gname, c=container: self._wl_delete_group(g, c)).pack(side="right")
            ttk.Button(hdr, text=t("gui.btn.rename"), width=8,
                       command=lambda g=gname, c=container: self._wl_rename_group(g, c)).pack(side="right")

            if not is_collapsed:
                if not tickers:
                    ttk.Label(container, text=t("gui.wl.empty_group"), foreground="gray").pack(anchor="w", padx=(20, 0))
                for ticker in tickers:
                    item = wl_map[ticker]
                    row = ttk.Frame(container)
                    row.pack(fill="x", pady=1, padx=(20, 0))
                    ttk.Label(row, text=f'{ticker:6} {item.get("name", "")}', width=32).pack(side="left")
                    ttk.Button(row, text="📁", width=3,
                               command=lambda t=ticker, c=container: self._wl_set_output_dir(t, c)).pack(side="left", padx=(2, 0))
                    out_dir = item.get("output_dir", "")
                    if out_dir:
                        parts = Path(out_dir).parts
                        short = os.sep.join(parts[-2:]) if len(parts) >= 2 else out_dir
                        path_text = f"…{os.sep}{short}"
                        path_fg = "black"
                    else:
                        path_text = t("gui.wl.default_group")
                        path_fg = "gray"
                    ttk.Label(row, text=path_text, foreground=path_fg, width=18).pack(side="left", padx=(4, 2))
                    ttk.Button(row, text="[x]", width=4,
                               command=lambda t=ticker, c=container: self._wl_remove(t, c)).pack(side="left")

        container.update_idletasks()
        if hasattr(self, "_wl_list_canvas"):
            self._wl_list_canvas.configure(scrollregion=self._wl_list_canvas.bbox("all"))

    def _wl_lookup(self):
        """Start async company-name lookup for the ticker in the add-company input."""
        ticker = self.wl_add_var.get().strip().upper()
        if not ticker:
            self.wl_lookup_label.config(text=t("gui.msg.enter_ticker"), foreground="red")
            return
        self.wl_lookup_label.config(text=t("gui.status.looking_up"), foreground="gray")
        self.wl_add_btn.config(state="disabled")
        self._wl_found_name = ""
        threading.Thread(target=lambda: self._wl_lookup_worker(ticker), daemon=True).start()

    def _wl_lookup_worker(self, ticker: str):
        """Background thread: resolve company name for watchlist add (local cache first, then live EDGAR)."""
        cache: dict[str, str] = {}
        if CACHE_PATH.exists():
            try:
                with open(CACHE_PATH, encoding="utf-8") as f:
                    cache = json.load(f).get("companies", {})
            except (json.JSONDecodeError, OSError):
                cache = {}
        if ticker in cache:
            self.msg_queue.put(("wl_lookup_result", ("ok", ticker, cache[ticker])))
            return
        try:
            from edgar import Company, set_identity
            set_identity(self.cfg.get("identity", "SEC Tool sec@example.com"))
            c = Company(ticker)
            name = c.name or ticker
            self.msg_queue.put(("wl_lookup_result", ("ok", ticker, name)))
        except Exception as e:
            self.msg_queue.put(("wl_lookup_result", ("error", str(e))))

    def _wl_set_output_dir(self, ticker: str, container):
        from tkinter import filedialog
        folder = filedialog.askdirectory(title=t("gui.dlg.choose_ticker_dir", ticker=ticker))
        if not folder:
            return
        for item in self._wl_draft.get("watchlist", []):
            if item["ticker"] == ticker:
                item["output_dir"] = folder
                break
        self._refresh_wl_popup_list(container)

    def _wl_remove(self, ticker: str, container):
        """Remove ticker from both the watchlist array and all groups in _wl_draft."""
        self._wl_draft["watchlist"] = [w for w in self._wl_draft.get("watchlist", []) if w["ticker"] != ticker]
        for g in self._wl_draft.get("groups", []):
            if ticker in g["tickers"]:
                g["tickers"].remove(ticker)
        self._refresh_wl_popup_list(container)

    def _wl_add(self):
        """Add the looked-up ticker to _wl_draft watchlist and assign it to the selected group."""
        ticker = self.wl_add_var.get().strip().upper()
        if not ticker or not self._wl_found_name:
            return
        if any(w["ticker"] == ticker for w in self._wl_draft.get("watchlist", [])):
            self.wl_lookup_label.config(text=t("gui.wl.already_added", ticker=ticker), foreground="orange")
            return
        self._wl_draft.setdefault("watchlist", []).append({"ticker": ticker, "name": self._wl_found_name})
        target = (_group_stored(self.wl_group_var.get())
                  if self.wl_group_var else UNCATEGORIZED)
        grp = next((g for g in self._wl_draft.get("groups", []) if g["name"] == target), None)
        if grp is None:
            self._wl_draft.setdefault("groups", []).append({"name": target, "tickers": [ticker]})
        elif ticker not in grp["tickers"]:
            grp["tickers"].append(ticker)
        self.wl_add_var.set("")
        self.wl_lookup_label.config(text=t("gui.wl.added", ticker=ticker, group=_group_display(target)), foreground="#1a7a34")
        self.wl_add_btn.config(state="disabled")
        self._wl_found_name = ""
        self._refresh_wl_popup_list(self._wl_list_container)

    def _wl_toggle_group(self, group_name: str, container):
        self._wl_group_collapsed[group_name] = not self._wl_group_collapsed.get(group_name, False)
        self._refresh_wl_popup_list(container)

    def _wl_add_group(self, container):
        from tkinter import simpledialog
        name = simpledialog.askstring(t("gui.dlg.add_group_title"), t("gui.dlg.add_group_prompt"), parent=container.winfo_toplevel())
        if not name or not name.strip():
            return
        name = name.strip()
        if any(g["name"] == name for g in self._wl_draft.get("groups", [])):
            messagebox.showwarning(t("gui.dlg.duplicate_title"), t("gui.wl.group_exists", name=name), parent=container.winfo_toplevel())
            return
        # 過一次 _group_stored：英文介面下有人輸入 "Uncategorized" 的話，
        # 那就是預設群組本身，不該建出第二個同名（顯示上）的空群組
        self._wl_draft.setdefault("groups", []).append(
            {"name": _group_stored(name), "tickers": []})
        self._refresh_group_dropdown()
        self._refresh_wl_popup_list(container)

    def _wl_rename_group(self, old_name: str, container):
        from tkinter import simpledialog
        new_name = simpledialog.askstring(t("gui.btn.rename"), t("gui.wl.rename_prompt", old=_group_display(old_name)),
                                           parent=container.winfo_toplevel())
        if not new_name or not new_name.strip() or new_name.strip() == old_name:
            return
        new_name = new_name.strip()
        if any(g["name"] == new_name for g in self._wl_draft.get("groups", [])):
            messagebox.showwarning(t("gui.dlg.duplicate_title"), t("gui.wl.group_exists", name=new_name), parent=container.winfo_toplevel())
            return
        for g in self._wl_draft.get("groups", []):
            if g["name"] == old_name:
                g["name"] = new_name
                break
        if old_name in self._wl_group_collapsed:
            self._wl_group_collapsed[new_name] = self._wl_group_collapsed.pop(old_name)
        self._refresh_group_dropdown()
        self._refresh_wl_popup_list(container)

    def _wl_delete_group(self, group_name: str, container):
        """Delete a group, moving its tickers to 未分類 with a confirmation dialog if the group is non-empty."""
        grp = next((g for g in self._wl_draft.get("groups", []) if g["name"] == group_name), None)
        if not grp:
            return
        if grp["tickers"]:
            if not messagebox.askyesno(t("gui.dlg.confirm_delete_title"),
                                        t("gui.wl.delete_confirm", name=_group_display(group_name),
                                          n=len(grp["tickers"]),
                                          fallback=t("gui.wl.uncategorized")),
                                        parent=container.winfo_toplevel()):
                return
            uncategorized = next((g for g in self._wl_draft["groups"] if g["name"] == UNCATEGORIZED), None)
            if uncategorized is None:
                self._wl_draft["groups"].append({"name": UNCATEGORIZED, "tickers": list(grp["tickers"])})
            else:
                uncategorized["tickers"].extend(grp["tickers"])
        self._wl_draft["groups"] = [g for g in self._wl_draft["groups"] if g["name"] != group_name]
        self._wl_group_collapsed.pop(group_name, None)
        self._refresh_group_dropdown()
        self._refresh_wl_popup_list(container)

    def _refresh_group_dropdown(self):
        if not self.wl_group_cb:
            return
        try:
            names = [_group_display(g["name"])
                     for g in self._get_groups_sorted(self._wl_draft)] or [_group_display(UNCATEGORIZED)]
            self.wl_group_cb["values"] = names
            if self.wl_group_var.get() not in names:
                self.wl_group_var.set(names[0])
        except tk.TclError:
            pass

    def _wl_save_close(self, popup: tk.Toplevel):
        """Commit _wl_draft to live config, save to disk, refresh Tab 2, and close the popup."""
        self.cfg["watchlist"] = self._wl_draft.get("watchlist", [])
        self.cfg["groups"]    = self._wl_draft.get("groups",    [])
        save_config(self.cfg, CONFIG_PATH)
        self._refresh_tab2_list()
        popup.destroy()

    def _wl_update_cache(self):
        self.wl_cache_label.config(text=t("gui.status.updating"), foreground="gray")
        threading.Thread(target=self._wl_update_cache_worker, daemon=True).start()

    def _wl_update_cache_worker(self):
        """Background thread: download full SEC ticker list from EDGAR and save to company_cache.json."""
        try:
            identity = self.cfg.get("identity") or "SEC Tool sec@example.com"
            url = "https://www.sec.gov/files/company_tickers.json"
            req = urllib.request.Request(url, headers={"User-Agent": identity})
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            companies = {v["ticker"].upper(): v["title"] for v in raw.values()}
            cache_data = {"last_updated": str(date.today()), "companies": companies}
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            self.msg_queue.put(("wl_cache_updated", (str(date.today()), len(companies))))
        except Exception as e:
            self.msg_queue.put(("wl_cache_update_error", str(e)))

    def _ensure_groups(self, cfg: dict) -> None:
        """Migrate old watchlist (no groups key) to groups structure."""
        if "groups" not in cfg:
            tickers = [w["ticker"] for w in cfg.get("watchlist", [])]
            cfg["groups"] = [{"name": UNCATEGORIZED, "tickers": tickers}] if tickers else []
        else:
            all_grouped = {t for g in cfg["groups"] for t in g["tickers"]}
            ungrouped = [w["ticker"] for w in cfg.get("watchlist", []) if w["ticker"] not in all_grouped]
            if ungrouped:
                for g in cfg["groups"]:
                    if g["name"] == UNCATEGORIZED:
                        g["tickers"].extend(ungrouped)
                        break
                else:
                    cfg["groups"].append({"name": UNCATEGORIZED, "tickers": ungrouped})

    def _get_groups_sorted(self, cfg: dict) -> list[dict]:
        """Return groups sorted A-Z, 未分類 always last."""
        groups = cfg.get("groups", [])
        known = sorted([g for g in groups if g["name"] != UNCATEGORIZED], key=lambda g: g["name"])
        uncategorized = [g for g in groups if g["name"] == UNCATEGORIZED]
        return known + uncategorized

    def _wl_cache_status(self) -> str:
        if CACHE_PATH.exists():
            try:
                with open(CACHE_PATH, encoding="utf-8") as f:
                    data = json.load(f)
                count = len(data.get("companies", {}))
                return t("gui.wl.cache_loaded", count=f"{count:,}",
                                 updated=data.get("last_updated") or t("gui.wl.unknown_time"))
            except (json.JSONDecodeError, OSError):
                return t("gui.wl.cache_corrupt")
        return t("gui.wl.cache_absent")

    # =========================================================
    # Cross-company comparison tab
    # =========================================================

    def _load_company_cache(self) -> dict[str, str]:
        """讀 company_cache.json，回傳 {ticker: 公司名}。讀不到就回空字典，
        不要讓自動完成功能因為快取檔案缺失而整個掛掉。"""
        if not CACHE_PATH.exists():
            return {}
        try:
            with open(CACHE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            return data.get("companies", {})
        except (json.JSONDecodeError, OSError):
            return {}

    def _build_tab4(self):
        """Build Tab 4 (跨公司比較): 選擇視窗按鈕、輸出設定、執行按鈕。

        進度條／log 沿用 Tab1/Tab2 共用的根層級元件（`self.progress_bar`／
        `self.log_text`，透過 `self._log()`／`self._set_progress()` 更新），
        不另外做一份 Tab4 專用的，這樣跟現有分頁的行為一致。
        """
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=t("gui.tab.compare"))
        tab.columnconfigure(0, weight=1)

        self.compare_selected_tickers: list[tuple[str, str]] = []
        self.compare_selected_metrics: list[str] = []
        self.compare_start_year = tk.StringVar(value="")
        self.compare_end_year = tk.StringVar(value="")
        self.compare_frequency = tk.StringVar(value="quarterly")
        self.compare_snapshot_date = tk.StringVar(value="")

        summary_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=8)
        summary_frame.grid(row=0, column=0, sticky="ew", pady=4)
        summary_frame.columnconfigure(0, weight=1)
        self.compare_summary_label = ttk.Label(summary_frame, text=t("gui.compare.no_selection"),
                                                justify="left")
        self.compare_summary_label.grid(row=0, column=0, sticky="w")

        ttk.Button(tab, text=t("gui.btn.compare_select"),
                   command=self._open_compare_selection_window).grid(row=1, column=0, pady=8)

        out_row = ttk.Frame(tab)
        out_row.grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Label(out_row, text=t("gui.lbl.save_location")).pack(side="left")
        self.compare_outdir_var = tk.StringVar(value=str(PROJECT_ROOT / "output" / "compare"))
        ttk.Entry(out_row, textvariable=self.compare_outdir_var).pack(
            side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(out_row, text=t("gui.btn.browse"), width=5,
                   command=self._browse_compare_output_dir).pack(side="left")

        self.compare_run_btn = ttk.Button(tab, text=t("gui.btn.compare_run"),
                                           command=self._run_comparison)
        self.compare_run_btn.grid(row=3, column=0, pady=8)

    def _browse_compare_output_dir(self):
        from tkinter import filedialog
        current = self.compare_outdir_var.get().strip() or str(PROJECT_ROOT / "output" / "compare")
        folder = filedialog.askdirectory(title=t("gui.dlg.choose_output_dir"), initialdir=current)
        if folder:
            self.compare_outdir_var.set(folder)

    def _update_compare_summary(self):
        if not self.compare_selected_tickers or not self.compare_selected_metrics:
            self.compare_summary_label.config(text=t("gui.compare.no_selection"))
            return
        tickers_str = "、".join(tk_ for tk_, _ in self.compare_selected_tickers)
        metrics_str = "、".join(self.compare_selected_metrics[:5])
        if len(self.compare_selected_metrics) > 5:
            metrics_str += f" ...({len(self.compare_selected_metrics)})"
        freq_label = (t("gui.compare.freq_quarterly") if self.compare_frequency.get() == "quarterly"
                      else t("gui.compare.freq_annual"))
        text = (f"{t('gui.compare.companies')}: {tickers_str}\n"
                f"{t('gui.compare.period')}: {self.compare_start_year.get()}"
                f"~{self.compare_end_year.get()} ({freq_label})\n"
                f"{t('gui.compare.metrics')}: {metrics_str}")
        self.compare_summary_label.config(text=text)

    def _open_compare_selection_window(self):
        win = tk.Toplevel(self.root)
        win.title(t("gui.compare.select_title"))
        win.geometry("560x640")

        # ── ① 選公司 ──────────────────────────────────────────────
        ttk.Label(win, text=t("gui.compare.step1_company"), font=("", 11, "bold")).pack(
            anchor="w", padx=10, pady=(10, 2))

        ticker_row = ttk.Frame(win)
        ticker_row.pack(fill="x", padx=10)
        ttk.Label(ticker_row, text=t("gui.compare.ticker_input")).pack(side="left")
        ticker_var = tk.StringVar()
        ticker_entry = ttk.Entry(ticker_row, textvariable=ticker_var, width=30)
        ticker_entry.pack(side="left", padx=(6, 0), fill="x", expand=True)

        suggest_listbox = tk.Listbox(win, height=4)
        cache = self._load_company_cache()

        def _on_ticker_type(*_):
            text = ticker_var.get().strip().upper()
            suggest_listbox.delete(0, "end")
            if not text or "," in text:
                return
            matches = [(tk_, name) for tk_, name in cache.items() if tk_.startswith(text)][:8]
            for tk_, name in matches:
                suggest_listbox.insert("end", f"{tk_}  {name}")

        ticker_var.trace_add("write", _on_ticker_type)

        chips_frame = ttk.Frame(win)
        chips_frame.pack(fill="x", padx=10, pady=(4, 0))

        def _refresh_company_chips():
            for child in chips_frame.winfo_children():
                child.destroy()
            for tk_, name in self.compare_selected_tickers:
                chip = ttk.Frame(chips_frame, relief="raised", borderwidth=1)
                chip.pack(side="left", padx=2, pady=2)
                ttk.Label(chip, text=f"{tk_} {name}").pack(side="left", padx=(4, 0))
                ttk.Button(chip, text="✕", width=2,
                           command=lambda t_=tk_: _remove_company(t_)).pack(side="left")

        def _add_company(ticker: str):
            ticker = ticker.strip().upper()
            if not ticker or any(t_ == ticker for t_, _ in self.compare_selected_tickers):
                return
            name = cache.get(ticker, "")
            if not name:
                messagebox.showwarning(
                    t("gui.compare.unknown_ticker_title"),
                    t("gui.compare.unknown_ticker_msg").format(ticker=ticker))
                return
            self.compare_selected_tickers.append((ticker, name))
            _refresh_company_chips()

        def _remove_company(ticker: str):
            self.compare_selected_tickers = [
                (t_, n) for t_, n in self.compare_selected_tickers if t_ != ticker
            ]
            _refresh_company_chips()

        def _on_ticker_submit(event=None):
            text = ticker_var.get().strip()
            if "," in text:
                for part in text.split(","):
                    _add_company(part)
            else:
                _add_company(text)
            ticker_var.set("")
            suggest_listbox.delete(0, "end")

        ticker_entry.bind("<Return>", _on_ticker_submit)

        def _on_suggest_pick(event):
            selection = suggest_listbox.curselection()
            if not selection:
                return
            picked = suggest_listbox.get(selection[0]).split()[0]
            _add_company(picked)
            ticker_var.set("")
            suggest_listbox.delete(0, "end")

        suggest_listbox.bind("<<ListboxSelect>>", _on_suggest_pick)
        suggest_listbox.pack(fill="x", padx=10)

        _refresh_company_chips()

        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ── ② 選指標 ──────────────────────────────────────────────
        ttk.Label(win, text=t("gui.compare.step2_metrics"), font=("", 11, "bold")).pack(
            anchor="w", padx=10)

        period_row = ttk.Frame(win)
        period_row.pack(fill="x", padx=10, pady=4)
        ttk.Label(period_row, text=t("gui.compare.start_year")).pack(side="left")
        ttk.Entry(period_row, textvariable=self.compare_start_year, width=6).pack(
            side="left", padx=(2, 8))
        ttk.Label(period_row, text=t("gui.compare.end_year")).pack(side="left")
        ttk.Entry(period_row, textvariable=self.compare_end_year, width=6).pack(
            side="left", padx=(2, 8))
        ttk.Label(period_row, text=t("gui.compare.frequency")).pack(side="left")
        freq_combo = ttk.Combobox(period_row, textvariable=self.compare_frequency,
                                   values=["quarterly", "annual"], state="readonly", width=10)
        freq_combo.pack(side="left", padx=(2, 0))

        category_row = ttk.Frame(win)
        category_row.pack(fill="x", padx=10, pady=4)
        ttk.Label(category_row, text=t("gui.compare.metric_category")).pack(side="left")

        from ratios import RATIO_CATEGORIES, RATIO_DEFS
        # 內部分類鍵一律英文（IS/BS/CF 是既有的科目分類標籤，RATIO_CATEGORIES
        # 是 ratios.py 的比率分類），下拉選單顯示的文字另外透過 t() 翻譯，
        # 不要把中文寫死進分類鍵本身。
        category_keys = ["IS", "BS", "CF"] + RATIO_CATEGORIES
        category_labels = {
            key: t(f"gui.compare.cat_{key.lower().replace(' ', '_')}") for key in category_keys
        }
        category_var = tk.StringVar(value=category_labels[category_keys[0]])
        category_combo = ttk.Combobox(category_row, textvariable=category_var,
                                       values=list(category_labels.values()),
                                       state="readonly", width=16)
        category_combo.pack(side="left", padx=(4, 0))

        metric_check_frame = ttk.Frame(win)
        metric_check_frame.pack(fill="x", padx=10)
        metric_vars: dict[str, tk.BooleanVar] = {}

        def _selected_category_key() -> str:
            label = category_var.get()
            return next((k for k, v in category_labels.items() if v == label), category_keys[0])

        def _metrics_for_category(category_key: str) -> list[str]:
            if category_key in ("IS", "BS", "CF"):
                return self._raw_concepts_for_tag(category_key)
            return [name for name, _, cat, _ in RATIO_DEFS if cat == category_key]

        def _refresh_metric_checkboxes(*_):
            for child in metric_check_frame.winfo_children():
                child.destroy()
            names = _metrics_for_category(_selected_category_key())
            for idx, name in enumerate(names):
                var = metric_vars.setdefault(name, tk.BooleanVar(
                    value=name in self.compare_selected_metrics))

                def _on_toggle(name_=name, var_=var):
                    if var_.get() and name_ not in self.compare_selected_metrics:
                        self.compare_selected_metrics.append(name_)
                    elif not var_.get() and name_ in self.compare_selected_metrics:
                        self.compare_selected_metrics.remove(name_)
                    _refresh_metric_chips()

                ttk.Checkbutton(metric_check_frame, text=name, variable=var,
                                command=_on_toggle).grid(
                    row=idx // 2, column=idx % 2, sticky="w", padx=4)

        category_combo.bind("<<ComboboxSelected>>", _refresh_metric_checkboxes)

        metric_chips_frame = ttk.Frame(win)
        metric_chips_frame.pack(fill="x", padx=10, pady=(4, 0))

        def _refresh_metric_chips():
            for child in metric_chips_frame.winfo_children():
                child.destroy()
            for name in self.compare_selected_metrics:
                chip = ttk.Frame(metric_chips_frame, relief="raised", borderwidth=1)
                chip.pack(side="left", padx=2, pady=2)
                ttk.Label(chip, text=name).pack(side="left", padx=(4, 0))

                def _remove(name_=name):
                    self.compare_selected_metrics.remove(name_)
                    if name_ in metric_vars:
                        metric_vars[name_].set(False)
                    _refresh_metric_chips()

                ttk.Button(chip, text="✕", width=2, command=_remove).pack(side="left")

        _refresh_metric_checkboxes()
        _refresh_metric_chips()

        snapshot_row = ttk.Frame(win)
        snapshot_row.pack(fill="x", padx=10, pady=6)
        ttk.Label(snapshot_row, text=t("gui.compare.snapshot_date")).pack(side="left")
        ttk.Entry(snapshot_row, textvariable=self.compare_snapshot_date, width=14).pack(
            side="left", padx=(4, 0))

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_row, text=t("gui.btn.cancel"), command=win.destroy).pack(
            side="right", padx=4)

        def _confirm():
            self._update_compare_summary()
            win.destroy()

        ttk.Button(btn_row, text=t("gui.btn.confirm"), command=_confirm).pack(
            side="right", padx=4)

    def _raw_concepts_for_tag(self, tag: str) -> list[str]:
        """從 fetcher_gaap 的科目定義表取某一類（IS/BS/CF）的欄位名稱清單。

        `IS_TEMPLATE`／`BS_TEMPLATE`／`CF_TEMPLATE` 是 fetcher_gaap.py 裡已經
        依報表類型分開的 module-level 清單，每筆 tuple 的第 0 欄就是顯示名稱。
        """
        from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
        source = {"IS": IS_TEMPLATE, "BS": BS_TEMPLATE, "CF": CF_TEMPLATE}[tag]
        return [row[0] for row in source]

    def _run_comparison(self):
        if not self.compare_selected_tickers:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.compare.no_company_warn"))
            return
        if not self.compare_selected_metrics:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.compare.no_metric_warn"))
            return
        identity = self.cfg.get("identity", "")
        if not identity:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.lbl.identity_missing"))
            return

        self.compare_run_btn.config(state="disabled")
        threading.Thread(target=self._compare_worker, daemon=True).start()

    def _compare_worker(self):
        from comparison import build_comparison
        from comparison_writer import write_comparison_workbook

        identity = self.cfg.get("identity", "")
        tickers = [t_ for t_, _ in self.compare_selected_tickers]
        metrics = list(self.compare_selected_metrics)
        start_year = int(self.compare_start_year.get()) if self.compare_start_year.get().strip() else None
        end_year = int(self.compare_end_year.get()) if self.compare_end_year.get().strip() else None
        frequency = self.compare_frequency.get()

        self._log(t("gui.compare.log_start", n=len(tickers)))
        self._set_progress(0, len(tickers), t("gui.compare.log_start", n=len(tickers)))
        try:
            result = build_comparison(
                tickers, identity, metrics, frequency=frequency,
                start_year=start_year, end_year=end_year,
            )
        except Exception as e:
            self.msg_queue.put(("compare_error", f"{type(e).__name__}{_exc_status(e)}"))
            return

        for failure in result.failures:
            self._log(t("gui.compare.log_company_failed",
                         ticker=failure.ticker, error_type=failure.error_type), "WARN")

        if not any(result.metrics.get(m) for m in metrics):
            self.msg_queue.put(("compare_error", t("gui.compare.nothing_fetched")))
            return

        out_dir = Path(self.compare_outdir_var.get().strip() or str(PROJECT_ROOT / "output" / "compare"))
        names = "_".join(tickers[:3])
        filename = f"比較_{names}_{date.today().strftime('%Y%m%d')}.xlsx"
        out_path = out_dir / filename

        try:
            write_comparison_workbook(
                result, metrics, out_path, snapshot_date=self.compare_snapshot_date.get().strip()
            )
        except Exception as e:
            self.msg_queue.put(("compare_error", f"{type(e).__name__}{_exc_status(e)}"))
            return

        self.msg_queue.put(("compare_done", str(out_path)))

    # =========================================================
    # Advanced settings tab
    # =========================================================

    def _build_tab3(self):
        """Build Tab 3 (進階設定). 2026-08-17 CTH 要求從彈出視窗改成頁籤。

        內容放進固定高度可捲動容器：Notebook 的高度取所有頁籤中最高的那個，
        設定內容比另外兩頁高不少，直接塞進去會把整個 Notebook 撐高，而下面
        「處理進度」的 log 是唯一 weight=1 的列，撐出來的高度全由它吸收。
        限高之後多的部分用捲的，另外兩頁的版面完全不受影響。

        存檔/還原按鈕（`_build_settings_footer`）刻意擺在可捲動容器**外面**，
        不放進 `popup`——這樣它們永遠貼在頁籤底部可見，不用捲到底才找得到
        （TODO E11）。`tab` 本身用 pack，容器是固定高度不會被撐大，footer
        接著 pack 上去自然落在容器下方。
        """
        tab = ttk.Frame(self.notebook, padding=(4, 6))
        self.notebook.add(tab, text=t("gui.tab.settings"))
        _, inner = _build_fixed_height_scrollable(tab, height=self._TAB3_HEIGHT)
        self._build_settings_panel(inner)
        self._build_settings_footer(tab)

    # 實測貼齊值（2026-08-18 隨 E16 一起重量）：Tab1 現在的實際內容需要
    # ~414px（比 2026-08-17 設計時的 343px 高，這段期間陸續加了 SEC Identity
    # 警告、輸出設定「設為預設」說明等新行）。存檔／還原按鈕搬到可捲動容器
    # 外面（`_build_settings_footer`）之後，容器本身不用再留空間給按鈕列，
    # 355 量出來的 Tab3 總高度 (~412px) 跟 Tab1 的 414px 只差 2px，兩頁看起來
    # 一樣高。改動任何一頁的版面之後要重量（寫一個 Tk 探針腳本：建
    # `SECFetcherApp`、`update_idletasks()`、讀各分頁 `winfo_reqheight()`）。
    _TAB3_HEIGHT = 355

    def _build_settings_footer(self, tab):
        """存檔／還原固定在頁籤最下面（TODO E11）——原本存檔鍵在可捲動內容的
        最後一列，要捲到底才看得到；搬出來後永遠可見，不用捲。

        「還原」不是「取消並關閉」：Tab3 是頁籤沒有可以關掉的視窗，取消鍵按了
        什麼都不會發生反而更混淆。「還原」把畫面上的欄位值改回上次 `save_config`
        存的值（`self.cfg`），讓使用者可以反悔本次還沒存檔的編輯。
        """
        footer = ttk.Frame(tab)
        footer.pack(fill="x", pady=(4, 4))
        ttk.Button(footer, text=t("gui.btn.save"),
                   command=self._save_settings, width=10).pack(side="left", padx=6)
        ttk.Button(footer, text=t("gui.btn.restore"),
                   command=self._restore_settings, width=10).pack(side="left")
        self.settings_saved_label = ttk.Label(footer, text="", foreground="#1a7a34")
        self.settings_saved_label.pack(side="left", padx=8)

    def _build_settings_panel(self, popup):
        """SEC identity、AI 設定、抓取上限、模板模式。

        參數名還叫 popup 是為了讓底下上百行的 .grid(in_=popup) 不用全改；
        它現在是頁籤裡的可捲動容器，不是 Toplevel。
        """
        pad = {"padx": 12, "pady": 4}
        # 同上：欄位要有 weight 視窗變寬才會撐開（TODO E16）
        popup.columnconfigure(0, weight=1)

        # Language — 最上方獨立一列。
        # 不另開 tab：主視窗高度鎖死 650px（見 __init__ 的 geometry 註解），
        # log 顯示區已緊到剩個位數 px，多一列會把 log 擠到全隱形。語言屬
        # 「設一次不再動」，與 Identity / API Key 同性質。
        #
        # 標籤固定英文、不翻譯（CTH 指定）；選項用各語言自稱，任何語言下
        # 都認得出哪個是哪個。選項由 i18n.LANGUAGES 動態生成——新增語言時
        # 這裡一個字都不必改。
        lang_frame = ttk.Frame(popup)
        lang_frame.grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(lang_frame, text="Language:").pack(side="left", padx=(0, 8))
        self._lang_choices = i18n.available_languages()
        # 讀 config 不讀 i18n.get_lang()：set_lang() 只在 __init__ 跑一次，
        # 使用者選了新語言但按 Later 不重啟時，runtime 語言還是舊的。
        # 用 runtime 值當基準，下次開設定按儲存會把選擇默默寫回去。
        self._lang_saved_code = (self.cfg.get("language")
                                 if i18n.is_supported(self.cfg.get("language", ""))
                                 else i18n.DEFAULT_LANG)
        names = [name for _, name in self._lang_choices]
        current_name = next((n for c, n in self._lang_choices if c == self._lang_saved_code),
                            names[0])
        self.settings_lang_var = tk.StringVar(value=current_name)
        ttk.Combobox(lang_frame, textvariable=self.settings_lang_var,
                     values=names, width=14, state="readonly").pack(side="left")

        # SEC Identity
        id_frame = ttk.LabelFrame(popup, text=" SEC EDGAR Identity ", padding=8)
        id_frame.grid(row=1, column=0, sticky="ew", **pad)
        id_frame.columnconfigure(1, weight=1)
        _identity_why_label = ttk.Label(id_frame, text=t("gui.lbl.identity_why"),
                  foreground="#555555", font=("", 10), justify="left")
        _identity_why_label.grid(row=0, column=0, columnspan=2, sticky="w")
        # wraplength 不能寫死一個像素數——視窗寬度是算比例來的（見 __init__），
        # 寫死的話小螢幕會被截、大螢幕又留一堆早換行的空白（TODO E16 的根因之
        # 一）。改成跟著 id_frame 實際寬度即時調整
        id_frame.bind("<Configure>",
                       lambda e: _identity_why_label.config(wraplength=max(200, e.width - 20)))
        ttk.Label(id_frame, text=t("gui.lbl.identity_hint"),
                  foreground="#555555", font=("", 10)).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Label(id_frame, text="Identity:").grid(row=2, column=0, sticky="w", pady=4)
        self.settings_identity_var = tk.StringVar(value=self.cfg.get("identity", ""))
        ttk.Entry(id_frame, textvariable=self.settings_identity_var, width=42).grid(row=2, column=1, sticky="ew", padx=(8, 0))

        # AI Config
        ai_frame = ttk.LabelFrame(popup, text=t("gui.frame.ai_settings"), padding=8)
        ai_frame.grid(row=2, column=0, sticky="ew", **pad)

        ttk.Label(ai_frame, text="Provider:").grid(row=0, column=0, sticky="w")
        self.settings_provider_var = tk.StringVar(value=self.cfg["ai"].get("provider", "google"))
        provider_cb = ttk.Combobox(ai_frame, textvariable=self.settings_provider_var,
                                   values=["google", "openai", "anthropic"], width=14, state="readonly")
        provider_cb.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=4)
        provider_cb.bind("<<ComboboxSelected>>", self._on_provider_change)

        ttk.Label(ai_frame, text="Model:").grid(row=1, column=0, sticky="w")
        self.settings_model_var = tk.StringVar(value=self.cfg["ai"].get("model", "gemini-flash-latest"))
        ttk.Entry(ai_frame, textvariable=self.settings_model_var, width=30).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(ai_frame, text="API Key:").grid(row=2, column=0, sticky="w")
        key_row = ttk.Frame(ai_frame)
        key_row.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=4)
        self.settings_key_var = tk.StringVar(value=self.cfg["ai"].get("api_key", ""))
        self.settings_key_entry = ttk.Entry(key_row, textvariable=self.settings_key_var, width=28, show="•")
        self.settings_key_entry.pack(side="left", padx=(0, 8))
        self.settings_key_toggle_btn = ttk.Button(key_row, text=t("gui.btn.show"), width=5, command=self._toggle_key_show)
        self.settings_key_toggle_btn.pack(side="left")
        tk.Label(ai_frame, text=t("gui.lbl.api_key_notice"),
                 foreground="#555555", font=("", 10)).grid(row=3, column=0, columnspan=2, sticky="w")

        test_row = ttk.Frame(ai_frame)
        test_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(test_row, text=t("gui.btn.test_connection"), command=self._test_ai_connection).pack(side="left")
        self.settings_test_label = ttk.Label(test_row, text="", foreground="gray")
        self.settings_test_label.pack(side="left", padx=10)

        # Fetch settings frame
        fetch_frame = ttk.LabelFrame(popup, text=t("gui.frame.fetch_settings"), padding=8)
        fetch_frame.grid(row=3, column=0, sticky="ew", **pad)
        fetch_frame.columnconfigure(2, weight=1)

        ttk.Label(fetch_frame, text=t("gui.lbl.max_filings")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.settings_max_filings_var = tk.IntVar(value=self.cfg.get("max_filings", 80))
        max_spin = ttk.Spinbox(fetch_frame, from_=4, to=320, increment=4,
                               textvariable=self.settings_max_filings_var, width=6)
        max_spin.grid(row=0, column=1, sticky="w")
        ttk.Label(fetch_frame, text=t("gui.lbl.max_filings_hint"), foreground="#555555").grid(
            row=0, column=2, sticky="w", padx=(4, 0))

        # Template mode
        ttk.Label(fetch_frame, text=t("gui.lbl.template")).grid(row=1, column=0, sticky="nw", pady=(10, 0))
        has_tpl = bool(self.cfg.get("template_path", ""))
        self.settings_template_mode_var = tk.StringVar(value="custom" if has_tpl else "default")
        self.settings_template_var = tk.StringVar(value=self.cfg.get("template_path", ""))

        tpl_frame = ttk.Frame(fetch_frame)
        tpl_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Radiobutton(tpl_frame, text=t("gui.radio.template_default"),
                        variable=self.settings_template_mode_var, value="default",
                        command=self._on_template_mode_change).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Radiobutton(tpl_frame, text=t("gui.radio.template_custom"),
                        variable=self.settings_template_mode_var, value="custom",
                        command=self._on_template_mode_change).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._tpl_entry = ttk.Entry(tpl_frame, textvariable=self.settings_template_var, width=24)
        self._tpl_entry.grid(row=1, column=1, sticky="ew", padx=(4, 4), pady=(4, 0))
        self._tpl_browse_btn = ttk.Button(tpl_frame, text=t("gui.btn.browse"), width=5,
                                           command=self._browse_template)
        self._tpl_browse_btn.grid(row=1, column=2, pady=(4, 0))
        self._on_template_mode_change()  # set initial enabled/disabled state
        # 存檔／還原按鈕搬到 `_build_settings_footer`（`tab` 層級，固定在頁籤
        # 最下面，不放在這個可捲動的 `popup` 裡），見 `_build_tab3` 的說明

    def _on_template_mode_change(self):
        is_custom = getattr(self, "settings_template_mode_var", None) and \
                    self.settings_template_mode_var.get() == "custom"
        state = "normal" if is_custom else "disabled"
        if hasattr(self, "_tpl_entry"):
            self._tpl_entry.config(state=state)
        if hasattr(self, "_tpl_browse_btn"):
            self._tpl_browse_btn.config(state=state)

    def _browse_template(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title=t("gui.dlg.choose_template"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path and hasattr(self, "settings_template_var"):
            self.settings_template_var.set(path)

    def _on_provider_change(self, _event=None):
        """Auto-fill the default model name when the AI provider selection changes."""
        provider = self.settings_provider_var.get()
        self.settings_model_var.set(PROVIDER_DEFAULTS.get(provider, ""))

    def _toggle_key_show(self):
        current = self.settings_key_entry.cget("show")
        new_show = "" if current else "•"
        self.settings_key_entry.config(show=new_show)
        if self.settings_key_toggle_btn:
            self.settings_key_toggle_btn.config(text=t("gui.btn.hide") if new_show == "" else t("gui.btn.show"))

    def _test_ai_connection(self):
        provider = self.settings_provider_var.get()
        model    = self.settings_model_var.get().strip()
        api_key  = self.settings_key_var.get().strip()
        if not api_key:
            self.settings_test_label.config(text=t("gui.msg.enter_api_key"), foreground="red")
            return
        self.settings_test_label.config(text=t("gui.status.testing"), foreground="gray")
        threading.Thread(
            target=lambda: self._test_ai_worker(provider, model, api_key), daemon=True
        ).start()

    def _test_ai_worker(self, provider: str, model: str, api_key: str):
        """Background thread: send a trivial prompt to verify AI API connectivity."""
        try:
            if provider == "google":
                import google.generativeai as genai
                genai.configure(api_key=api_key)
                m = genai.GenerativeModel(model)
                m.generate_content("Reply with one word: OK")
            elif provider == "openai":
                from openai import OpenAI
                OpenAI(api_key=api_key).chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": "Reply with one word: OK"}],
                    max_tokens=5,
                )
            elif provider == "anthropic":
                import anthropic
                anthropic.Anthropic(api_key=api_key).messages.create(
                    model=model, max_tokens=5,
                    messages=[{"role": "user", "content": "Reply with one word: OK"}],
                )
            self.msg_queue.put(("ai_test_result", ("ok", None)))
        except Exception as e:
            self.msg_queue.put(("ai_test_result", ("error", str(e))))

    def _save_settings(self, popup: tk.Toplevel | None = None):
        """存下設定頁的內容。popup 保留給還在用彈出視窗的呼叫端（目前沒有）。"""
        self.cfg["identity"]       = self.settings_identity_var.get().strip()
        self.cfg["ai"]["provider"] = self.settings_provider_var.get()
        self.cfg["ai"]["model"]    = self.settings_model_var.get().strip()
        self.cfg["ai"]["api_key"]  = self.settings_key_var.get().strip()
        try:
            self.cfg["max_filings"] = int(self.settings_max_filings_var.get())
        except (ValueError, tk.TclError):
            self.cfg["max_filings"] = 80
        if hasattr(self, "settings_template_mode_var"):
            if self.settings_template_mode_var.get() == "custom":
                self.cfg["template_path"] = self.settings_template_var.get().strip()
            else:
                self.cfg["template_path"] = ""
        new_lang = self._selected_lang_code()
        lang_changed = new_lang != self._lang_saved_code
        self.cfg["language"] = new_lang
        # 存檔後要更新這個基準值，不然「還原」（`_restore_settings`）會把語言
        # 還原回啟動當下的舊值，而不是這次剛存的新值——兩者現在都代表
        # 「上次已存檔的值」，必須同步
        self._lang_saved_code = new_lang
        save_config(self.cfg, CONFIG_PATH)
        self._update_identity_warnings()
        if popup is not None:
            popup.destroy()
        if getattr(self, "settings_saved_label", None) is not None:
            self.settings_saved_label.config(text=t("gui.status.settings_saved"))
            # 3 秒後自己消失。留著不動的話下次改完設定會分不出這句是這次的
            # 還是上次殘留的。
            self.root.after(3000, lambda: self.settings_saved_label.config(text=""))
        # 只有語言真的變更才打擾使用者——改 API Key 不該跳重啟視窗
        if lang_changed:
            self._prompt_restart_for_language()

    def _restore_settings(self):
        """「還原」按鈕（TODO E11）：把畫面上的欄位值改回 `self.cfg` 目前存的值

        `self.cfg` 只在 `_save_settings` 呼叫 `save_config` 時才會被改，所以
        它天生就是「上次已存檔的值」，不用另外記一份備份——把每個
        `settings_*_var` 重新從 `self.cfg` 讀一次、`.set()` 回填即可，等於把
        `_build_settings_panel` 裡讀初始值那段邏輯再跑一次。逐一列出要還原
        的變數，漏一個就是還原不完整：
          language / identity / ai(provider, model, api_key) / max_filings /
          template(mode, path)
        """
        self.settings_identity_var.set(self.cfg.get("identity", ""))
        self.settings_provider_var.set(self.cfg["ai"].get("provider", "google"))
        self.settings_model_var.set(self.cfg["ai"].get("model", "gemini-flash-latest"))
        self.settings_key_var.set(self.cfg["ai"].get("api_key", ""))
        self.settings_max_filings_var.set(self.cfg.get("max_filings", 80))

        has_tpl = bool(self.cfg.get("template_path", ""))
        self.settings_template_mode_var.set("custom" if has_tpl else "default")
        self.settings_template_var.set(self.cfg.get("template_path", ""))
        self._on_template_mode_change()  # 重新套用還原後的 mode 到 Entry/Browse 的 enable 狀態

        current_name = next((n for c, n in self._lang_choices if c == self._lang_saved_code),
                            self._lang_choices[0][1])
        self.settings_lang_var.set(current_name)

        if getattr(self, "settings_saved_label", None) is not None:
            self.settings_saved_label.config(text=t("gui.status.settings_restored"))
            self.root.after(3000, lambda: self.settings_saved_label.config(text=""))

    def _selected_lang_code(self) -> str:
        """把下拉選單顯示的語言名稱換回代號。取不到就維持原設定，不亂改。"""
        chosen = self.settings_lang_var.get()
        for code, name in self._lang_choices:
            if name == chosen:
                return code
        return self._lang_saved_code

    def _prompt_restart_for_language(self):
        """語言變更後問是否重啟。

        視窗全英文：此刻介面還是舊語言、使用者要的是新語言，用任一方都尷尬，
        英文最中立。
        """
        if messagebox.askyesno(
            "Language Changed",
            "Restart the app to apply the new language.\n\nRestart now?",
        ):
            self._restart_app()

    def _restart_app(self):
        """起一個新行程再關掉自己。

        不用 os.execv：Windows 上它會就地覆寫當前行程，tkinter 還沒釋放的
        視窗 handle 可能殘留，看起來像關不掉的殭屍視窗。Popen + destroy
        讓 tkinter 走完自己的關閉流程。

        ⚠ 已知的小瑕疵：launcher.ps1 是同步等 `python src\\main.py` 結束的，
        我們一 destroy 它就往下跑到收尾段，其中 `Remove-Item __pycache__`
        可能因為新行程正在用那些 .pyc 而在主控台印一行紅字。新視窗本身正常，
        exit code 也是 0（不會誤記 ERROR log）。要根治得改 launcher，
        不值得為一個選用的重啟路徑動啟動流程。
        """
        try:
            subprocess.Popen([sys.executable, *sys.argv], close_fds=True)
        except OSError:
            # 起不了新行程就什麼都不做——使用者下次自己開一樣會生效，
            # 這裡把舊視窗關掉反而讓人以為程式壞了
            return
        self.root.destroy()

    # =========================================================
    # Output path helpers
    # =========================================================

    def _lookup_company_name(self, ticker: str) -> str:
        """Look up company name: watchlist → cache → fallback to ticker."""
        for item in self.cfg.get("watchlist", []):
            if item["ticker"] == ticker:
                name = item.get("name", "")
                if name:
                    return name
        if CACHE_PATH.exists():
            try:
                with open(CACHE_PATH, encoding="utf-8") as f:
                    cache = json.load(f).get("companies", {})
                if ticker in cache:
                    return cache[ticker]
            except (json.JSONDecodeError, OSError):
                pass
        return ticker

    def _build_output_path(self, ticker: str) -> Path:
        """Build output file path. Priority: watchlist item output_dir → ticker_paths → global output_dir."""
        # 1. watchlist item output_dir
        for item in self.cfg.get("watchlist", []):
            if item["ticker"] == ticker and item.get("output_dir"):
                output_dir = Path(item["output_dir"])
                break
        else:
            # 2. legacy ticker_paths
            ticker_dir = self.cfg.get("ticker_paths", {}).get(ticker)
            if ticker_dir:
                output_dir = Path(ticker_dir)
            else:
                # 3. global output_dir
                output_dir = PROJECT_ROOT / self.cfg.get("output_dir", "output")

        fmt = self.cfg.get("filename_format", "ticker_name")
        if fmt == "ticker_name":
            name = self._lookup_company_name(ticker)
            safe_name = re.sub(r'[\\/:*?"<>|]', "", name).strip()
            filename = f"{ticker} {safe_name} data.xlsx"
        elif fmt == "custom":
            custom = re.sub(r'[\\/:*?"<>|]', "", self.cfg.get("filename_custom", "")).strip()
            filename = f"{custom}.xlsx" if custom else f"{ticker}.xlsx"
        else:
            filename = f"{ticker}.xlsx"
        return output_dir / filename

    # =========================================================
    # Open output folder
    # =========================================================

    def _open_output_folder(self):
        folder = self._last_output_folder or PROJECT_ROOT / self.cfg.get("output_dir", "output")
        if folder.exists():
            os.startfile(str(folder))

    # =========================================================
    # Run actions
    # =========================================================

    def _confirm_overwrite(self, message: str) -> bool:
        """覆蓋前的確認視窗。回 True 代表繼續覆蓋，False 代表取消。

        必須在主執行緒呼叫（`_run_single` / `_run_batch` 裡，還沒開 worker 前）
        ——tkinter 不是 thread-safe，在背景執行緒開視窗會當掉。
        """
        return messagebox.askokcancel(
            t("gui.dlg.overwrite_title"), message, parent=self.root, icon="warning"
        )

    def _run_single(self):
        """Validate inputs then launch the single-ticker fetch+write worker in a background thread."""
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.enter_ticker"))
            return
        if not self.fetch_gaap_var.get() and not self.fetch_nongaap_var.get():
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.pick_gaap_or_nongaap"))
            return
        fetch_gaap    = self.fetch_gaap_var.get()
        fetch_nongaap = self.fetch_nongaap_var.get()
        if fetch_nongaap and not self.cfg["ai"].get("api_key"):
            messagebox.showwarning(
                t("gui.dlg.need_key_title"),
                t("gui.msg.nongaap_need_key_body")
            )
            return
        max_filings = self.cfg.get("max_filings", 80)
        fetch_q = self.tab1_fetch_q_var.get() if self.tab1_fetch_q_var else True
        fetch_k = self.tab1_fetch_k_var.get() if self.tab1_fetch_k_var else True
        if not fetch_q and not fetch_k:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.pick_q_or_k"))
            return
        try:
            start_year = int(self.tab1_start_year_var.get()) if self.tab1_start_year_var and self.tab1_start_year_var.get().strip() else None
            end_year   = int(self.tab1_end_year_var.get())   if self.tab1_end_year_var   and self.tab1_end_year_var.get().strip()   else None
        except ValueError:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.bad_year"))
            return
        if start_year is not None and end_year is not None and start_year > end_year:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.year_range_reversed", start=start_year, end=end_year))
            return
        excluded = {
            name for name, var in self._sheet_check_vars.items()
            if not var.get() and name not in self._FIXED_SHEETS
        }
        # 覆蓋確認要在這裡問，不能放進 worker——worker 在背景執行緒，
        # tkinter 不是 thread-safe，在那裡開視窗會當掉。
        out_path = self._build_output_path(ticker)
        if out_path.exists():
            if not self._confirm_overwrite(t("gui.msg.overwrite_single", name=out_path.name)):
                return
        self._start_worker(lambda: self._worker_single(
            ticker, fetch_gaap, fetch_nongaap, max_filings,
            fetch_q=fetch_q, fetch_k=fetch_k,
            start_year=start_year, end_year=end_year,
            excluded_sheets=excluded,
        ))

    def _run_batch(self):
        selected = [t for t, v in self.tab2_check_vars.items() if v.get()]
        if not selected:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.pick_a_company"))
            return
        fetch_nongaap = self.batch_nongaap_var.get()
        fetch_q = self.batch_fetch_q_var.get() if self.batch_fetch_q_var else True
        fetch_k = self.batch_fetch_k_var.get() if self.batch_fetch_k_var else True
        if not fetch_q and not fetch_k:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.pick_q_or_k"))
            return
        try:
            start_year = int(self.batch_start_year_var.get()) if self.batch_start_year_var and self.batch_start_year_var.get().strip() else None
            end_year   = int(self.batch_end_year_var.get())   if self.batch_end_year_var   and self.batch_end_year_var.get().strip()   else None
        except ValueError:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.bad_year"))
            return
        if start_year is not None and end_year is not None and start_year > end_year:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.year_range_reversed", start=start_year, end=end_year))
            return
        if fetch_nongaap and not self.cfg["ai"].get("api_key"):
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.batch_nongaap_need_key"))
            return
        # 批量一次問完，不逐檔跳視窗——12 檔要點 12 次沒人受得了。
        hits = existing_outputs([self._build_output_path(s) for s in selected])
        if hits:
            _SHOWN = 10          # 太長的清單塞不進對話框，其餘用數字帶過
            names = "\n".join(f"　• {p.name}" for p in hits[:_SHOWN])
            if len(hits) > _SHOWN:
                names += f"\n　… +{len(hits) - _SHOWN}"
            msg = t("gui.msg.overwrite_batch", total=len(selected),
                    n=len(hits), names=names)
            if not self._confirm_overwrite(msg):
                return
        self._start_worker(lambda: self._worker_batch(
            selected, fetch_nongaap,
            fetch_q=fetch_q, fetch_k=fetch_k,
            start_year=start_year, end_year=end_year,
        ))

    def _show_scan_help(self):
        """Explain what 快速掃描 does — triggered by the '？' label next to the button."""
        win = tk.Toplevel(self.root)
        win.title(t("gui.help.scan_title"))
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text=t("gui.help.scan_heading"), font=("Microsoft JhengHei", 12, "bold")).pack(
            anchor="w", padx=20, pady=(16, 4)
        )
        lines = [
            t("gui.help.scan_l1"),
            t("gui.help.scan_l2"),
            "",
            t("gui.help.scan_l3"),
            t("gui.help.scan_l4"),
            t("gui.help.scan_l5"),
            "",
            t("gui.help.scan_l6"),
            "",
            t("gui.help.scan_l7"),
            t("gui.help.scan_l8"),
        ]
        for line in lines:
            ttk.Label(win, text=line, justify="left").pack(anchor="w", padx=20, pady=1)

        ttk.Button(win, text=t("gui.btn.close"), command=win.destroy).pack(pady=(12, 16))

    def _run_preview_scan(self):
        """Start background preview scan for the current ticker."""
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.enter_ticker_first"))
            return
        identity = self.cfg.get("identity", "")
        if not identity:
            messagebox.showerror(t("gui.dlg.error_title"), t("gui.msg.set_identity"))
            return
        if self._scan_btn:
            self._scan_btn.config(state="disabled", text=t("gui.status.scanning"))
        if self._scan_hint_label:
            self._scan_hint_label.config(text=t("gui.status.scan_hint"))
        # 名稱查詢由這裡自己發動，不再只依賴 ticker 欄位的 <FocusOut>。
        # 使用者若是用 Enter 確認過公司、焦點早就不在欄位上，點掃描不會再觸發
        # focusout，名稱那格就停在舊值；反過來若焦點還在欄位上，focusout 與這裡
        # 會各發一次查詢——同一個 ticker 走本機快取，重複一次的成本可忽略，
        # 換到的是「點一次必定兩件事都做」這個確定性。
        self._confirm_company()
        if self._sheet_panel_frame:
            self._sheet_panel_frame.grid_remove()
            self._sheet_panel_frame.configure(text=self._SHEET_PANEL_TITLE_BASE)
        self._sheet_check_vars = {}
        threading.Thread(
            target=lambda: self._preview_scan_worker(ticker, identity), daemon=True
        ).start()

    def _preview_scan_worker(self, ticker: str, identity: str):
        """Background thread: call preview_sheets() and push result to queue.

        2026-08-18 補上耗時 log（TODO E3）：CTH 回報「第一次按只顯示代號，第二
        次才查到期間」，但實際模擬點擊測試（含真的打 EDGAR）第一次就正常顯示，
        根因懷疑是 CTH 那台機器上第一次連線特別慢（防火牆/代理偵測之類，這邊
        環境重現不了）。原本這段完全沒寫 log，出事只能用猜的。現在起訖都記一
        行，耗時異常長的話 `logs/app.log` 就看得出來，不用再靠使用者口述秒數。
        """
        _write_log_header(f"查可用期間 {ticker}")
        t_start = time.time()
        try:
            from fetcher_gaap import preview_sheets
            result = preview_sheets(ticker, identity)
            elapsed = time.time() - t_start
            _write_log(f"{ticker} 查可用期間完成，耗時 {elapsed:.1f} 秒", "OK")
            self.msg_queue.put(("preview_scan_done", result))
        except Exception as e:
            elapsed = time.time() - t_start
            # 不把 str(e) 原文丟給使用者——edgartools 的 CompanyNotFoundError 訊息
            # 挾帶 "Tip: Search by name with find_company(...)" 這種給開發者看的 API
            # 建議，使用者看了只會更困惑。只留類型名，UI 端自己組使用者看得懂的話。
            # log 檔不受這個限制，寫完整例外內容方便事後查根因。
            _write_log(f"{ticker} 查可用期間失敗，耗時 {elapsed:.1f} 秒 -> {type(e).__name__}: {e}", "ERROR")
            self.msg_queue.put(("preview_scan_error", (ticker, type(e).__name__)))

    _FIXED_SHEETS = frozenset({"Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"})

    # 視窗高度鎖死不會自動撐大（見 __init__ 的 geometry() 註解）。可選 Sheet 面板
    # 展開時 Tab 1 需要的高度比閒置時多，而下面「處理進度」的 log 是唯一 weight=1
    # 的列，多出來的高度全由它吸收，壓到剩 1px 就等於消失。
    #
    # 原本的解法是面板展開/收合時把視窗在 700x650 與 700x800 之間切換。那會讓
    # 視窗在掃描完成的瞬間自己長高 150px，CTH 回報的「視窗現在很高」就是掃完
    # 之後的那個狀態。改成**單一尺寸、永不跳動**，靠三件事把高度需求壓下來：
    #
    #   1. 寬度改算比例後普遍比舊版 900 寬，可選 Sheet 面板維持 4 欄
    #   2. 面板的固定高度容器 60px（4 欄排得下 8 張 sheet，再多還是可以捲，
    #      不會撐開視窗）
    #   3. 高度改算比例（見 __init__）：2026-08-18 併入 E16 修正後量測，
    #      Tab1 實際內容需要 ~414px（比舊版設計時的 343px 高，這段期間陸續加了
    #      SEC Identity 警告、輸出設定「設為預設」說明等新行），拉高視窗才能讓
    #      log 區不要被壓縮回原本的 4~5 行
    #
    # 下面四個是 __init__ 算 _WIN_W/_WIN_H 用的比例與夾限（clamp），不是視窗本身
    # 的尺寸——實際尺寸依螢幕大小算，小螢幕會自動縮小，大螢幕不會無限撐大。
    _WIN_W_RATIO = 2 / 3
    _WIN_W_MIN, _WIN_W_MAX = 900, 1300
    _WIN_H_RATIO = 0.8
    _WIN_H_MIN, _WIN_H_MAX = 780, 900

    def _build_sheet_panel(self, sheet_names: list[str]):
        """Populate sheet selection panel with checkboxes. Fixed sheets are disabled."""
        if not self._sheet_panel_frame:
            return
        if not sheet_names:
            self._sheet_panel_frame.grid_remove()
            return
        for w in self._sheet_panel_inner.winfo_children():
            w.destroy()
        self._sheet_check_vars = {}

        # 4 欄排列而非單欄直排——sheet 數一多（segment 軸拆出多張）減少捲動需求。
        # 視窗加寬到 900 之後才排得下第 4 欄，這也是高度容器能從 90px 收到 60px
        # 的前提（見 _WIN_W 的註解）。外層 _sheet_panel_inner 是固定高度可捲動
        # 容器，超過還是可以捲，不會撐高視窗。
        _COLS = 4
        for i, name in enumerate(sheet_names):
            var = tk.BooleanVar(value=True)
            self._sheet_check_vars[name] = var
            is_fixed = name in self._FIXED_SHEETS
            cb = ttk.Checkbutton(
                self._sheet_panel_inner, text=name, variable=var,
                state="disabled" if is_fixed else "normal",
            )
            cb.grid(row=i // _COLS, column=i % _COLS, sticky="w", padx=4, pady=1)

        self._sheet_panel_frame.grid()

    def _start_worker(self, target):
        """Clear log, disable run buttons, and start target as a daemon thread. Guards against double-runs."""
        if self.is_running:
            return
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self.btn_open_folder.pack_forget()
        self.progress_bar["value"] = 0
        self.progress_label.config(text=t("gui.status.preparing"))
        self.is_running = True
        self.btn_run_single.config(state="disabled")
        self.btn_run_batch.config(state="disabled")
        if self._scan_btn:
            self._scan_btn.config(state="disabled")
        threading.Thread(target=target, daemon=True).start()

    def _worker_single(self, ticker: str, fetch_gaap: bool, fetch_nongaap: bool,
                       max_filings: int = 80, fetch_q: bool = True, fetch_k: bool = True,
                       start_year: int | None = None, end_year: int | None = None,
                       excluded_sheets: set[str] | None = None):
        """Background thread: orchestrate GAAP and/or Non-GAAP fetch, then write to Excel."""
        try:
            identity = self.cfg.get("identity", "")
            if not identity:
                self._log(t("gui.log.need_identity_full"))
                self._done(False)
                return

            # ---- 任務起始行：只記 ticker + 設定，不記 payload / key ----
            srcs = [s for s, on in (("GAAP", fetch_gaap), ("NonGAAP", fetch_nongaap)) if on]
            kinds = [k for k, on in (("10-Q", fetch_q), ("10-K", fetch_k)) if on]
            scope = f"{start_year or ''}-{end_year or ''}" if (start_year or end_year) else f"max{max_filings}筆"
            task_start = time.time()
            _write_log_header(f"抓取 {ticker} | {'+'.join(srcs) or '無'} | {'/'.join(kinds) or '無'} | {scope}")

            tables = []
            output_path = self._build_output_path(ticker)
            output_dir  = output_path.parent
            output_dir.mkdir(parents=True, exist_ok=True)

            # 抓之前先確認寫得進去。原本的失敗點在最後一步的 wb.save()，
            # 檔案被 Excel 開著時使用者要白等一分多鐘才看到 PermissionError。
            lock_msg = check_output_writable(output_path)
            if lock_msg:
                self._log(f"[{ticker}] ✗ {lock_msg}")
                _write_log(f"{ticker} 輸出檔無法寫入，未開始抓取", "ERROR")
                return

            total_steps = sum([fetch_gaap, fetch_nongaap]) + 1  # +1 for write
            step = 0

            if fetch_gaap:
                self._log(t("gui.log.fetching_gaap", ticker=ticker))
                self._set_progress(step, total_steps, t("gui.status.fetching_gaap"))

                # 逐份 filing 推進度（TODO E12）：GAAP 抓取要對每份 10-Q/10-K
                # 分別建 IS/BS/CF 三張表，幾十份 filing 跑下來可能要幾分鐘，
                # 中途原本進度條完全不動、看起來像卡死。這裡直接借用進度條，
                # 抓取這段期間顯示「第幾份/共幾份」，跟 NonGAAP 那段的
                # `_ng_progress` 是同一個做法——暫時把進度條的刻度換成這段自己
                # 的 current/total，離開這段之後下一次 `_set_progress` 呼叫
                # （NonGAAP 開始或寫檔開始）會自然把刻度換回 total_steps 那個
                # 大尺度，不用特地在這裡復原。不逐行寫 log——幾十次 tick 全部
                # 寫進畫面上的 log 會洗版，只有進度條本身跟著動就夠了。
                def _gaap_progress(current, total, _label):
                    self._set_progress(current, total,
                                        t("gui.status.fetching_gaap_n", current=current, total=total))

                # 開帳本才拿得到缺漏明細。不開的話 fetch_gaap_statements 會
                # 自己開一本，但那本在函式回傳後就沒了，這裡讀不到。
                with collect_gaps() as gaps, report_progress(_gaap_progress):
                    gaap_tables = fetch_gaap_statements(
                        ticker, identity, max_filings=max_filings,
                        ai_config=self.cfg.get("ai", {}),
                        start_year=start_year, end_year=end_year,
                        fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                        excluded_sheets=excluded_sheets or set(),
                    )
                tables.extend(gaap_tables)
                self._log(t("gui.log.gaap_got", ticker=ticker, n=len(gaap_tables)))
                if gaps.has_gaps:
                    # 橘字警告 + 落檔。使用者不必自己去比對少了哪幾期。
                    self._log(gaps.summary(), "WARN", to_file=True)
                if ticker.upper() in _FINANCIAL_SECTOR_TICKERS:
                    self._log(t("gui.log.financial_sector_warning", ticker=ticker))
                step += 1

            if fetch_nongaap and NONGAAP_ENABLED:
                from fetcher_nongaap import fetch_nongaap_statements
                ai_config = self.cfg.get("ai", {})
                self._log(t("gui.log.fetching_nongaap", ticker=ticker))
                self._set_progress(step, total_steps, t("gui.status.fetching_nongaap"))

                def _ng_progress(current, total, label):
                    self._log(f"[{ticker}] {label}")
                    self._set_progress(current, total, label)

                ng_tables = fetch_nongaap_statements(
                    ticker, identity, ai_config,
                    output_dir=output_dir,
                    progress_cb=_ng_progress,
                    max_filings=max_filings,
                    start_year=start_year, end_year=end_year,
                )
                tables.extend(ng_tables)
                self._log(t("gui.log.nongaap_sheets", ticker=ticker, n=len(ng_tables)))
                step += 1

            # 一期都沒抓到就不寫——空殼 Excel 會蓋掉使用者原本好好的舊檔。
            # 缺幾期是可以接受的（上面已經警告過），全空不行。
            if not has_any_data(tables):
                self._log(t("gui.log.nothing_to_write"), "ERROR", to_file=True)
                self._done(False)
                return

            _append_ratio_table(tables)

            self._log(t("gui.log.writing_excel", ticker=ticker))
            self._set_progress(step, total_steps, t("gui.status.writing_excel"))
            tpl = self.cfg.get("template_path", "") or None
            write_statements(tables, output_path, template_path=tpl)
            self._log(t("gui.log.done_file", ticker=ticker, name=output_path.name))
            self._set_progress(total_steps, total_steps, t("gui.status.done"))
            # ---- 任務結果行：成功 + 耗時 ----
            _elapsed = int(time.time() - task_start)
            _write_log(f"{ticker} 成功，耗時 {_elapsed // 60}分{_elapsed % 60}秒", "OK")
            self.msg_queue.put(("last_output_folder", output_path.parent))
            self._done(True)

        except Exception as e:
            # ---- 錯誤行：只記 type + status，禁止 {e} 全文（會挾帶 URL/response/key）----
            self._log(t("gui.log.fetch_failed", ticker=ticker,
                                exc=f"{type(e).__name__}{_exc_status(e)}"), "ERROR", to_file=True)
            try:
                _elapsed = int(time.time() - task_start)
                _write_log(f"{ticker} 失敗，耗時 {_elapsed // 60}分{_elapsed % 60}秒", "FAIL")
            except NameError:
                pass  # 例外發生在 task_start 賦值前（identity 檢查等）
            self._done(False)

    def _worker_batch(self, tickers: list[str], fetch_nongaap: bool = False,
                      fetch_q: bool = True, fetch_k: bool = True,
                      start_year: int | None = None, end_year: int | None = None):
        """Background thread: fetch GAAP (and optionally Non-GAAP) for each ticker."""
        total = len(tickers)
        identity = self.cfg.get("identity", "")
        if not identity:
            self._log(t("gui.log.need_identity"))
            self._done(False)
            return
        max_filings = self.cfg.get("max_filings", 80)
        ai_config   = self.cfg.get("ai", {})

        srcs = "GAAP+NonGAAP" if fetch_nongaap else "GAAP"
        kinds = [k for k, on in (("10-Q", fetch_q), ("10-K", fetch_k)) if on]
        scope = f"{start_year or ''}-{end_year or ''}" if (start_year or end_year) else f"max{max_filings}筆"

        for i, ticker in enumerate(tickers, 1):
            self._set_progress(i - 1, total,
                               t("gui.status.processing", ticker=ticker, i=i, total=total))
            self._log("\n" + t("gui.log.starting", ticker=ticker))
            # ---- 任務起始行：只記 ticker + 設定 ----
            task_start = time.time()
            _write_log_header(f"批量抓取 {ticker} ({i}/{total}) | {srcs} | {'/'.join(kinds) or '無'} | {scope}")
            try:
                # 批次一樣會卡在單一 ticker 抓很久（TODO E12），同一顆 callback
                # 沿用單一公司那邊的做法：抓這個 ticker 期間暫時把進度條換成
                # 「這個 ticker 內第幾份/共幾份」，抓完下一輪迴圈開頭的
                # `self._set_progress(i-1, total, ...)` 會換回「第幾間公司」
                # 那個大尺度
                def _gaap_progress(current, total_n, _label):
                    self._set_progress(current, total_n,
                                        t("gui.status.fetching_gaap_n", current=current, total=total_n))

                with collect_gaps() as gaps, report_progress(_gaap_progress):
                    tables = fetch_gaap_statements(
                        ticker, identity, max_filings=max_filings, ai_config=ai_config,
                        start_year=start_year, end_year=end_year,
                        fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                    )
                if gaps.has_gaps:
                    self._log(f"[{ticker}] {gaps.summary()}", "WARN", to_file=True)

                if ticker.upper() in _FINANCIAL_SECTOR_TICKERS:
                    self._log(t("gui.log.financial_sector_warning", ticker=ticker))

                if fetch_nongaap and NONGAAP_ENABLED:
                    from fetcher_nongaap import fetch_nongaap_statements
                    output_path = self._build_output_path(ticker)
                    output_dir  = output_path.parent
                    output_dir.mkdir(parents=True, exist_ok=True)

                    def _ng_cb(current, total_ng, label, _t=ticker):
                        self._log(f"[{_t}] {label}")

                    ng_tables = fetch_nongaap_statements(
                        ticker, identity, ai_config,
                        output_dir=output_dir,
                        progress_cb=_ng_cb,
                        max_filings=max_filings,
                        start_year=start_year, end_year=end_year,
                    )
                    tables.extend(ng_tables)
                    self._log(t("gui.log.nongaap_sheets", ticker=ticker, n=len(ng_tables)))

                if not has_any_data(tables):
                    # 批次不中斷整批，只跳過這一家（跟輸出檔被鎖同樣處理）
                    self._log(f"[{ticker}] {t('gui.log.nothing_to_write')}",
                              "ERROR", to_file=True)
                    _write_log(f"{ticker} 一期都沒抓到，未寫出檔案", "FAIL")
                    continue

                _append_ratio_table(tables)
                output_path = self._build_output_path(ticker)
                lock_msg = check_output_writable(output_path)
                if lock_msg:
                    # 批次模式不中斷整批，只跳過這一家
                    self._log(f"[{ticker}] ✗ {lock_msg}")
                    _write_log(f"{ticker} 輸出檔無法寫入，已跳過", "ERROR")
                    continue
                tpl = self.cfg.get("template_path", "") or None
                write_statements(tables, output_path, template_path=tpl)
                self._log(t("gui.log.done_count", ticker=ticker, n=len(tables)))
                # ---- 任務結果行：成功 + 耗時 ----
                _elapsed = int(time.time() - task_start)
                _write_log(f"{ticker} 成功，耗時 {_elapsed // 60}分{_elapsed % 60}秒", "OK")
            except Exception as e:
                # ---- 錯誤行：只記 type + status，禁止 {e} 全文 ----
                self._log(t("gui.log.fetch_failed", ticker=ticker,
                                exc=f"{type(e).__name__}{_exc_status(e)}"), "ERROR", to_file=True)
                _elapsed = int(time.time() - task_start)
                _write_log(f"{ticker} 失敗，耗時 {_elapsed // 60}分{_elapsed % 60}秒", "FAIL")

        self._set_progress(total, total, t("gui.status.batch_done", total=total))
        self.msg_queue.put(("last_output_folder", self._build_output_path(tickers[-1]).parent))
        self._done(True)

    # =========================================================
    # Thread-safe queue helpers
    # =========================================================

    def _log(self, msg: str, level: str = "INFO", to_file: bool = False):
        """Queue a log line for the UI; optionally also append it to logs/app.log.

        預設 to_file=False（fail-closed）：畫面上的進度／成功訊息只推 UI，不落檔。
        真正該落檔的只有任務起始、錯誤、結果三種，由呼叫端明確傳 to_file=True，
        或直接呼叫 _write_log_header / _write_log。這樣可避免任何進度訊息意外寫上磁碟。
        """
        if to_file:
            _write_log(msg, level)
        self.msg_queue.put(("log", msg))

    def _init_log(self, msg: str):
        self.log_text.config(state="normal")
        self.log_text.insert("1.0", msg + "\n")
        self.log_text.config(state="disabled")

    def _set_progress(self, current: int, total: int, label: str):
        """Queue a progress bar update to be applied in the main thread."""
        self.msg_queue.put(("progress", (current, total, label)))

    def _done(self, success: bool):
        """Queue a run-complete signal so _poll_queue re-enables buttons."""
        self.msg_queue.put(("done", success))

    def _poll_queue(self):
        """Drain msg_queue and apply GUI updates. Runs every 100ms on the main thread.

        All worker threads push results here rather than touching Tkinter directly,
        because Tkinter is not thread-safe.
        """
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()

                if msg_type == "log":
                    self.log_text.config(state="normal")
                    self.log_text.insert("end", data + "\n")
                    self.log_text.see("end")
                    self.log_text.config(state="disabled")

                elif msg_type == "progress":
                    current, total, label = data
                    self.progress_bar["maximum"] = total
                    self.progress_bar["value"]   = current
                    self.progress_label.config(text=label)

                elif msg_type == "done":
                    success = data
                    self.is_running = False
                    self.btn_run_single.config(state="normal")
                    self.btn_run_batch.config(state="normal")
                    if self._scan_btn:
                        self._scan_btn.config(state="normal")
                    if success:
                        self.btn_open_folder.pack(side="left")
                        self.progress_label.config(text=t("gui.status.done"))
                    else:
                        self.progress_label.config(text=t("gui.status.error_see_log"))

                elif msg_type == "tab1_name_result":
                    status, looked_ticker, name = data
                    current = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
                    if self.tab1_name_label and current == looked_ticker:
                        if status == "ok":
                            self.tab1_name_label.config(text=f"　{name}", foreground="#1a7a34")
                            # 自動帶出已記憶的路徑
                            saved_path = self.cfg.get("ticker_paths", {}).get(looked_ticker)
                            if saved_path and self.tab1_outdir_var:
                                self.tab1_outdir_var.set(saved_path)
                        else:
                            self.tab1_name_label.config(text=t("gui.status.ticker_not_found"), foreground="orange")
                        self._update_tab1_preview()
                    if self.btn_confirm_company:
                        self.btn_confirm_company.config(state="normal")

                elif msg_type == "wl_lookup_result":
                    status = data[0]
                    if status == "ok":
                        _, ticker, name = data
                        self._wl_found_name = name
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=t("gui.wl.found", name=name), foreground="#1a7a34")
                        if self.wl_add_btn:
                            self.wl_add_btn.config(state="normal")
                        self._wl_add()
                    else:
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=t("gui.wl.lookup_failed", reason=data[1]), foreground="red")

                elif msg_type == "wl_cache_updated":
                    update_date, count = data
                    if self.wl_cache_label:
                        self.wl_cache_label.config(
                            text=t("gui.wl.cache_loaded", count=f"{count:,}", updated=update_date),
                            foreground="gray"
                        )

                elif msg_type == "wl_cache_update_error":
                    if self.wl_cache_label:
                        self.wl_cache_label.config(text=t("gui.wl.update_failed", reason=data), foreground="red")

                elif msg_type == "last_output_folder":
                    self._last_output_folder = data

                elif msg_type == "ai_test_result":
                    ok, err = data
                    if self.settings_test_label:
                        if ok == "ok":
                            self.settings_test_label.config(text=t("gui.msg.connection_ok"), foreground="#1a7a34")
                        else:
                            self.settings_test_label.config(text=t("gui.msg.failed", reason=str(err)[:60]), foreground="red")

                elif msg_type == "preview_scan_done":
                    self._build_sheet_panel(data["sheets"])
                    if self._sheet_panel_frame:
                        label, end, fdate = data["latest_label"], data["latest_period_end"], data["filing_date"]
                        info = (t("gui.status.latest_data", label=label, end=end, filed=fdate)
                                if label else t("gui.status.latest_unknown"))
                        self._sheet_panel_frame.configure(text=f"{self._SHEET_PANEL_TITLE_BASE} ｜ {info}")
                    if self._scan_btn:
                        self._scan_btn.config(state="normal", text=t("gui.btn.scan"))
                    if self._scan_hint_label:
                        self._scan_hint_label.config(text="")

                elif msg_type == "preview_scan_error":
                    if self._scan_btn:
                        self._scan_btn.config(state="normal", text=t("gui.btn.scan"))
                    if self._scan_hint_label:
                        self._scan_hint_label.config(text="")
                    ticker, exc_name = data
                    if exc_name in ("CompanyNotFoundError", "ValueError"):
                        msg = t("gui.msg.ticker_not_found", ticker=ticker)
                    else:
                        msg = t("gui.msg.scan_failed", reason=exc_name)
                    messagebox.showerror(t("gui.dlg.scan_failed_title"), msg)

                elif msg_type == "compare_error":
                    self._log(f"{t('gui.compare.select_title')}: {data}", "ERROR")
                    self.progress_label.config(text=t("gui.status.error_see_log"))
                    self.compare_run_btn.config(state="normal")

                elif msg_type == "compare_done":
                    self._log(t("gui.compare.log_done", path=data))
                    self.progress_label.config(text=t("gui.status.done"))
                    self.compare_run_btn.config(state="normal")

        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# =========================================================
# Entry point
# =========================================================

def _pick_language_on_first_run(root: tk.Tk) -> None:
    """首次啟動時問一次語言，選完寫進 config.json，之後不再出現。

    「還沒選過」的判斷依據是 `config.json` 的 `language` **不是合法代號**
    （空字串、缺這個鍵、或舊版留下的怪值）。不另外開一個 `language_chosen`
    布林值——兩個欄位描述同一件事，遲早會不同步。

    視窗本身刻意**不翻譯**：這時候還不知道使用者要哪個語言，用任一種當說明
    都在賭。所以只有一個英文抬頭，其餘全是各語言的自稱，看得懂哪個就點哪個。

    直接關掉視窗＝接受第一個選項（繁體中文）並且**照樣存檔**——需求是「選完
    就記住不要再跳」，關掉還一直跳才是煩人。選錯了在「進階設定」隨時能改。
    """
    _migrate_config_if_needed()
    cfg = load_config(CONFIG_PATH)
    if i18n.is_supported(cfg.get("language", "")):
        return                      # 選過了，直接進主畫面

    choices = i18n.available_languages()
    chosen = {"code": choices[0][0]}

    dlg = tk.Toplevel(root)
    dlg.title("Language")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)

    ttk.Label(dlg, text="Select your language",
              font=("", 12, "bold")).pack(padx=28, pady=(20, 4))
    ttk.Label(dlg, text="You can change this later in Settings.",
              foreground="#555555").pack(padx=28, pady=(0, 14))

    def _choose(code: str) -> None:
        chosen["code"] = code
        dlg.destroy()

    for code, name in choices:
        ttk.Button(dlg, text=name, width=20,
                   command=lambda c=code: _choose(c)).pack(padx=28, pady=3)
    ttk.Frame(dlg, height=10).pack()

    _center_on_parent(dlg, root)

    dlg.grab_set()
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)   # 關掉＝用預設值，照樣存
    root.wait_window(dlg)

    cfg["language"] = chosen["code"]
    save_config(cfg, CONFIG_PATH)


def main():
    """Entry point: show banner, create Tk root, launch app, enter event loop."""
    show_cth_banner()
    # edgartools 預設不設 HTTP 逾時，落到 httpx 自己的 5 秒——抓大份 filing
    # 時太短，會誤判成逾時再被白白重試三遍。給一個明確的值。
    configure_timeouts()
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.update()
    root.attributes("-topmost", False)
    _pick_language_on_first_run(root)
    SECFetcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
