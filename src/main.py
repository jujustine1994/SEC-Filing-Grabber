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
from fetcher_gaap import fetch_gaap_statements
from output_tables import append_ratio_table

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
        _area = work_area()
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

        # Persistent buttons
        frame_persist = tk.Frame(self.root)
        frame_persist.grid(row=1, column=0, pady=4)
        ttk.Button(frame_persist, text=t("gui.btn.manage_watchlist"), command=self._open_watchlist_popup, width=18).pack(side="left", padx=6)
        ttk.Button(frame_persist, text=t("gui.btn.settings"),       command=self._open_settings_popup,  width=14).pack(side="left", padx=6)

        # Progress log
        frame_log = ttk.LabelFrame(self.root, text=t("gui.frame.progress"), padding=8)
        frame_log.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 4))
        frame_log.rowconfigure(2, weight=1)
        frame_log.columnconfigure(0, weight=1)
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

        # Row 0: Ticker + inline company name
        row_ticker = ttk.Frame(tab)
        row_ticker.grid(row=0, column=0, sticky="ew", pady=4)
        ttk.Label(row_ticker, text="Ticker:").pack(side="left", padx=(0, 8))
        self.ticker_var = tk.StringVar()
        self.ticker_entry = ttk.Entry(row_ticker, textvariable=self.ticker_var, width=12, foreground="grey")
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

        # Row 1: Checkboxes
        row_type = ttk.Frame(tab)
        row_type.grid(row=1, column=0, sticky="ew", pady=4)
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

        # Row 2: 進階設定 toggle
        adv_toggle_row1 = ttk.Frame(tab)
        adv_toggle_row1.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        self._tab1_adv_toggle_btn = ttk.Button(adv_toggle_row1, text=t("gui.btn.adv_collapsed"),
                                                command=self._toggle_tab1_adv, width=12)
        self._tab1_adv_toggle_btn.pack(side="left")

        # Row 3: 進階設定 content — report type (hidden by default)
        self._tab1_adv_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=(8, 4))
        self._tab1_adv_frame.grid(row=3, column=0, sticky="ew", pady=(0, 4))
        self.tab1_fetch_q_var = tk.BooleanVar(value=True)
        self.tab1_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self._tab1_adv_frame, text=t("gui.chk.quarterly"), variable=self.tab1_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(self._tab1_adv_frame, text=t("gui.chk.annual"), variable=self.tab1_fetch_k_var).pack(side="left")
        self._tab1_adv_frame.grid_remove()

        # Row 4: Date range
        row_date = ttk.Frame(tab)
        row_date.grid(row=4, column=0, sticky="ew", pady=(2, 4))
        ttk.Label(row_date, text=t("gui.lbl.year_from")).pack(side="left", padx=(0, 4))
        self.tab1_start_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_start_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text=t("gui.lbl.year_to")).pack(side="left", padx=(8, 4))
        self.tab1_end_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_end_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text=t("gui.lbl.year_hint"), foreground="#555555").pack(side="left", padx=(4, 0))

        # Row 5: Sheet selection panel (hidden until scan completes)
        # 最新季度／送件日不能另開一行 Label——這個視窗鎖死 650px 高、不會自動撐大
        # （見 __init__ 的 geometry() 註解），下面「處理進度」的 log 區已經很緊繃，
        # 實測顯示 sheet 面板一展開，log 可視高度就只剩個位數 px；多加一行 23px
        # 會直接把 log 擠到全隱形。改寫進 LabelFrame 自己的標題列，不佔新的一行，
        # 高度成本是 0
        self._SHEET_PANEL_TITLE_BASE = t("gui.frame.optional_sheets")
        self._sheet_panel_frame = ttk.LabelFrame(tab, text=self._SHEET_PANEL_TITLE_BASE, padding=6)
        self._sheet_panel_frame.grid(row=5, column=0, sticky="ew", pady=(0, 4))
        _, self._sheet_panel_inner = _build_fixed_height_scrollable(self._sheet_panel_frame, height=60)
        self._sheet_panel_frame.grid_remove()

        # Row 6: Non-GAAP warning (hidden by default)
        self.nongaap_warn_label = ttk.Label(
            tab, text=t("gui.lbl.nongaap_need_key"),
            foreground="orange", font=("", 10)
        )
        self.nongaap_warn_label.grid(row=6, column=0, sticky="w", padx=2)
        self.nongaap_warn_label.grid_remove()

        # Row 7: Output settings toggle
        self._out_collapsed = False
        out_toggle_row = ttk.Frame(tab)
        out_toggle_row.grid(row=7, column=0, sticky="ew", pady=(8, 0))
        self._out_toggle_btn = ttk.Button(out_toggle_row, text=t("gui.btn.output_expanded"),
                                           command=self._toggle_out_settings, width=12)
        self._out_toggle_btn.pack(side="left")

        # Row 8: Output settings content (collapsible)
        out_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=8)
        out_frame.grid(row=8, column=0, sticky="ew", pady=(0, 4))
        self._out_settings_frame = out_frame

        # Storage location row
        loc_row = ttk.Frame(out_frame)
        loc_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(loc_row, text=t("gui.lbl.save_location")).pack(side="left")
        self.tab1_outdir_var = tk.StringVar(value=self.cfg.get("output_dir", "output"))
        ttk.Entry(loc_row, textvariable=self.tab1_outdir_var, width=26).pack(side="left", padx=(0, 6))
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

        # Row 5: Execute button
        self.btn_run_single = ttk.Button(tab, text=t("gui.btn.run"), command=self._run_single, width=16)
        self.btn_run_single.grid(row=9, column=0, pady=(8, 4))

    def _toggle_tab1_adv(self):
        self._tab1_adv_collapsed = not self._tab1_adv_collapsed
        if self._tab1_adv_collapsed:
            self._tab1_adv_frame.grid_remove()
            self._tab1_adv_toggle_btn.config(text=t("gui.btn.adv_collapsed"))
        else:
            self._tab1_adv_frame.grid()
            self._tab1_adv_toggle_btn.config(text=t("gui.btn.adv_expanded"))

    def _toggle_tab2_adv(self):
        self._tab2_adv_collapsed = not self._tab2_adv_collapsed
        if self._tab2_adv_collapsed:
            self._tab2_adv_frame.grid_remove()
            self._tab2_adv_toggle_btn.config(text=t("gui.btn.adv_collapsed"))
        else:
            self._tab2_adv_frame.grid()
            self._tab2_adv_toggle_btn.config(text=t("gui.btn.adv_expanded"))

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

        self.tab2_list_frame = ttk.LabelFrame(tab, text=" Watchlist ", padding=6)
        self.tab2_list_frame.grid(row=0, column=0, sticky="ew", pady=4)
        tab.columnconfigure(0, weight=1)
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

        row_sel = ttk.Frame(tab)
        row_sel.grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(row_sel, text=t("gui.btn.select_all"),   command=self._select_all,   width=8).pack(side="left", padx=(0, 8))
        ttk.Button(row_sel, text=t("gui.btn.select_none"), command=self._deselect_all, width=8).pack(side="left")

        row_opts = ttk.Frame(tab)
        row_opts.grid(row=2, column=0, sticky="w", pady=(4, 0))
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

        # Row 3: 進階設定 toggle
        adv_toggle_row2 = ttk.Frame(tab)
        adv_toggle_row2.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self._tab2_adv_toggle_btn = ttk.Button(adv_toggle_row2, text=t("gui.btn.adv_collapsed"),
                                                command=self._toggle_tab2_adv, width=12)
        self._tab2_adv_toggle_btn.pack(side="left")

        # Row 4: 進階設定 content — report type (hidden by default)
        self._tab2_adv_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=(8, 4))
        self._tab2_adv_frame.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        self.batch_fetch_q_var = tk.BooleanVar(value=True)
        self.batch_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self._tab2_adv_frame, text=t("gui.chk.quarterly"), variable=self.batch_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(self._tab2_adv_frame, text=t("gui.chk.annual"), variable=self.batch_fetch_k_var).pack(side="left")
        self._tab2_adv_frame.grid_remove()

        # Row 5: Date range
        row_date2 = ttk.Frame(tab)
        row_date2.grid(row=5, column=0, sticky="ew", pady=(2, 0))
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
        self.btn_run_batch.grid(row=6, column=0, pady=(8, 4))

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
        """Enable/disable the custom filename entry and refresh the preview when format radio changes."""
        is_custom = self.tab1_fmt_var.get() == "custom"
        if self.tab1_custom_entry:
            self.tab1_custom_entry.config(state="normal" if is_custom else "disabled")
        self._save_tab1_output_settings()
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
        """Open folder picker and save selection globally and as a per-ticker path memory."""
        from tkinter import filedialog
        current = self.tab1_outdir_var.get().strip() if self.tab1_outdir_var else "output"
        initial = str(PROJECT_ROOT / current) if not os.path.isabs(current) else current
        folder = filedialog.askdirectory(title=t("gui.dlg.choose_output_dir"), initialdir=initial)
        if folder:
            self.tab1_outdir_var.set(folder)
            # 記住這個 ticker 的路徑
            ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
            if ticker:
                if "ticker_paths" not in self.cfg:
                    self.cfg["ticker_paths"] = {}
                self.cfg["ticker_paths"][ticker] = folder
            self._save_tab1_output_settings()

    def _save_tab1_output_settings(self):
        """Persist Tab 1 output settings (dir, filename format, custom name) to config.json."""
        if self.tab1_outdir_var:
            self.cfg["output_dir"] = self.tab1_outdir_var.get().strip() or "output"
        if self.tab1_fmt_var:
            self.cfg["filename_format"] = self.tab1_fmt_var.get()
        if self.tab1_custom_var:
            self.cfg["filename_custom"] = self.tab1_custom_var.get().strip()
        save_config(self.cfg, CONFIG_PATH)

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

        wl_canvas = tk.Canvas(list_frame, height=200, highlightthickness=0)
        wl_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=wl_canvas.yview)
        wl_canvas.configure(yscrollcommand=wl_scrollbar.set)
        wl_canvas.grid(row=0, column=0, sticky="ew")
        wl_scrollbar.grid(row=0, column=1, sticky="ns")
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

            # Group header
            hdr = ttk.Frame(container)
            hdr.pack(fill="x", pady=(6, 0))
            arrow = "▶" if is_collapsed else "▼"
            ttk.Button(hdr, text=f"{arrow} {_group_display(gname)}", width=16,
                       command=lambda g=gname, c=container: self._wl_toggle_group(g, c)).pack(side="left")
            ttk.Button(hdr, text=t("gui.btn.rename"), width=8,
                       command=lambda g=gname, c=container: self._wl_rename_group(g, c)).pack(side="left", padx=(4, 0))
            if gname != UNCATEGORIZED:
                ttk.Button(hdr, text=t("gui.btn.delete_group"), width=8,
                           command=lambda g=gname, c=container: self._wl_delete_group(g, c)).pack(side="left", padx=(4, 0))

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
    # Advanced settings popup
    # =========================================================

    def _open_settings_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title(t("gui.btn.settings"))
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)
        popup.update()
        popup.attributes("-topmost", False)
        popup.bind("<Escape>", lambda e: popup.destroy())
        self._build_settings_popup(popup)

    def _build_settings_popup(self, popup: tk.Toplevel):
        """Build settings popup: SEC identity, AI config, fetch limits, template mode."""
        pad = {"padx": 12, "pady": 4}

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
        ttk.Label(id_frame, text=t("gui.lbl.identity_hint"),
                  foreground="#555555", font=("", 10)).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(id_frame, text="Identity:").grid(row=1, column=0, sticky="w", pady=4)
        self.settings_identity_var = tk.StringVar(value=self.cfg.get("identity", ""))
        ttk.Entry(id_frame, textvariable=self.settings_identity_var, width=42).grid(row=1, column=1, sticky="ew", padx=(8, 0))

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

        # Buttons
        btn_row = ttk.Frame(popup)
        btn_row.grid(row=4, column=0, pady=10)
        ttk.Button(btn_row, text=t("gui.btn.save"), command=lambda: self._save_settings(popup), width=10).pack(side="left", padx=6)
        ttk.Button(btn_row, text=t("gui.btn.cancel"), command=popup.destroy, width=10).pack(side="left", padx=6)

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

    def _save_settings(self, popup: tk.Toplevel):
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
        save_config(self.cfg, CONFIG_PATH)
        popup.destroy()
        # 只有語言真的變更才打擾使用者——改 API Key 不該跳重啟視窗
        if lang_changed:
            self._prompt_restart_for_language()

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
        if self.is_running:
            messagebox.showwarning(t("gui.dlg.info_title"), t("gui.msg.wait_for_current_run"))
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
        """Background thread: call preview_sheets() and push result to queue."""
        try:
            from fetcher_gaap import preview_sheets
            result = preview_sheets(ticker, identity)
            self.msg_queue.put(("preview_scan_done", result))
        except Exception as e:
            # 不把 str(e) 原文丟給使用者——edgartools 的 CompanyNotFoundError 訊息
            # 挾帶 "Tip: Search by name with find_company(...)" 這種給開發者看的 API
            # 建議，使用者看了只會更困惑。只留類型名，UI 端自己組使用者看得懂的話。
            self.msg_queue.put(("preview_scan_error", (ticker, type(e).__name__)))

    _FIXED_SHEETS = frozenset({"Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"})

    # 視窗高度鎖死不會自動撐大（見 __init__ 的 geometry() 註解）。可選 Sheet 面板
    # 展開時 Tab 1 需要的高度比閒置時多，而下面「處理進度」的 log 是唯一 weight=1
    # 的列，多出來的高度全由它吸收，壓到剩 1px 就等於消失。
    #
    # 原本的解法是面板展開/收合時把視窗在 700x650 與 700x800 之間切換。那會讓
    # 視窗在掃描完成的瞬間自己長高 150px，CTH 回報的「視窗現在很高」就是掃完
    # 之後的那個狀態。現在改成**單一尺寸、永不跳動**，靠三件事把高度需求壓下來：
    #
    #   1. 寬度 700 -> 900，可選 Sheet 面板從 3 欄改 4 欄
    #   2. 面板的固定高度容器 90px -> 60px（4 欄之後兩列就放得下 8 張 sheet，
    #      再多還是可以捲，不會撐開視窗）
    #   3. 總高度取 720：面板展開時 log 仍有 4~5 行可視，收合時約 9 行
    #
    # 這兩個值是 fit_geometry 的輸入，實際擺出來的尺寸在小螢幕上會被縮。
    _WIN_W = 900
    _WIN_H = 720

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
                gaap_tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings,
                    ai_config=self.cfg.get("ai", {}),
                    start_year=start_year, end_year=end_year,
                    fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                    excluded_sheets=excluded_sheets or set(),
                )
                tables.extend(gaap_tables)
                self._log(t("gui.log.gaap_got", ticker=ticker, n=len(gaap_tables)))
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

            if not tables:
                self._log(t("gui.log.nothing_to_write"))
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
                tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings, ai_config=ai_config,
                    start_year=start_year, end_year=end_year,
                    fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                )

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

    dlg.update_idletasks()
    # 置中在主視窗上，不要跑到螢幕角落
    x = root.winfo_rootx() + (root.winfo_width() - dlg.winfo_width()) // 2
    y = root.winfo_rooty() + (root.winfo_height() - dlg.winfo_height()) // 3
    dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    dlg.grab_set()
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)   # 關掉＝用預設值，照樣存
    root.wait_window(dlg)

    cfg["language"] = chosen["code"]
    save_config(cfg, CONFIG_PATH)


def main():
    """Entry point: show banner, create Tk root, launch app, enter event loop."""
    show_cth_banner()
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.update()
    root.attributes("-topmost", False)
    _pick_language_on_first_run(root)
    SECFetcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
