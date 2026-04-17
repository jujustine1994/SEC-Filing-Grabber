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
import threading
import tkinter as tk
import urllib.request
from datetime import date
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

from config import load_config, save_config, CONFIG_PATH
from excel_writer import write_statements
from fetcher_gaap import fetch_gaap_statements

SCRIPT_DIR = Path(__file__).parent
CACHE_PATH = SCRIPT_DIR / "company_cache.json"


def _migrate_config_if_needed():
    """If old config.json exists in project dir, move it to APPDATA."""
    old_path = SCRIPT_DIR / "config.json"
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

class SECFetcherApp:
    TICKER_PH = "輸入 Ticker（如 AAPL）"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SEC Financial Fetcher")
        self.root.resizable(True, True)
        self.root.minsize(520, 600)

        _migrate_config_if_needed()
        self.cfg = load_config(CONFIG_PATH)
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
        self.tab1_outdir_var = None
        self.tab1_fmt_var = None
        self.tab1_custom_var = None
        self.tab1_custom_entry = None
        self.tab1_preview_label = None

        self._build_ui()
        self._poll_queue()

    # =========================================================
    # UI Construction
    # =========================================================

    def _build_ui(self):
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
        ttk.Button(frame_persist, text="管理 Watchlist", command=self._open_watchlist_popup, width=18).pack(side="left", padx=6)
        ttk.Button(frame_persist, text="進階設定",       command=self._open_settings_popup,  width=14).pack(side="left", padx=6)

        # Progress log
        frame_log = ttk.LabelFrame(self.root, text=" 處理進度 ", padding=8)
        frame_log.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 4))
        frame_log.rowconfigure(2, weight=1)
        frame_log.columnconfigure(0, weight=1)
        self.progress_label = ttk.Label(frame_log, text="等待開始...")
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
            frame_output, text="開啟輸出資料夾", command=self._open_output_folder
        )

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

    def _build_tab1(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  單一公司  ")

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

        # Row 1: Checkboxes
        row_type = ttk.Frame(tab)
        row_type.grid(row=1, column=0, sticky="ew", pady=4)
        self.fetch_gaap_var    = tk.BooleanVar(value=True)
        self.fetch_nongaap_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row_type, text="GAAP 財報",               variable=self.fetch_gaap_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(row_type, text="Non-GAAP（需設定 AI API）", variable=self.fetch_nongaap_var).pack(side="left")
        self.fetch_nongaap_var.trace_add("write", self._on_nongaap_toggle)

        # Row 2: Non-GAAP warning (hidden by default)
        self.nongaap_warn_label = ttk.Label(
            tab, text="⚠ Non-GAAP 需先在「進階設定」填入 AI API Key",
            foreground="orange", font=("", 10)
        )
        self.nongaap_warn_label.grid(row=2, column=0, sticky="w", padx=2)
        self.nongaap_warn_label.grid_remove()

        # Row 3: Output settings toggle
        self._out_collapsed = False
        out_toggle_row = ttk.Frame(tab)
        out_toggle_row.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self._out_toggle_btn = ttk.Button(out_toggle_row, text="▼ 輸出設定",
                                           command=self._toggle_out_settings, width=12)
        self._out_toggle_btn.pack(side="left")

        # Row 4: Output settings content (collapsible)
        out_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=8)
        out_frame.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        self._out_settings_frame = out_frame

        # Storage location row
        loc_row = ttk.Frame(out_frame)
        loc_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(loc_row, text="儲存位置：").pack(side="left")
        self.tab1_outdir_var = tk.StringVar(value=self.cfg.get("output_dir", "output"))
        ttk.Entry(loc_row, textvariable=self.tab1_outdir_var, width=26).pack(side="left", padx=(0, 6))
        ttk.Button(loc_row, text="瀏覽", width=5, command=self._browse_output_dir).pack(side="left")

        # Filename format radios
        ttk.Label(out_frame, text="檔名格式：").grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.tab1_fmt_var = tk.StringVar(value=self.cfg.get("filename_format", "ticker_name"))

        ttk.Radiobutton(out_frame, text="Ticker + 公司名稱（如 AAPL Apple Inc. data.xlsx）",
                        variable=self.tab1_fmt_var, value="ticker_name",
                        command=self._on_tab1_fmt_change).grid(row=2, column=0, sticky="w", padx=(16, 0))
        ttk.Radiobutton(out_frame, text="僅 Ticker（如 AAPL.xlsx）",
                        variable=self.tab1_fmt_var, value="ticker_only",
                        command=self._on_tab1_fmt_change).grid(row=3, column=0, sticky="w", padx=(16, 0))

        custom_row = ttk.Frame(out_frame)
        custom_row.grid(row=4, column=0, sticky="w", padx=(16, 0))
        ttk.Radiobutton(custom_row, text="自訂：",
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
        self.btn_run_single = ttk.Button(tab, text="▶  執行", command=self._run_single, width=16)
        self.btn_run_single.grid(row=5, column=0, pady=(8, 4))

    def _toggle_out_settings(self):
        self._out_collapsed = not self._out_collapsed
        if self._out_collapsed:
            self._out_settings_frame.grid_remove()
            self._out_toggle_btn.config(text="▶ 輸出設定")
        else:
            self._out_settings_frame.grid()
            self._out_toggle_btn.config(text="▼ 輸出設定")

    def _on_nongaap_toggle(self, *_args):
        if self.fetch_nongaap_var.get() and not self.cfg["ai"].get("api_key"):
            self.nongaap_warn_label.grid()
        else:
            self.nongaap_warn_label.grid_remove()

    def _build_tab2(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  批量更新  ")

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
        ttk.Button(row_sel, text="全選",   command=self._select_all,   width=8).pack(side="left", padx=(0, 8))
        ttk.Button(row_sel, text="全不選", command=self._deselect_all, width=8).pack(side="left")

        self.btn_run_batch = ttk.Button(tab, text="▶  開始批量更新", command=self._run_batch, width=20)
        self.btn_run_batch.grid(row=2, column=0, pady=(8, 4))

    # =========================================================
    # Placeholder helpers
    # =========================================================

    def _ph_in(self, entry, var, placeholder):
        if var.get() == placeholder:
            var.set("")
            entry.configure(foreground="black")
            if entry is self.ticker_entry and self.tab1_name_label:
                self.tab1_name_label.config(text="")

    def _ph_out(self, entry, var, placeholder):
        if not var.get().strip():
            var.set(placeholder)
            entry.configure(foreground="grey")

    def _on_ticker_focusout(self, event):
        self._ph_out(self.ticker_entry, self.ticker_var, self.TICKER_PH)
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            if self.tab1_name_label:
                self.tab1_name_label.config(text="")
            self._update_tab1_preview()
            return
        if self.tab1_name_label:
            self.tab1_name_label.config(text="查詢中...", foreground="#555555")
        if self.btn_confirm_company:
            self.btn_confirm_company.config(state="disabled")
        self._update_tab1_preview()
        threading.Thread(target=lambda: self._tab1_lookup_worker(ticker), daemon=True).start()

    def _confirm_company(self):
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            return
        if self.tab1_name_label:
            self.tab1_name_label.config(text="查詢中...", foreground="#555555")
        if self.btn_confirm_company:
            self.btn_confirm_company.config(state="disabled")
        threading.Thread(target=lambda: self._tab1_lookup_worker(ticker), daemon=True).start()

    def _tab1_lookup_worker(self, ticker: str):
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
        is_custom = self.tab1_fmt_var.get() == "custom"
        if self.tab1_custom_entry:
            self.tab1_custom_entry.config(state="normal" if is_custom else "disabled")
        self._save_tab1_output_settings()
        self._update_tab1_preview()

    def _update_tab1_preview(self):
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
                preview = "TICKER 公司名稱 data.xlsx"
        elif fmt == "ticker_only":
            preview = f"{ticker}.xlsx" if ticker else "TICKER.xlsx"
        else:  # custom
            custom = self.tab1_custom_var.get().strip() if self.tab1_custom_var else ""
            preview = f"{custom}.xlsx" if custom else "（請輸入檔名）"
        self.tab1_preview_label.config(text=f"預覽：{preview}")

    def _browse_output_dir(self):
        from tkinter import filedialog
        current = self.tab1_outdir_var.get().strip() if self.tab1_outdir_var else "output"
        initial = str(SCRIPT_DIR / current) if not os.path.isabs(current) else current
        folder = filedialog.askdirectory(title="選擇儲存位置", initialdir=initial)
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
        for w in self._tab2_inner.winfo_children():
            w.destroy()
        self.tab2_check_vars = {}
        watchlist = self.cfg.get("watchlist", [])
        if not watchlist:
            ttk.Label(self._tab2_inner, text="Watchlist 為空，請先在「管理 Watchlist」新增公司。",
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
            ttk.Label(hdr, text=gname, font=("", 11, "bold"), foreground="#333").pack(side="left")
            ttk.Button(hdr, text="全選", width=5,
                       command=lambda ts=tickers: self._select_group(ts, True)).pack(side="left", padx=(8, 2))
            ttk.Button(hdr, text="全不選", width=6,
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
        import copy
        self._ensure_groups(self.cfg)
        self._wl_draft = copy.deepcopy({
            "watchlist": self.cfg.get("watchlist", []),
            "groups":    self.cfg.get("groups",    []),
        })
        self._wl_group_collapsed = {}
        popup = tk.Toplevel(self.root)
        popup.title("管理 Watchlist")
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)
        popup.update()
        popup.attributes("-topmost", False)
        popup.bind("<Escape>", lambda e: popup.destroy())
        self._build_watchlist_popup(popup)

    def _build_watchlist_popup(self, popup: tk.Toplevel):
        pad = {"padx": 12, "pady": 4}
        popup.columnconfigure(0, weight=1)

        # ── Watchlist (scrollable, groups) ──────────────────────────
        list_frame = ttk.LabelFrame(popup, text=" 目前 Watchlist ", padding=6)
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

        ttk.Button(popup, text="＋ 新增群組",
                   command=lambda: self._wl_add_group(wl_inner)).grid(
            row=1, column=0, sticky="w", padx=12, pady=(0, 4))

        # ── Add company ────────────────────────────────────────────
        add_frame = ttk.LabelFrame(popup, text=" 新增公司 ", padding=6)
        add_frame.grid(row=2, column=0, sticky="ew", **pad)
        row_add = ttk.Frame(add_frame)
        row_add.grid(row=0, column=0, sticky="ew")
        ttk.Label(row_add, text="Ticker:").pack(side="left", padx=(0, 6))
        self.wl_add_var = tk.StringVar()
        wl_entry = ttk.Entry(row_add, textvariable=self.wl_add_var, width=10)
        wl_entry.pack(side="left", padx=(0, 8))
        wl_entry.bind("<Return>", lambda e: self._wl_lookup())
        ttk.Label(row_add, text="群組:").pack(side="left", padx=(8, 4))
        group_names = [g["name"] for g in self._get_groups_sorted(self._wl_draft)] or ["未分類"]
        self.wl_group_var = tk.StringVar(value=group_names[0])
        self.wl_group_cb = ttk.Combobox(row_add, textvariable=self.wl_group_var,
                                         values=group_names, width=12, state="readonly")
        self.wl_group_cb.pack(side="left", padx=(0, 8))
        ttk.Button(row_add, text="查詢", command=lambda: self._wl_lookup()).pack(side="left")
        self.wl_lookup_label = ttk.Label(add_frame, text="", foreground="gray")
        self.wl_lookup_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.wl_add_btn = ttk.Button(add_frame, text="加入 Watchlist", command=self._wl_add, state="disabled")
        self.wl_add_btn.grid(row=2, column=0, sticky="w", pady=4)
        self._wl_found_name = ""

        # ── Cache status ───────────────────────────────────────────
        cache_frame = ttk.Frame(popup)
        cache_frame.grid(row=3, column=0, sticky="ew", **pad)
        self.wl_cache_label = ttk.Label(cache_frame, text=self._wl_cache_status(), foreground="#555555")
        self.wl_cache_label.pack(side="left")
        ttk.Button(cache_frame, text="更新名稱庫（下載完整美股清單）",
                   command=self._wl_update_cache).pack(side="left", padx=10)

        # ── Save / discard ─────────────────────────────────────────
        btn_row = ttk.Frame(popup)
        btn_row.grid(row=4, column=0, pady=8)
        ttk.Button(btn_row, text="儲存關閉", width=12,
                   command=lambda: self._wl_save_close(popup)).pack(side="left", padx=6)
        ttk.Button(btn_row, text="放棄關閉", width=12,
                   command=popup.destroy).pack(side="left", padx=6)

    def _refresh_wl_popup_list(self, container):
        for w in container.winfo_children():
            w.destroy()
        watchlist = self._wl_draft.get("watchlist", [])
        wl_map = {w["ticker"]: w for w in watchlist}
        groups = self._get_groups_sorted(self._wl_draft)

        if not watchlist and not any(g["tickers"] for g in groups):
            ttk.Label(container, text="（空）", foreground="gray").pack(anchor="w")
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
            ttk.Button(hdr, text=f"{arrow} {gname}", width=16,
                       command=lambda g=gname, c=container: self._wl_toggle_group(g, c)).pack(side="left")
            ttk.Button(hdr, text="重新命名", width=8,
                       command=lambda g=gname, c=container: self._wl_rename_group(g, c)).pack(side="left", padx=(4, 0))
            if gname != "未分類":
                ttk.Button(hdr, text="刪除群組", width=8,
                           command=lambda g=gname, c=container: self._wl_delete_group(g, c)).pack(side="left", padx=(4, 0))

            if not is_collapsed:
                if not tickers:
                    ttk.Label(container, text="  （空群組）", foreground="gray").pack(anchor="w", padx=(20, 0))
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
                        path_text = "（預設）"
                        path_fg = "gray"
                    ttk.Label(row, text=path_text, foreground=path_fg, width=18).pack(side="left", padx=(4, 2))
                    ttk.Button(row, text="[x]", width=4,
                               command=lambda t=ticker, c=container: self._wl_remove(t, c)).pack(side="left")

        container.update_idletasks()
        if hasattr(self, "_wl_list_canvas"):
            self._wl_list_canvas.configure(scrollregion=self._wl_list_canvas.bbox("all"))

    def _wl_lookup(self):
        ticker = self.wl_add_var.get().strip().upper()
        if not ticker:
            self.wl_lookup_label.config(text="請輸入 Ticker", foreground="red")
            return
        self.wl_lookup_label.config(text="查詢中...", foreground="gray")
        self.wl_add_btn.config(state="disabled")
        self._wl_found_name = ""
        threading.Thread(target=lambda: self._wl_lookup_worker(ticker), daemon=True).start()

    def _wl_lookup_worker(self, ticker: str):
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
        folder = filedialog.askdirectory(title=f"選擇 {ticker} 的輸出資料夾")
        if not folder:
            return
        for item in self._wl_draft.get("watchlist", []):
            if item["ticker"] == ticker:
                item["output_dir"] = folder
                break
        self._refresh_wl_popup_list(container)

    def _wl_remove(self, ticker: str, container):
        self._wl_draft["watchlist"] = [w for w in self._wl_draft.get("watchlist", []) if w["ticker"] != ticker]
        for g in self._wl_draft.get("groups", []):
            if ticker in g["tickers"]:
                g["tickers"].remove(ticker)
        self._refresh_wl_popup_list(container)

    def _wl_add(self):
        ticker = self.wl_add_var.get().strip().upper()
        if not ticker or not self._wl_found_name:
            return
        if any(w["ticker"] == ticker for w in self._wl_draft.get("watchlist", [])):
            self.wl_lookup_label.config(text=f"{ticker} 已在 Watchlist 中", foreground="orange")
            return
        self._wl_draft.setdefault("watchlist", []).append({"ticker": ticker, "name": self._wl_found_name})
        target = self.wl_group_var.get() if self.wl_group_var else "未分類"
        grp = next((g for g in self._wl_draft.get("groups", []) if g["name"] == target), None)
        if grp is None:
            self._wl_draft.setdefault("groups", []).append({"name": target, "tickers": [ticker]})
        elif ticker not in grp["tickers"]:
            grp["tickers"].append(ticker)
        self.wl_add_var.set("")
        self.wl_lookup_label.config(text=f"✓ 已加入 {ticker} 到「{target}」", foreground="#1a7a34")
        self.wl_add_btn.config(state="disabled")
        self._wl_found_name = ""
        self._refresh_wl_popup_list(self._wl_list_container)

    def _wl_toggle_group(self, group_name: str, container):
        self._wl_group_collapsed[group_name] = not self._wl_group_collapsed.get(group_name, False)
        self._refresh_wl_popup_list(container)

    def _wl_add_group(self, container):
        from tkinter import simpledialog
        name = simpledialog.askstring("新增群組", "群組名稱：", parent=container.winfo_toplevel())
        if not name or not name.strip():
            return
        name = name.strip()
        if any(g["name"] == name for g in self._wl_draft.get("groups", [])):
            messagebox.showwarning("重複", f"群組「{name}」已存在", parent=container.winfo_toplevel())
            return
        self._wl_draft.setdefault("groups", []).append({"name": name, "tickers": []})
        self._refresh_group_dropdown()
        self._refresh_wl_popup_list(container)

    def _wl_rename_group(self, old_name: str, container):
        from tkinter import simpledialog
        new_name = simpledialog.askstring("重新命名", f"新名稱（原：{old_name}）：",
                                           parent=container.winfo_toplevel())
        if not new_name or not new_name.strip() or new_name.strip() == old_name:
            return
        new_name = new_name.strip()
        if any(g["name"] == new_name for g in self._wl_draft.get("groups", [])):
            messagebox.showwarning("重複", f"群組「{new_name}」已存在", parent=container.winfo_toplevel())
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
        grp = next((g for g in self._wl_draft.get("groups", []) if g["name"] == group_name), None)
        if not grp:
            return
        if grp["tickers"]:
            if not messagebox.askyesno("確認刪除",
                                        f"刪除「{group_name}」後，其中 {len(grp['tickers'])} 支股票將移至「未分類」。確定嗎？",
                                        parent=container.winfo_toplevel()):
                return
            uncategorized = next((g for g in self._wl_draft["groups"] if g["name"] == "未分類"), None)
            if uncategorized is None:
                self._wl_draft["groups"].append({"name": "未分類", "tickers": list(grp["tickers"])})
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
            names = [g["name"] for g in self._get_groups_sorted(self._wl_draft)] or ["未分類"]
            self.wl_group_cb["values"] = names
            if self.wl_group_var.get() not in names:
                self.wl_group_var.set(names[0])
        except tk.TclError:
            pass

    def _wl_save_close(self, popup: tk.Toplevel):
        self.cfg["watchlist"] = self._wl_draft.get("watchlist", [])
        self.cfg["groups"]    = self._wl_draft.get("groups",    [])
        save_config(self.cfg, CONFIG_PATH)
        self._refresh_tab2_list()
        popup.destroy()

    def _wl_update_cache(self):
        self.wl_cache_label.config(text="更新中...", foreground="gray")
        threading.Thread(target=self._wl_update_cache_worker, daemon=True).start()

    def _wl_update_cache_worker(self):
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
            cfg["groups"] = [{"name": "未分類", "tickers": tickers}] if tickers else []
        else:
            all_grouped = {t for g in cfg["groups"] for t in g["tickers"]}
            ungrouped = [w["ticker"] for w in cfg.get("watchlist", []) if w["ticker"] not in all_grouped]
            if ungrouped:
                for g in cfg["groups"]:
                    if g["name"] == "未分類":
                        g["tickers"].extend(ungrouped)
                        break
                else:
                    cfg["groups"].append({"name": "未分類", "tickers": ungrouped})

    def _get_groups_sorted(self, cfg: dict) -> list[dict]:
        """Return groups sorted A-Z, 未分類 always last."""
        groups = cfg.get("groups", [])
        known = sorted([g for g in groups if g["name"] != "未分類"], key=lambda g: g["name"])
        uncategorized = [g for g in groups if g["name"] == "未分類"]
        return known + uncategorized

    def _wl_cache_status(self) -> str:
        if CACHE_PATH.exists():
            try:
                with open(CACHE_PATH, encoding="utf-8") as f:
                    data = json.load(f)
                count = len(data.get("companies", {}))
                return f"已載入 {count:,} 間公司，上次更新：{data.get('last_updated', '未知')}"
            except (json.JSONDecodeError, OSError):
                return "名稱庫：檔案損毀"
        return "名稱庫：尚未建立（建議先點「更新名稱庫」下載完整清單）"

    # =========================================================
    # Advanced settings popup
    # =========================================================

    def _open_settings_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("進階設定")
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)
        popup.update()
        popup.attributes("-topmost", False)
        popup.bind("<Escape>", lambda e: popup.destroy())
        self._build_settings_popup(popup)

    def _build_settings_popup(self, popup: tk.Toplevel):
        pad = {"padx": 12, "pady": 4}

        # SEC Identity
        id_frame = ttk.LabelFrame(popup, text=" SEC EDGAR Identity ", padding=8)
        id_frame.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Label(id_frame, text="格式：姓名 空格 信箱（如 John Smith john@example.com）",
                  foreground="#555555", font=("", 10)).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(id_frame, text="Identity:").grid(row=1, column=0, sticky="w", pady=4)
        self.settings_identity_var = tk.StringVar(value=self.cfg.get("identity", ""))
        ttk.Entry(id_frame, textvariable=self.settings_identity_var, width=42).grid(row=1, column=1, sticky="ew", padx=(8, 0))

        # AI Config
        ai_frame = ttk.LabelFrame(popup, text=" AI 設定（Non-GAAP 功能需要，未設定不影響 GAAP）", padding=8)
        ai_frame.grid(row=1, column=0, sticky="ew", **pad)

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
        self.settings_key_toggle_btn = ttk.Button(key_row, text="顯示", width=5, command=self._toggle_key_show)
        self.settings_key_toggle_btn.pack(side="left")
        tk.Label(ai_frame, text="API Key 僅存於本機 config.json，請勿分享給他人。",
                 foreground="#555555", font=("", 10)).grid(row=3, column=0, columnspan=2, sticky="w")

        test_row = ttk.Frame(ai_frame)
        test_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(test_row, text="測試連線", command=self._test_ai_connection).pack(side="left")
        self.settings_test_label = ttk.Label(test_row, text="", foreground="gray")
        self.settings_test_label.pack(side="left", padx=10)

        # Fetch settings frame
        fetch_frame = ttk.LabelFrame(popup, text=" 抓取設定 ", padding=8)
        fetch_frame.grid(row=2, column=0, sticky="ew", **pad)
        fetch_frame.columnconfigure(2, weight=1)

        ttk.Label(fetch_frame, text="最多季報數量:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.settings_max_filings_var = tk.IntVar(value=self.cfg.get("max_filings", 80))
        max_spin = ttk.Spinbox(fetch_frame, from_=4, to=320, increment=4,
                               textvariable=self.settings_max_filings_var, width=6)
        max_spin.grid(row=0, column=1, sticky="w")
        ttk.Label(fetch_frame, text="筆（預設 80，約 20 年）", foreground="#555555").grid(
            row=0, column=2, sticky="w", padx=(4, 0))

        # Template mode
        ttk.Label(fetch_frame, text="著色模板:").grid(row=1, column=0, sticky="nw", pady=(10, 0))
        has_tpl = bool(self.cfg.get("template_path", ""))
        self.settings_template_mode_var = tk.StringVar(value="custom" if has_tpl else "default")
        self.settings_template_var = tk.StringVar(value=self.cfg.get("template_path", ""))

        tpl_frame = ttk.Frame(fetch_frame)
        tpl_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Radiobutton(tpl_frame, text="預設模板（Python 自動著色）",
                        variable=self.settings_template_mode_var, value="default",
                        command=self._on_template_mode_change).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Radiobutton(tpl_frame, text="自訂模板：",
                        variable=self.settings_template_mode_var, value="custom",
                        command=self._on_template_mode_change).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._tpl_entry = ttk.Entry(tpl_frame, textvariable=self.settings_template_var, width=24)
        self._tpl_entry.grid(row=1, column=1, sticky="ew", padx=(4, 4), pady=(4, 0))
        self._tpl_browse_btn = ttk.Button(tpl_frame, text="瀏覽", width=5,
                                           command=self._browse_template)
        self._tpl_browse_btn.grid(row=1, column=2, pady=(4, 0))
        self._on_template_mode_change()  # set initial enabled/disabled state

        # Buttons
        btn_row = ttk.Frame(popup)
        btn_row.grid(row=3, column=0, pady=10)
        ttk.Button(btn_row, text="儲存", command=lambda: self._save_settings(popup), width=10).pack(side="left", padx=6)
        ttk.Button(btn_row, text="取消", command=popup.destroy, width=10).pack(side="left", padx=6)

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
            title="選擇著色模板",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path and hasattr(self, "settings_template_var"):
            self.settings_template_var.set(path)

    def _on_provider_change(self, _event=None):
        provider = self.settings_provider_var.get()
        self.settings_model_var.set(PROVIDER_DEFAULTS.get(provider, ""))

    def _toggle_key_show(self):
        current = self.settings_key_entry.cget("show")
        new_show = "" if current else "•"
        self.settings_key_entry.config(show=new_show)
        if self.settings_key_toggle_btn:
            self.settings_key_toggle_btn.config(text="隱藏" if new_show == "" else "顯示")

    def _test_ai_connection(self):
        provider = self.settings_provider_var.get()
        model    = self.settings_model_var.get().strip()
        api_key  = self.settings_key_var.get().strip()
        if not api_key:
            self.settings_test_label.config(text="請輸入 API Key", foreground="red")
            return
        self.settings_test_label.config(text="測試中...", foreground="gray")
        threading.Thread(
            target=lambda: self._test_ai_worker(provider, model, api_key), daemon=True
        ).start()

    def _test_ai_worker(self, provider: str, model: str, api_key: str):
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
        save_config(self.cfg, CONFIG_PATH)
        popup.destroy()

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
                output_dir = SCRIPT_DIR / self.cfg.get("output_dir", "output")

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
        folder = self._last_output_folder or SCRIPT_DIR / self.cfg.get("output_dir", "output")
        if folder.exists():
            os.startfile(str(folder))

    # =========================================================
    # Run actions
    # =========================================================

    def _run_single(self):
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            messagebox.showerror("錯誤", "請輸入 Ticker")
            return
        if not self.fetch_gaap_var.get() and not self.fetch_nongaap_var.get():
            messagebox.showerror("錯誤", "請至少勾選 GAAP 或 Non-GAAP")
            return
        fetch_gaap    = self.fetch_gaap_var.get()
        fetch_nongaap = self.fetch_nongaap_var.get()
        if fetch_nongaap and not self.cfg["ai"].get("api_key"):
            messagebox.showwarning(
                "需要 API Key",
                "Non-GAAP 功能需要 AI API Key。\n請先至「進階設定」填入 API Key 後再執行。"
            )
            return
        max_filings = self.cfg.get("max_filings", 80)
        self._start_worker(lambda: self._worker_single(ticker, fetch_gaap, fetch_nongaap, max_filings))

    def _run_batch(self):
        selected = [t for t, v in self.tab2_check_vars.items() if v.get()]
        if not selected:
            messagebox.showerror("錯誤", "請至少勾選一間公司")
            return
        self._start_worker(lambda: self._worker_batch(selected))

    def _start_worker(self, target):
        if self.is_running:
            return
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self.btn_open_folder.pack_forget()
        self.progress_bar["value"] = 0
        self.progress_label.config(text="準備中...")
        self.is_running = True
        self.btn_run_single.config(state="disabled")
        self.btn_run_batch.config(state="disabled")
        threading.Thread(target=target, daemon=True).start()

    def _worker_single(self, ticker: str, fetch_gaap: bool, fetch_nongaap: bool, max_filings: int = 80):
        try:
            identity = self.cfg.get("identity", "")
            if not identity:
                self._log("[ERROR] 請先在進階設定填入 Identity（姓名 + 信箱）")
                self._done(False)
                return

            tables = []
            output_path = self._build_output_path(ticker)
            output_dir  = output_path.parent
            output_dir.mkdir(parents=True, exist_ok=True)

            total_steps = sum([fetch_gaap, fetch_nongaap]) + 1  # +1 for write
            step = 0

            if fetch_gaap:
                self._log(f"[{ticker}] 抓取 GAAP 財報中...")
                self._set_progress(step, total_steps, "抓取 GAAP...")
                gaap_tables = fetch_gaap_statements(ticker, identity, max_filings=max_filings)
                tables.extend(gaap_tables)
                self._log(f"[{ticker}] GAAP：取得 {len(gaap_tables)} 份財報")
                step += 1

            if fetch_nongaap:
                from fetcher_nongaap import fetch_nongaap_statements
                ai_config = self.cfg.get("ai", {})
                self._log(f"[{ticker}] 抓取 Non-GAAP 財報中...")
                self._set_progress(step, total_steps, "抓取 Non-GAAP...")

                def _ng_progress(current, total, label):
                    self._log(f"[{ticker}] {label}")
                    self._set_progress(current, total, label)

                ng_tables = fetch_nongaap_statements(
                    ticker, identity, ai_config,
                    output_dir=output_dir,
                    progress_cb=_ng_progress,
                )
                tables.extend(ng_tables)
                self._log(f"[{ticker}] Non-GAAP：{len(ng_tables)} 張 sheet")
                step += 1

            if not tables:
                self._log("[WARNING] 無資料可寫入")
                self._done(False)
                return

            self._log(f"[{ticker}] 寫入 Excel...")
            self._set_progress(step, total_steps, "寫入 Excel...")
            tpl = self.cfg.get("template_path", "") or None
            write_statements(tables, output_path, template_path=tpl)
            self._log(f"[{ticker}] 完成 → {output_path.name}")
            self._set_progress(total_steps, total_steps, "完成！")
            self.msg_queue.put(("last_output_folder", output_path.parent))
            self._done(True)

        except Exception as e:
            self._log(f"[ERROR] {e}")
            self._done(False)

    def _worker_batch(self, tickers: list[str]):
        total = len(tickers)
        identity = self.cfg.get("identity", "")
        if not identity:
            self._log("[ERROR] 請先在進階設定填入 Identity")
            self._done(False)
            return
        max_filings = self.cfg.get("max_filings", 80)

        for i, ticker in enumerate(tickers, 1):
            self._set_progress(i - 1, total, f"處理中：{ticker} ({i}/{total})")
            self._log(f"\n[{ticker}] 開始...")
            try:
                tables      = fetch_gaap_statements(ticker, identity, max_filings=max_filings)
                output_path = self._build_output_path(ticker)
                tpl = self.cfg.get("template_path", "") or None
                write_statements(tables, output_path, template_path=tpl)
                self._log(f"[{ticker}] 完成（{len(tables)} 份財報）")
            except Exception as e:
                self._log(f"[{ticker}] 錯誤：{e}")

        self._set_progress(total, total, f"完成：共處理 {total} 間公司")
        self.msg_queue.put(("last_output_folder", self._build_output_path(tickers[-1]).parent))
        self._done(True)

    # =========================================================
    # Thread-safe queue helpers
    # =========================================================

    def _log(self, msg: str):
        self.msg_queue.put(("log", msg))

    def _init_log(self, msg: str):
        self.log_text.config(state="normal")
        self.log_text.insert("1.0", msg + "\n")
        self.log_text.config(state="disabled")

    def _set_progress(self, current: int, total: int, label: str):
        self.msg_queue.put(("progress", (current, total, label)))

    def _done(self, success: bool):
        self.msg_queue.put(("done", success))

    def _poll_queue(self):
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
                        self.progress_label.config(text="完成！")
                    else:
                        self.progress_label.config(text="發生錯誤，請查看上方記錄")

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
                            self.tab1_name_label.config(text="　查無此 Ticker，請確認後再試", foreground="orange")
                        self._update_tab1_preview()
                    if self.btn_confirm_company:
                        self.btn_confirm_company.config(state="normal")

                elif msg_type == "wl_lookup_result":
                    status = data[0]
                    if status == "ok":
                        _, ticker, name = data
                        self._wl_found_name = name
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=f"查到：{name}", foreground="#1a7a34")
                        if self.wl_add_btn:
                            self.wl_add_btn.config(state="normal")
                        self._wl_add()
                    else:
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=f"查詢失敗：{data[1]}", foreground="red")

                elif msg_type == "wl_cache_updated":
                    update_date, count = data
                    if self.wl_cache_label:
                        self.wl_cache_label.config(
                            text=f"已載入 {count:,} 間公司，上次更新：{update_date}",
                            foreground="gray"
                        )

                elif msg_type == "wl_cache_update_error":
                    if self.wl_cache_label:
                        self.wl_cache_label.config(text=f"更新失敗：{data}", foreground="red")

                elif msg_type == "last_output_folder":
                    self._last_output_folder = data

                elif msg_type == "ai_test_result":
                    ok, err = data
                    if self.settings_test_label:
                        if ok == "ok":
                            self.settings_test_label.config(text="連線成功！", foreground="#1a7a34")
                        else:
                            self.settings_test_label.config(text=f"失敗：{str(err)[:60]}", foreground="red")

        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# =========================================================
# Entry point
# =========================================================

def main():
    show_cth_banner()
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.update()
    root.attributes("-topmost", False)
    SECFetcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
