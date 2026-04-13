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
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

from config import load_config, save_config
from excel_writer import write_statements
from fetcher_gaap import fetch_gaap_statements

SCRIPT_DIR  = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
CACHE_PATH  = SCRIPT_DIR / "company_cache.json"

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
        self.root.resizable(False, False)

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
        self.settings_identity_var = None
        self.settings_provider_var = None
        self.settings_model_var = None
        self.settings_key_var = None
        self.settings_key_entry = None
        self.settings_outdir_var = None
        self.settings_test_label = None

        self._build_ui()
        self._poll_queue()

    # =========================================================
    # UI Construction
    # =========================================================

    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

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
        frame_log.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 4))
        self.progress_label = ttk.Label(frame_log, text="等待開始...")
        self.progress_label.pack(anchor="w")
        self.progress_bar = ttk.Progressbar(frame_log, mode="determinate", length=440)
        self.progress_bar.pack(fill="x", pady=(4, 8))
        self.log_text = scrolledtext.ScrolledText(
            frame_log, width=60, height=10, state="disabled", font=("Consolas", 9)
        )
        self.log_text.pack(fill="x")

        # Open folder button (shown after completion)
        frame_output = tk.Frame(self.root)
        frame_output.grid(row=3, column=0, pady=(0, 12))
        self.btn_open_folder = ttk.Button(
            frame_output, text="開啟輸出資料夾", command=self._open_output_folder
        )

        self.root.columnconfigure(0, weight=1)
        self._init_log("請輸入 Ticker 或選擇 Watchlist 後按執行。")

    def _build_tab1(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  單一公司  ")

        row_ticker = ttk.Frame(tab)
        row_ticker.grid(row=0, column=0, sticky="ew", pady=4)
        ttk.Label(row_ticker, text="Ticker:").pack(side="left", padx=(0, 8))
        self.ticker_var = tk.StringVar()
        self.ticker_entry = ttk.Entry(row_ticker, textvariable=self.ticker_var, width=18, foreground="grey")
        self.ticker_entry.pack(side="left")
        self.ticker_var.set(self.TICKER_PH)
        self.ticker_entry.bind("<FocusIn>",  lambda e: self._ph_in(self.ticker_entry, self.ticker_var, self.TICKER_PH))
        self.ticker_entry.bind("<FocusOut>", lambda e: self._ph_out(self.ticker_entry, self.ticker_var, self.TICKER_PH))

        row_type = ttk.Frame(tab)
        row_type.grid(row=1, column=0, sticky="ew", pady=4)
        self.fetch_gaap_var    = tk.BooleanVar(value=True)
        self.fetch_nongaap_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row_type, text="GAAP 財報",               variable=self.fetch_gaap_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(row_type, text="Non-GAAP（需設定 AI API）", variable=self.fetch_nongaap_var).pack(side="left")

        self.btn_run_single = ttk.Button(tab, text="▶  執行", command=self._run_single, width=16)
        self.btn_run_single.grid(row=2, column=0, pady=(8, 4))

    def _build_tab2(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  批量更新  ")

        self.tab2_list_frame = ttk.LabelFrame(tab, text=" Watchlist ", padding=6)
        self.tab2_list_frame.grid(row=0, column=0, sticky="ew", pady=4)
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

    def _ph_out(self, entry, var, placeholder):
        if not var.get().strip():
            var.set(placeholder)
            entry.configure(foreground="grey")

    def _get_ph_value(self, var, placeholder) -> str:
        v = var.get().strip()
        return "" if v == placeholder else v

    # =========================================================
    # Tab 2 watchlist list
    # =========================================================

    def _refresh_tab2_list(self):
        for w in self.tab2_list_frame.winfo_children():
            w.destroy()
        self.tab2_check_vars = {}
        watchlist = self.cfg.get("watchlist", [])
        if not watchlist:
            ttk.Label(self.tab2_list_frame, text="Watchlist 為空，請先在「管理 Watchlist」新增公司。",
                      foreground="gray").pack(anchor="w")
            return
        for item in watchlist:
            var = tk.BooleanVar(value=True)
            self.tab2_check_vars[item["ticker"]] = var
            ttk.Checkbutton(
                self.tab2_list_frame,
                text=f'{item["ticker"]:6}  {item.get("name", "")}',
                variable=var,
            ).pack(anchor="w")

    def _select_all(self):
        for v in self.tab2_check_vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self.tab2_check_vars.values():
            v.set(False)

    # =========================================================
    # Watchlist popup
    # =========================================================

    def _open_watchlist_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("管理 Watchlist")
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)
        popup.update()
        popup.attributes("-topmost", False)
        self._build_watchlist_popup(popup)

    def _build_watchlist_popup(self, popup: tk.Toplevel):
        pad = {"padx": 12, "pady": 4}

        list_frame = ttk.LabelFrame(popup, text=" 目前 Watchlist ", padding=6)
        list_frame.grid(row=0, column=0, sticky="ew", **pad)
        self._wl_list_container = list_frame
        self._refresh_wl_popup_list(list_frame)

        add_frame = ttk.LabelFrame(popup, text=" 新增公司 ", padding=6)
        add_frame.grid(row=1, column=0, sticky="ew", **pad)
        row_add = ttk.Frame(add_frame)
        row_add.grid(row=0, column=0, sticky="ew")
        ttk.Label(row_add, text="Ticker:").pack(side="left", padx=(0, 6))
        self.wl_add_var = tk.StringVar()
        ttk.Entry(row_add, textvariable=self.wl_add_var, width=10).pack(side="left", padx=(0, 8))
        ttk.Button(row_add, text="查詢", command=lambda: self._wl_lookup()).pack(side="left")
        self.wl_lookup_label = ttk.Label(add_frame, text="", foreground="gray")
        self.wl_lookup_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.wl_add_btn = ttk.Button(add_frame, text="加入 Watchlist", command=self._wl_add, state="disabled")
        self.wl_add_btn.grid(row=2, column=0, sticky="w", pady=4)
        self._wl_found_name = ""

        cache_frame = ttk.Frame(popup)
        cache_frame.grid(row=2, column=0, sticky="ew", **pad)
        self.wl_cache_label = ttk.Label(cache_frame, text=self._wl_cache_status(), foreground="gray")
        self.wl_cache_label.pack(side="left")
        ttk.Button(cache_frame, text="更新名稱庫", command=self._wl_update_cache).pack(side="left", padx=10)

        ttk.Button(popup, text="關閉", command=popup.destroy, width=10).grid(row=3, column=0, pady=8)

    def _refresh_wl_popup_list(self, container):
        for w in container.winfo_children():
            w.destroy()
        watchlist = self.cfg.get("watchlist", [])
        if not watchlist:
            ttk.Label(container, text="（空）", foreground="gray").pack(anchor="w")
            return
        for item in watchlist:
            row = ttk.Frame(container)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f'{item["ticker"]:6} {item.get("name", "")}', width=36).pack(side="left")
            ticker = item["ticker"]
            ttk.Button(row, text="[x]", width=4,
                       command=lambda t=ticker, c=container: self._wl_remove(t, c)).pack(side="left")

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
            with open(CACHE_PATH, encoding="utf-8") as f:
                cache = json.load(f).get("companies", {})
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

    def _wl_remove(self, ticker: str, container):
        self.cfg["watchlist"] = [w for w in self.cfg["watchlist"] if w["ticker"] != ticker]
        save_config(self.cfg, CONFIG_PATH)
        self._refresh_wl_popup_list(container)
        self._refresh_tab2_list()

    def _wl_add(self):
        ticker = self.wl_add_var.get().strip().upper()
        if not ticker or not self._wl_found_name:
            return
        if any(w["ticker"] == ticker for w in self.cfg["watchlist"]):
            self.wl_lookup_label.config(text=f"{ticker} 已在 Watchlist 中", foreground="orange")
            return
        self.cfg["watchlist"].append({"ticker": ticker, "name": self._wl_found_name})
        save_config(self.cfg, CONFIG_PATH)
        self.wl_add_var.set("")
        self.wl_lookup_label.config(text="", foreground="gray")
        self.wl_add_btn.config(state="disabled")
        self._wl_found_name = ""
        self._refresh_wl_popup_list(self._wl_list_container)
        self._refresh_tab2_list()

    def _wl_update_cache(self):
        self.wl_cache_label.config(text="更新中...", foreground="gray")
        threading.Thread(target=self._wl_update_cache_worker, daemon=True).start()

    def _wl_update_cache_worker(self):
        from edgar import Company, set_identity
        identity = self.cfg.get("identity", "SEC Tool sec@example.com")
        set_identity(identity)
        companies: dict[str, str] = {}
        for item in self.cfg.get("watchlist", []):
            ticker = item["ticker"]
            try:
                c = Company(ticker)
                companies[ticker] = c.name or ticker
            except Exception:
                companies[ticker] = item.get("name", ticker)
        cache_data = {"last_updated": str(date.today()), "companies": companies}
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
        self.msg_queue.put(("wl_cache_updated", str(date.today())))

    def _wl_cache_status(self) -> str:
        if CACHE_PATH.exists():
            with open(CACHE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            return f"上次更新：{data.get('last_updated', '未知')}"
        return "名稱庫：尚未建立"

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
        self._build_settings_popup(popup)

    def _build_settings_popup(self, popup: tk.Toplevel):
        pad = {"padx": 12, "pady": 4}

        # SEC Identity
        id_frame = ttk.LabelFrame(popup, text=" SEC EDGAR Identity ", padding=8)
        id_frame.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Label(id_frame, text="格式：姓名 空格 信箱（如 John Smith john@example.com）",
                  foreground="gray", font=("", 8)).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(id_frame, text="Identity:").grid(row=1, column=0, sticky="w", pady=4)
        self.settings_identity_var = tk.StringVar(value=self.cfg.get("identity", ""))
        ttk.Entry(id_frame, textvariable=self.settings_identity_var, width=42).grid(row=1, column=1, sticky="ew", padx=(8, 0))

        # AI Config
        ai_frame = ttk.LabelFrame(popup, text=" AI 設定（Non-GAAP 功能需要）", padding=8)
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
        ttk.Button(key_row, text="顯示", width=5, command=self._toggle_key_show).pack(side="left")
        tk.Label(ai_frame, text="API Key 僅存於本機 config.json，請勿分享給他人。",
                 foreground="gray", font=("", 8)).grid(row=3, column=0, columnspan=2, sticky="w")

        test_row = ttk.Frame(ai_frame)
        test_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(test_row, text="測試連線", command=self._test_ai_connection).pack(side="left")
        self.settings_test_label = ttk.Label(test_row, text="", foreground="gray")
        self.settings_test_label.pack(side="left", padx=10)

        # Output dir
        out_frame = ttk.LabelFrame(popup, text=" 輸出資料夾 ", padding=8)
        out_frame.grid(row=2, column=0, sticky="ew", **pad)
        ttk.Label(out_frame, text="路徑:").grid(row=0, column=0, sticky="w")
        self.settings_outdir_var = tk.StringVar(value=self.cfg.get("output_dir", "output"))
        ttk.Entry(out_frame, textvariable=self.settings_outdir_var, width=36).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        # Buttons
        btn_row = ttk.Frame(popup)
        btn_row.grid(row=3, column=0, pady=10)
        ttk.Button(btn_row, text="儲存", command=lambda: self._save_settings(popup), width=10).pack(side="left", padx=6)
        ttk.Button(btn_row, text="取消", command=popup.destroy, width=10).pack(side="left", padx=6)

    def _on_provider_change(self, _event=None):
        provider = self.settings_provider_var.get()
        self.settings_model_var.set(PROVIDER_DEFAULTS.get(provider, ""))

    def _toggle_key_show(self):
        current = self.settings_key_entry.cget("show")
        self.settings_key_entry.config(show="" if current else "•")

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
        self.cfg["output_dir"]     = self.settings_outdir_var.get().strip() or "output"
        self.cfg["ai"]["provider"] = self.settings_provider_var.get()
        self.cfg["ai"]["model"]    = self.settings_model_var.get().strip()
        self.cfg["ai"]["api_key"]  = self.settings_key_var.get().strip()
        save_config(self.cfg, CONFIG_PATH)
        popup.destroy()

    # =========================================================
    # Open output folder
    # =========================================================

    def _open_output_folder(self):
        folder = SCRIPT_DIR / self.cfg.get("output_dir", "output")
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
        self._start_worker(lambda: self._worker_single(ticker))

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

    def _worker_single(self, ticker: str):
        try:
            identity = self.cfg.get("identity", "")
            if not identity:
                self._log("[ERROR] 請先在進階設定填入 Identity（姓名 + 信箱）")
                self._done(False)
                return

            tables = []

            if self.fetch_gaap_var.get():
                self._log(f"[{ticker}] 抓取 GAAP 財報中...")
                self._set_progress(0, 2, "抓取 GAAP...")
                gaap_tables = fetch_gaap_statements(ticker, identity)
                tables.extend(gaap_tables)
                self._log(f"[{ticker}] GAAP：取得 {len(gaap_tables)} 份財報")

            if self.fetch_nongaap_var.get():
                self._log(f"[{ticker}] Non-GAAP 功能尚未實作（Phase 2）")

            if not tables:
                self._log("[WARNING] 無資料可寫入")
                self._done(False)
                return

            output_dir  = SCRIPT_DIR / self.cfg.get("output_dir", "output")
            output_path = output_dir / f"{ticker}.xlsx"
            self._log(f"[{ticker}] 寫入 Excel...")
            write_statements(tables, output_path)
            self._log(f"[{ticker}] 完成 → {output_path.name}")
            self._set_progress(2, 2, "完成！")
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

        for i, ticker in enumerate(tickers, 1):
            self._set_progress(i - 1, total, f"處理中：{ticker} ({i}/{total})")
            self._log(f"\n[{ticker}] 開始...")
            try:
                tables      = fetch_gaap_statements(ticker, identity)
                output_dir  = SCRIPT_DIR / self.cfg.get("output_dir", "output")
                output_path = output_dir / f"{ticker}.xlsx"
                write_statements(tables, output_path)
                self._log(f"[{ticker}] 完成（{len(tables)} 份財報）")
            except Exception as e:
                self._log(f"[{ticker}] 錯誤：{e}")

        self._set_progress(total, total, f"完成：共處理 {total} 間公司")
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

                elif msg_type == "wl_lookup_result":
                    status = data[0]
                    if status == "ok":
                        _, ticker, name = data
                        self._wl_found_name = name
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=f"查到：{name}", foreground="#2ecc71")
                        if self.wl_add_btn:
                            self.wl_add_btn.config(state="normal")
                    else:
                        if self.wl_lookup_label:
                            self.wl_lookup_label.config(text=f"查詢失敗：{data[1]}", foreground="red")

                elif msg_type == "wl_cache_updated":
                    if self.wl_cache_label:
                        self.wl_cache_label.config(text=f"上次更新：{data}", foreground="gray")

                elif msg_type == "ai_test_result":
                    ok, err = data
                    if self.settings_test_label:
                        if ok == "ok":
                            self.settings_test_label.config(text="連線成功！", foreground="#2ecc71")
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
