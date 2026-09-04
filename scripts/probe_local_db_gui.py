# -*- coding: utf-8 -*-
"""Tk 探針：把「更新本地庫」那條**執行路徑**真的跑一次（TODO J1-J4，不連網）。

專案現況是「GUI 對話框沿用 Tk 探針手動驗，不寫自動測試」，這支就是那個探針。
驗的不是「畫得出來」，是動起來之後的事：

    按鈕鎖 → 背景執行緒 → msg_queue → db_done → 按鈕解鎖 → log 內容

外加兩道前置檢查（名單空的、identity 沒填）與 J4 的版本提醒對話框。
`local_db.update_local_db` 整個換掉，所以**完全不打 SEC**，幾秒跑完。

    ./venv/Scripts/python.exe scripts/probe_local_db_gui.py

回傳碼 0＝24 項全過，1＝有項目沒過（會列出是哪幾項）。改 `main.py` 的
`_start_local_db_update` / `_local_db_worker` / `_poll_queue` 的 `db_done`
分支、或 `_warn_if_edgartools_changed` 之後跑一次。

⚠ 會**暫時**改動 `app.cfg` 的 `identity` 與 `local_db_tickers`（記憶體裡的那份），
結尾會還原，而且全程不呼叫 `save_config()`，不會動到你的 config.json。
"""
import sys, threading, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import tkinter as tk
from tkinter import messagebox

import main as m
import local_db

FAILURES = []


def check(name, cond, extra=""):
    print(("  OK   " if cond else "  FAIL ") + name + ("" if cond else f"  <- {extra}"))
    if not cond:
        FAILURES.append(name)


def pump(root, seconds=3.0, until=None):
    """跑 Tk 事件迴圈（含 _poll_queue 的 100ms after），直到條件成立或逾時。"""
    end = time.time() + seconds
    while time.time() < end:
        root.update()
        if until is not None and until():
            return True
        time.sleep(0.02)
    return until is None


root = tk.Tk()
root.withdraw()
app = m.SECFetcherApp(root)
orig_cfg_list = list(app.cfg.get("local_db_tickers") or [])
orig_identity = app.cfg.get("identity", "")

# ── 1. 前置檢查：名單是空的 → 只跳提示，不開執行緒 ────────────────────────
print("\n[1] 名單空的時候不該發動")
shown = []
m.messagebox.showinfo = lambda *a, **k: shown.append(("info", a))
m.messagebox.showerror = lambda *a, **k: shown.append(("error", a))
m.messagebox.askyesno = lambda *a, **k: True

local_db.set_update_list(app.cfg, [])
app._start_local_db_update()
check("名單空 → 跳 info、沒開始跑", shown and shown[-1][0] == "info" and not app.is_running,
      f"shown={shown} is_running={app.is_running}")

# ── 2. 前置檢查：identity 沒填 ────────────────────────────────────────────
print("\n[2] identity 沒填的時候不該發動")
shown.clear()
local_db.set_update_list(app.cfg, ["AAPL", "NVDA"])
app.cfg["identity"] = ""
app._start_local_db_update()
check("identity 空 → 跳 error、沒開始跑", shown and shown[-1][0] == "error" and not app.is_running,
      f"shown={shown} is_running={app.is_running}")

# ── 3. 正常跑一趟（update_local_db 換成假的，不連網）────────────────────
print("\n[3] 正常跑一趟：按鈕鎖 → 背景 → db_done → 解鎖")
app.cfg["identity"] = "Probe probe@example.com"
calls = {}
# 卡住背景執行緒，才觀察得到「跑到一半」的按鈕狀態——不卡的話假 worker
# 0.1 秒就跑完了，檢查時早就解鎖了（第一版探針就是這樣誤判成 FAIL）
release = threading.Event()


def fake_update(tickers, identity, progress=None, **kw):
    release.wait(10)
    calls["tickers"] = list(tickers)
    calls["identity"] = identity
    total = len(tickers)
    for i, tk_ in enumerate(tickers):
        progress({"event": "ticker_start", "ticker": tk_, "index": i, "total": total})
        time.sleep(0.05)
        progress({"event": "ticker_done", "ticker": tk_,
                  "status": "skipped" if i == 0 else "updated",
                  "index": i, "total": total, "new_filings": i * 3, "gaps": i})
    return local_db.UpdateReport(results=[
        local_db.TickerResult("AAPL", "skipped"),
        local_db.TickerResult("NVDA", "updated", new_filings=3, gaps=1),
    ])


local_db.update_local_db = fake_update
app._start_local_db_update()
root.update()
check("開跑後 is_running=True", app.is_running is True)
check("開跑後清除鈕被鎖", str(app._cache_clear_all_btn.cget("state")) == "disabled",
      app._cache_clear_all_btn.cget("state"))
check("開跑後「更新本地庫」自己也被鎖", str(app._localdb_run_btn.cget("state")) == "disabled",
      app._localdb_run_btn.cget("state"))
check("開跑後 Tab1 抓取鈕被鎖", str(app.btn_run_single.cget("state")) == "disabled",
      app.btn_run_single.cget("state"))
release.set()

done = pump(root, seconds=8.0, until=lambda: not app.is_running)
check("跑完 is_running 回 False（db_done 有被消化）", done and not app.is_running)
check("跑完清除鈕解鎖", str(app._cache_clear_all_btn.cget("state")) == "normal",
      app._cache_clear_all_btn.cget("state"))
check("跑完 Tab1 抓取鈕解鎖", str(app.btn_run_single.cget("state")) == "normal",
      app.btn_run_single.cget("state"))
check("ticker 有正確傳進去", calls.get("tickers") == ["AAPL", "NVDA"], calls)

log = app.log_text.get("1.0", "end")
check("log 有起始行", "開始更新本地庫" in log, log[:200])
check("log 有逐家的結果", "[AAPL]" in log and "[NVDA]" in log, log[:400])
check("log 有完成行含耗時", "更新本地庫完成" in log, log[-300:])
check("log 有缺漏提醒（D11）", "建議之後單獨重跑" in log and "NVDA" in log, log[-300:])
check("進度標籤是完成不是錯誤", app.progress_label.cget("text") == m.t("gui.status.done"),
      app.progress_label.cget("text"))

# 這條是刻意的：更新本地庫不產 Excel，「開啟輸出資料夾」不該冒出來
check("「開啟輸出資料夾」沒有被顯示出來",
      not app.btn_open_folder.winfo_ismapped(), "db_done 不該走 done 那條路")

# ── 4. 例外時走 FAIL 路徑，按鈕一樣要解鎖 ────────────────────────────────
print("\n[4] update_local_db 拋例外時也要解鎖")


def boom(*a, **k):
    raise RuntimeError("probe boom")


local_db.update_local_db = boom
app._start_local_db_update()
root.update()
pump(root, seconds=8.0, until=lambda: not app.is_running)
check("例外後 is_running 回 False", not app.is_running)
check("例外後按鈕解鎖", str(app._localdb_run_btn.cget("state")) == "normal")
check("例外後進度標籤是錯誤",
      app.progress_label.cget("text") == m.t("gui.status.error_see_log"),
      app.progress_label.cget("text"))
check("log 有失敗行", "更新本地庫失敗" in app.log_text.get("1.0", "end"))

# ── 5. J4：版本不符的提醒對話框 ──────────────────────────────────────────
print("\n[5] J4 版本提醒對話框")
warned = []
m.messagebox.showwarning = lambda *a, **k: warned.append(a)

local_db.stale_cache_summary = lambda: {
    "current": "5.31.0", "companies": ["AAPL", "NVDA"], "n_companies": 2,
    "n_filings": 132, "size_bytes": 11 * 1024 * 1024,
    "old_versions": ["5.29.0"], "estimated_seconds": 132 * 2.8}
m._warn_if_edgartools_changed(root)
check("版本不符會跳警告", len(warned) == 1, warned)
body = warned[0][1] if warned else ""
check("訊息有講幾家幾份", "2" in body and "132" in body, body[:200])
check("訊息有附回退指令", "pip install edgartools==5.29.0" in body, body[-200:])
check("訊息沒有給「照用舊快取」的選項", "照用" not in body, body[:400])

warned.clear()
local_db.stale_cache_summary = lambda: {
    "current": "5.29.0", "companies": [], "n_companies": 0, "n_filings": 0,
    "size_bytes": 0, "old_versions": [], "estimated_seconds": 0}
m._warn_if_edgartools_changed(root)
check("版本相符時不該跳任何東西", warned == [], warned)

local_db.stale_cache_summary = lambda: (_ for _ in ()).throw(OSError("nope"))
m._warn_if_edgartools_changed(root)
check("偵測本身炸掉時不該擋住啟動", True)

# ── 收尾：還原被探針改掉的設定 ───────────────────────────────────────────
app.cfg["identity"] = orig_identity
local_db.set_update_list(app.cfg, orig_cfg_list)
root.destroy()

print("\n" + ("PROBE FAILED: " + ", ".join(FAILURES) if FAILURES else "PROBE OK（全部通過）"))
sys.exit(1 if FAILURES else 0)
