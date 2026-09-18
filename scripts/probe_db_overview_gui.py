# -*- coding: utf-8 -*-
"""Tk 探針：資料庫總覽分頁的**執行路徑**（TODO J7，不連網）。

專案現況是「GUI 用 Tk 探針手動驗，純函式寫自動測試」，這支是總覽分頁那支。
`db_overview_cells` / `db_overview_csv` / `overview_rows` 的邏輯在
`tests/test_local_db.py` 已經測過，這裡驗的是 Treeview 真的動起來之後的事：

    切頁觸發重讀 → Treeview 有列 → 點標題排序 → 搜尋框篩選 → 匯出 CSV

⚠ 特別釘一條：**`_on_tab_changed` 用分頁身分比對、不是寫死 index**。原本那行
是 `== 3`，插入這個新分頁時就會靜默壞掉（3 從「進階設定」變成「資料庫總覽」，
快取清單再也不刷新，而且不報錯）。

    ./venv/Scripts/python.exe scripts/probe_db_overview_gui.py

回傳碼 0＝全過，1＝有項目沒過。跑之前會在 tmp 造一個假資料庫
（`SEC_LOCAL_DB_ROOT` 導過去），**不會碰到你真正的 local_db/**。
"""
import json
import os
import sys
import tempfile
import time
from pathlib import Path

# ⚠ 一定要在 import main 之前設好——`cache_root()` 每次呼叫重讀環境變數，
# 但 app 建構時就會掃一次資料庫，太晚設會掃到真的那份。
_TMP = Path(tempfile.mkdtemp(prefix="probe_db_overview_"))
os.environ["SEC_LOCAL_DB_ROOT"] = str(_TMP)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import tkinter as tk          # noqa: E402

import pandas as pd           # noqa: E402
import filing_cache           # noqa: E402
import local_db               # noqa: E402
import main as m              # noqa: E402

FAILURES = []


def check(name, cond, extra=""):
    print(("  OK   " if cond else "  FAIL ") + name + ("" if cond else f"  <- {extra}"))
    if not cond:
        FAILURES.append(name)


def seed(ticker, n_filings, first_year):
    """造一家假公司。走真的 `save_filing` 路徑太重，直接寫檔就好——
    `_meta.json` 讓 `load_meta()` 自己生，跟真實情況一致。"""
    d = filing_cache.ticker_dir(ticker)
    d.mkdir(parents=True, exist_ok=True)
    for i in range(n_filings):
        acc = f"0000320193-{i // 100 % 100:02d}-{i % 100:06d}"
        (d / f"{acc}.json").write_text(json.dumps({
            "schema_version": filing_cache.SCHEMA_VERSION,
            "accession_no": acc,
            "form": "10-K" if i % 4 == 0 else "10-Q",
            "filing_date": f"{first_year + i // 4}-05-01",
            "cached_at": "2026-09-05T00:00:00+08:00",
            "cik": 320193, "edgartools_version": "5.29.0",
            "has_financials": True,
            # 真實欄名：期末日 03-29，跟收件日 05-01 差一整季——驗得出
            # 「顯示的是期間不是收件日」
            "dataframes": {
                "income_statement": filing_cache.df_to_payload(pd.DataFrame({
                    "concept": ["Revenue"],
                    f"{first_year + i // 4}-03-29 (Q1)": [1.0],
                    f"{first_year + i // 4 - 1}-03-30 (Q1)": [0.9],
                })),
                "balance_sheet": None, "cashflow_statement": None,
            },
        }, ensure_ascii=False), encoding="utf-8")
    local_db.load_meta(ticker)


# AAPL 長歷史、BLK 短歷史（模擬 J6 換過 CIK）、NVDA 中等
seed("AAPL", 40, 2008)
seed("BLK", 4, 2024)
seed("NVDA", 12, 2020)

root = tk.Tk()
root.withdraw()
app = m.SECFetcherApp(root)
root.update()

# ── 1. 分頁存在，而且切過去會觸發重讀 ────────────────────────────────────
print("\n[1] 分頁建得出來、切過去會重讀")
tabs = [app.notebook.tab(i, "text") for i in range(app.notebook.index("end"))]
check("多了一個分頁", len(tabs) == 5, f"tabs={tabs}")
check("分頁名稱是資料庫總覽", any("資料庫總覽" in x for x in tabs), f"tabs={tabs}")

app._db_rows = []                       # 清掉，看切頁是不是真的會重讀
app.notebook.select(app._tab_database)
root.update()
check("切過去後有讀到 3 家", len(app._db_rows) == 3, f"rows={len(app._db_rows)}")

# ── 2. `_on_tab_changed` 用身分比對，不是寫死 index ──────────────────────
print("\n[2] 切回進階設定，快取面板仍會刷新（原本寫死 index==3 的那行）")
app._cache_panel_refreshed = False
_orig_refresh = app._refresh_cache_panel


def _spy():
    app._cache_panel_refreshed = True
    _orig_refresh()


app._refresh_cache_panel = _spy
app.notebook.select(app._tab_settings)
root.update()
check("切到進階設定有刷新快取面板", app._cache_panel_refreshed,
      "插入新分頁後 index 位移，寫死 index 的話這裡會靜默失敗")
app._refresh_cache_panel = _orig_refresh
app.notebook.select(app._tab_database)
root.update()

# ── 3. Treeview 真的有列，而且內容對得上 ────────────────────────────────
print("\n[3] Treeview 的內容")
items = app._db_tree.get_children()
check("三家都在表上", len(items) == 3, f"items={len(items)}")
first = app._db_tree.item(items[0], "values")
check("預設照代號升冪，第一列是 AAPL", first[0] == "AAPL", f"first={first}")
check("份數欄顯示 40", first[1] == "40", f"first={first}")
check("申報日期是範圍格式", "~" in first[2], f"first={first}")
check("年數算得出來", first[3] != "—", f"first={first}")

blk = [app._db_tree.item(i, "values") for i in items
       if app._db_tree.item(i, "values")[0] == "BLK"][0]
check("BLK 年數明顯比 AAPL 短（J6 的情況看得出來）",
      float(blk[3].split()[0]) < float(first[3].split()[0]),
      f"BLK={blk[3]} AAPL={first[3]}")

# ── 4. 點欄位標題排序 ───────────────────────────────────────────────────
print("\n[4] 點標題排序")
app._sort_db_overview("years")
root.update()
years_asc = [app._db_tree.item(i, "values")[0] for i in app._db_tree.get_children()]
check("年數升冪，最短的 BLK 在第一個", years_asc[0] == "BLK", f"order={years_asc}")

app._sort_db_overview("years")          # 同一欄再點一次 → 反向
root.update()
years_desc = [app._db_tree.item(i, "values")[0] for i in app._db_tree.get_children()]
check("再點一次變降冪", years_desc == list(reversed(years_asc)),
      f"asc={years_asc} desc={years_desc}")

app._sort_db_overview("ticker")
root.update()

# ── 5. 搜尋框即時篩選 ───────────────────────────────────────────────────
print("\n[5] 搜尋")
app._db_search_var.set("nv")
root.update()
shown = [app._db_tree.item(i, "values")[0] for i in app._db_tree.get_children()]
check("小寫 nv 找得到 NVDA", shown == ["NVDA"], f"shown={shown}")
check("摘要列有標示篩出幾家", "1" in app._db_summary_label.cget("text"),
      app._db_summary_label.cget("text"))

app._db_search_var.set("")
root.update()
check("清空搜尋後三家都回來",
      len(app._db_tree.get_children()) == 3,
      f"items={len(app._db_tree.get_children())}")

# ── 6. 匯出 CSV（檔案對話框換掉，不跳視窗）───────────────────────────────
print("\n[6] 匯出 CSV")
out = _TMP / "overview.csv"
import tkinter.filedialog as fd        # noqa: E402
fd.asksaveasfilename = lambda **kw: str(out)

app._db_search_var.set("a")            # 只匯出篩選後的：AAPL、NVDA 含 a
root.update()
app._export_db_overview()
check("檔案寫出來了", out.exists(), str(out))
if out.exists():
    raw = out.read_bytes()
    text = out.read_text(encoding="utf-8-sig")
    lines = text.strip().splitlines()
    check("有 BOM（Excel 開中文才不會亂碼）", raw.startswith(b"\xef\xbb\xbf"),
          f"head={raw[:6]!r}")
    check("匯出的是篩選後的 2 家 + 標題列", len(lines) == 3, f"lines={lines}")
    check("標題列寫的是「財報期間」（不是收件日、也不是含糊的「涵蓋期間」）",
          "財報期間" in lines[0] and "涵蓋" not in lines[0], lines[0])
    check("提示列顯示匯出筆數", "2" in app._db_hint_label.cget("text"),
          app._db_hint_label.cget("text"))

# ── 6.5 「更新選中的」只跑選中那幾家 ────────────────────────────────────
print("\n[6.5] 更新選中的")
app._db_search_var.set("")
root.update()

shown_dialogs = []
m.messagebox.showinfo = lambda *a, **k: shown_dialogs.append(("info", a))
m.messagebox.showerror = lambda *a, **k: shown_dialogs.append(("error", a))
m.messagebox.askyesno = lambda *a, **k: True

# 沒選任何一列 → 只提示，不發動
app._db_tree.selection_set(())
app._update_selected_companies()
check("沒選就按 → 跳提示、不發動",
      shown_dialogs and shown_dialogs[-1][0] == "info" and not app.is_running,
      f"dialogs={shown_dialogs} running={app.is_running}")

# 選兩列 → 只有那兩家進 worker
started = {}
app.cfg["identity"] = "Probe probe@example.com"
app._start_worker = lambda fn: started.setdefault("called", True)
_orig_update = app._start_local_db_update


def _spy_update(tickers=None):
    started["tickers"] = list(tickers) if tickers is not None else None
    return None


app._start_local_db_update = _spy_update
items = app._db_tree.get_children()
picked = [app._db_tree.item(i, "values")[0] for i in items[:2]]
app._db_tree.selection_set(items[:2])
app._update_selected_companies()
check("只把選中的兩家傳下去", started.get("tickers") == picked,
      f"傳下去的={started.get('tickers')} 選中的={picked}")
check("不是傳 None（那會變成跑整份名單 218 家）",
      started.get("tickers") is not None)
app._start_local_db_update = _orig_update

# ── 6.6 上次更新時間有顯示 ──────────────────────────────────────────────
print("\n[6.6] 上次更新時間")
col = m.DB_OVERVIEW_COLUMNS.index("checked")
vals = [app._db_tree.item(i, "values")[col] for i in app._db_tree.get_children()]
check("每一列都有『上次查』的值", all(v for v in vals), f"vals={vals}")
check("剛 seed 的顯示今天", any("今天" in v for v in vals), f"vals={vals}")

# ── 6.7 財報期間欄顯示的是期間、不是收件日 ──────────────────────────────
print("\n[6.7] 財報期間")
pcol = m.DB_OVERVIEW_COLUMNS.index("period_span")
aapl = [app._db_tree.item(i, "values") for i in app._db_tree.get_children()
        if app._db_tree.item(i, "values")[0] == "AAPL"][0]
check("AAPL 有財報期間", "~" in aapl[pcol], f"AAPL={aapl[pcol]}")
check("期間的月份是 03（季末）不是 05（收件月）",
      "-03" in aapl[pcol], f"AAPL={aapl[pcol]}（seed 的收件日是 05-01、期末是 03-29）")

# ── 7. 空資料庫不炸 ─────────────────────────────────────────────────────
print("\n[7] 空資料庫")
for ticker in ("AAPL", "BLK", "NVDA"):
    filing_cache.clear_ticker(ticker)
app._db_search_var.set("")
app._refresh_db_overview()
root.update()
check("清空後沒有列、也沒有例外", len(app._db_tree.get_children()) == 0)
check("摘要列顯示 0 家", "0" in app._db_summary_label.cget("text"),
      app._db_summary_label.cget("text"))

root.destroy()

print("\n" + "=" * 60)
if FAILURES:
    print(f"{len(FAILURES)} 項沒過：" + "、".join(FAILURES))
    sys.exit(1)
print("全部通過")
sys.exit(0)
