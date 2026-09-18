# -*- coding: utf-8 -*-
"""check_filer_type.py — 掃更新名單，找出這個工具抓不到的公司（TODO D9）。

    ./venv/Scripts/python.exe scripts/check_filer_type.py

**加新公司進更新名單之前先跑這支**，比事後在 10 小時的抓取裡看它失敗便宜。

判準用實際行為而不是靠猜：這個工具是用 10-Q 抓季報的，`fetch_gaap_statements()`
拿不到 10-Q 就直接 raise。所以「有幾份 10-Q」就是會不會出事的直接指標。

順便記 10-K 與 20-F 的份數，分辨三種情況：
  - 有 10-Q          → 正常
  - 無 10-Q、有 20-F → D9 外國私人發行人，抓不到（預期，不是 bug）
  - 無 10-Q、無 20-F → 其他問題（SEC 查不到這個代號？剛上市？）

⚠ **不可以靠註冊地或公司名稱猜**，兩個方向都會猜錯。2026-09-18 實測 218 家：
  - 猜錯方向一：NXPI（荷蘭）、LIN（英國）、ACN（愛爾蘭）、TEL（瑞士）
    註冊地都在國外，**但全部報 10-Q**——SEC 認定它們是美國國內申報人
  - 猜錯方向二：ARM Holdings 2023 才在 NASDAQ 上市，看起來像正常美股，
    **實際只有 20-F**。當時憑印象列「三家有問題」就漏掉了它

決定申報身分的是股東結構與管理層所在地，不是註冊地，所以只能問 SEC。
每家 3 次請求（10-Q／10-K／20-F 各一），218 家約 5 分鐘。

2026-09-18 首跑（218 家）：ARM／ASML／STM 只有 20-F，IFX 在 SEC 查無此代號
（Infineon 在美國是 OTC ADR `IFNNY`，`IFX` 是德國交易所代號）。四家已移出
更新名單，剩 214 家。
"""
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, r"C:\Users\CTH\Documents\Code\SEC Financial Tools\src")

from config import load_config          # noqa: E402
import local_db                          # noqa: E402
from edgar import Company, set_identity  # noqa: E402

set_identity(load_config()["identity"])
tickers = local_db.get_update_list(load_config())
print(f"掃 {len(tickers)} 家…", flush=True)

rows = []
for i, ticker in enumerate(tickers, 1):
    row = {"ticker": ticker, "10-Q": None, "10-K": None, "20-F": None,
           "error": "", "name": ""}
    try:
        c = Company(ticker)
        row["name"] = str(getattr(c, "name", "") or "")[:40]
        for form in ("10-Q", "10-K", "20-F"):
            row[form] = len(list(c.get_filings(form=form)))
    except Exception as exc:                       # noqa: BLE001
        row["error"] = f"{type(exc).__name__}: {exc}"[:80]
    rows.append(row)
    if row["error"] or not row["10-Q"]:
        print(f"  [{i}/{len(tickers)}] ⚠ {ticker:<6} "
              f"10-Q={row['10-Q']} 10-K={row['10-K']} 20-F={row['20-F']} "
              f"{row['error']}", flush=True)
    elif i % 40 == 0:
        print(f"  [{i}/{len(tickers)}] …", flush=True)

out = ROOT / "scan_20f.json"
out.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"\n寫入 {out}")

bad = [r for r in rows if not r["10-Q"] or r["error"]]
print(f"\n=== 沒有 10-Q 或查詢失敗的：{len(bad)} 家 ===")
for r in bad:
    kind = ("D9 外國私人發行人（只有 20-F）" if r["20-F"] else
            "其他問題" if not r["error"] else "查詢失敗")
    print(f"  {r['ticker']:<6} {r['name']:<34} 10-Q={r['10-Q']} "
          f"20-F={r['20-F']}  → {kind}")
