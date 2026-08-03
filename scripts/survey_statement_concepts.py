"""survey_statement_concepts.py — 調查美股三表實際出現的 XBRL 科目與覆蓋率。

用途：用數據決定 `IS_TEMPLATE` / `BS_TEMPLATE` / `CF_TEMPLATE` 該收哪些列。

判準（使用者 2026-08-03 定調）：
  「多數公司都有的項目，就算某些公司沒有，也該放進固定模板當空白列。
    只有極少數特例才有的，才該落到 overflow。」

方法：抓一批公司最新的 10-Q，逐一取 IS / BS / CF 的 XBRL 科目，統計**跨公司
出現率**。再比對現行模板，列出「高覆蓋率但沒收進模板」的科目——那些就是該補的。

**不呼叫 AI**：全程只走 EDGAR + edgartools 的 XBRL 解析，不耗用任何 API 額度。

用法：
    ./venv/Scripts/python.exe scripts/survey_statement_concepts.py [輸出.json]
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8")

import config
from edgar import Company, set_identity
from fetcher_gaap import (BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE,
                          _consolidated_mask)

# 跨產業 + 跨市值。刻意含金融股（三表結構不同）與公用事業、REIT。
TICKERS = [
    # 大型科技
    "AAPL", "MSFT", "NVDA", "GOOGL", "AMZN", "META", "AVGO", "ORCL", "TSLA",
    # 軟體 / SaaS
    "CRM", "NOW", "PANW", "ADBE", "WDAY",
    # 半導體
    "INTC", "AMD", "MU", "LRCX", "KLAC", "ADI", "TXN",
    # 中小型硬體 / 光通訊
    "COHR", "LITE", "ARLO", "FORM", "POWI", "AEIS",
    # 消費
    "COST", "WMT", "PG", "KO", "NKE", "SBUX",
    # 工業
    "CAT", "DE", "HON", "GE",
    # 能源 / 公用
    "XOM", "CVX", "NEE",
    # 醫療
    "JNJ", "PFE", "UNH", "ABBV",
    # 金融（三表結構不同，單獨看）
    "JPM", "BAC", "GS",
    # 電信 / REIT
    "T", "VZ", "PLD",
]

FINANCIALS = {"JPM", "BAC", "GS"}

STATEMENTS = ("IS", "BS", "CF")
TEMPLATES = {"IS": IS_TEMPLATE, "BS": BS_TEMPLATE, "CF": CF_TEMPLATE}


def _template_covers(statement: str, concept: str, std_concept: str) -> str | None:
    """這個 XBRL 科目有沒有被現行模板收走？回傳模板列名，沒有回 None。

    比對邏輯與 `_match_is_row` 同源：先比 standard_concept，再比 concept 的
    fallback 正則。只是這裡不需要 label_hint / match 這些精修參數。
    """
    for row in TEMPLATES[statement]:
        name, std, fallback, source, _match, _hint = row
        if source == "DERIVED":
            continue
        if std and std_concept and std_concept == std:
            return name
        if fallback and concept:
            try:
                if re.search(fallback, concept, re.IGNORECASE):
                    return name
            except re.error:
                if fallback.lower() in concept.lower():
                    return name
    return None


def _statement_frames(filing):
    """回傳 {'IS': df, 'BS': df, 'CF': df}，取不到的省略。"""
    out = {}
    try:
        fin = filing.obj().financials
    except Exception as exc:
        print(f"      obj() 失敗: {type(exc).__name__}", file=sys.stderr)
        return out
    for key, getter in (("IS", "income_statement"),
                        ("BS", "balance_sheet"),
                        ("CF", "cashflow_statement")):
        try:
            stmt = getattr(fin, getter)()
            df = stmt.get_dataframe() if hasattr(stmt, "get_dataframe") else stmt.to_dataframe()
            if df is not None and not df.empty:
                out[key] = df
        except Exception:
            continue
    return out


def survey(tickers: list[str]) -> dict:
    # {statement: {concept: {"label": str, "std": str, "tickers": set}}}
    seen: dict[str, dict[str, dict]] = {s: defaultdict(
        lambda: {"label": "", "std": "", "tickers": set()}) for s in STATEMENTS}
    ok, failed = [], []

    for i, ticker in enumerate(tickers, 1):
        print(f"[{i}/{len(tickers)}] {ticker}", flush=True)
        try:
            filings = Company(ticker).get_filings(form="10-Q")
            filing = filings[0] if len(filings) else None
        except Exception as exc:
            print(f"      清單失敗: {type(exc).__name__}", file=sys.stderr)
            failed.append(ticker)
            continue
        if filing is None:
            print("      找不到 10-Q")
            failed.append(ticker)
            continue

        frames = _statement_frames(filing)
        if not frames:
            failed.append(ticker)
            continue
        ok.append(ticker)

        for stmt, df in frames.items():
            try:
                rows = df[_consolidated_mask(df)]
            except Exception:
                rows = df
            for _, row in rows.iterrows():
                concept = str(row.get("concept", "") or "").strip()
                if not concept:
                    continue
                entry = seen[stmt][concept]
                entry["tickers"].add(ticker)
                if not entry["label"]:
                    entry["label"] = str(row.get("label", "") or "")[:70]
                std = str(row.get("standard_concept", "") or "")
                if std and std != "nan" and not entry["std"]:
                    entry["std"] = std
        print(f"      IS/BS/CF 科目數 "
              f"{'/'.join(str(len(frames.get(s, []))) for s in STATEMENTS)}")

    result = {"companies": ok, "failed": failed, "statements": {}}
    n_all = len(ok)
    n_nonfin = len([t for t in ok if t not in FINANCIALS])

    for stmt in STATEMENTS:
        items = []
        for concept, e in seen[stmt].items():
            tk = e["tickers"]
            nonfin = {t for t in tk if t not in FINANCIALS}
            items.append({
                "concept": concept,
                "label": e["label"],
                "std_concept": e["std"],
                "n_all": len(tk),
                "n_nonfin": len(nonfin),
                "pct_nonfin": round(len(nonfin) / n_nonfin * 100, 1) if n_nonfin else 0,
                "covered_by": _template_covers(stmt, concept, e["std"]),
                "tickers": sorted(tk),
            })
        items.sort(key=lambda x: -x["n_nonfin"])
        result["statements"][stmt] = items

    result["n_all"] = n_all
    result["n_nonfin"] = n_nonfin
    return result


def main() -> None:
    cfg = config.load_config()
    set_identity(cfg["identity"])
    result = survey(TICKERS)

    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("statement_concepts.json")
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + "=" * 78)
    print(f"成功 {result['n_all']} 家（非金融 {result['n_nonfin']} 家）；"
          f"失敗 {', '.join(result['failed']) or '無'}")
    for stmt in STATEMENTS:
        items = result["statements"][stmt]
        missing = [x for x in items if x["covered_by"] is None and x["pct_nonfin"] >= 30]
        print(f"\n### {stmt}：共 {len(items)} 個科目，"
              f"其中覆蓋率 ≥30% 但**模板沒收**的有 {len(missing)} 個")
        for x in missing[:30]:
            print(f"   {x['pct_nonfin']:>5.0f}%  {x['concept'][:52]:54} {x['label'][:40]}")
    print(f"\n完整結果：{out}")


if __name__ == "__main__":
    main()
