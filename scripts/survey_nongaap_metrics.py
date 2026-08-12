"""survey_nongaap_metrics.py — 調查美股 8-K 新聞稿實際使用的 Non-GAAP 指標。

用途：決定 `Data_NonGAAP` 的固定模板（Non-GAAP Core）要收哪些行。
方法：抓一批公司最新的 Item 2.02 8-K 新聞稿，用**純文字比對**統計指標出現頻率。

**不呼叫 AI**——這是刻意的。要決定「哪些指標夠通用」必須看原文實際怎麼寫，
用 AI 抽一遍等於讓 AI 的偏好污染統計結果，而且會吃配額。

輸出：
  1. 每家公司偵測到的 Non-GAAP 指標
  2. 跨公司出現頻率表（決定 core 的依據）
  3. 完全沒有 Non-GAAP 段落的公司清單（決定「沒有 Non-GAAP 時怎麼呈現」的依據）

用法：
    ./venv/Scripts/python.exe scripts/survey_nongaap_metrics.py [輸出.json]
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")

import config
from edgar import Company, set_identity

# 大中小型 + 跨產業。刻意包含金融股與消費股——它們多半不報 Non-GAAP，
# 正好用來回答「公司沒有提供 Non-GAAP 報表時怎麼呈現」。
TICKERS = [
    # 大型科技
    "AAPL", "MSFT", "NVDA", "GOOGL", "AMZN", "META", "AVGO", "ORCL", "TSLA",
    # 軟體 / SaaS
    "CRM", "NOW", "PANW", "WDAY", "ADBE",
    # 半導體
    "INTC", "AMD", "MU", "LRCX", "KLAC", "ADI", "SWKS", "QRVO",
    # 中小型硬體 / 光通訊
    "COHR", "LITE", "ARLO", "FORM", "POWI", "AEIS",
    # 非科技（對照組：多半不報 Non-GAAP）
    "JPM", "PG", "COST", "XOM",
]

# 指標片語：抓「non-GAAP / adjusted <名詞片語>」直到數字、標點或行尾。
#
# ⚠ 終止詞一定要加 \b。第一版沒加，"gross profit" 被切成 "gross pr"（後面的
# "ofit" 開頭就是終止詞 of）、"per share" 被切成 "per sh"（"are"）。
_STOP_WORDS = (
    "of|was|were|is|are|be|to|totaled|totalled|increased|decreased|grew|"
    "rose|fell|declined|improved|for|in|from|per|which|that|excludes|excluding"
)
_METRIC_RE = re.compile(
    r"\b(non-?GAAP|adjusted)\s+"
    r"([A-Za-z][A-Za-z0-9\-&/ ]{2,45}?)"
    r"(?=\s+(?:" + _STOP_WORDS + r")\b|\s*[:,.;()\[\]]|\s*\$|\s*\d|\s*%|$)",
    re.IGNORECASE,
)

# 這些字尾是句子殘骸不是指標名，砍掉
_TAIL_NOISE = re.compile(
    r"\s+(?:results?|measures?|financial|information|reconciliation|basis|"
    r"guidance|outlook|and|or|to|for|in|on|the|a|an|that|which|as|from|with)$",
    re.IGNORECASE,
)


def _clean(phrase: str) -> str:
    out = " ".join(phrase.split()).strip(" -&/")
    prev = None
    while prev != out:
        prev = out
        out = _TAIL_NOISE.sub("", out).strip()
    return out.lower()


def _metrics_in(text: str) -> list[str]:
    """從新聞稿全文抽出 Non-GAAP 指標候選名稱，依出現次數排序。"""
    found: Counter = Counter()
    for m in _METRIC_RE.finditer(text):
        name = _clean(f"{m.group(1)} {m.group(2)}")
        if len(name) > 6:
            found[name] += 1
    return [n for n, _ in found.most_common()]


def _press_release_text(filing) -> str | None:
    try:
        ek = filing.obj()
    except Exception as exc:
        print(f"      obj() 失敗: {type(exc).__name__}", file=sys.stderr)
        return None
    for pr in (getattr(ek, "press_releases", None) or []):
        try:
            text = pr.markdown() if hasattr(pr, "markdown") else pr.text()
        except Exception:
            continue
        if text:
            return unicodedata.normalize("NFKC", text)
    return None


def _latest_earnings_filing(ticker: str):
    company = Company(ticker)
    for filing in company.get_filings(form="8-K"):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" in items:
            return filing
    return None


def survey(tickers: list[str], cache_dir: Path | None = None) -> dict:
    per_company: dict[str, list[str]] = {}
    no_nongaap: list[str] = []
    failed: list[str] = []

    for i, ticker in enumerate(tickers, 1):
        print(f"[{i}/{len(tickers)}] {ticker}", flush=True)
        cached = cache_dir / f"{ticker}.md" if cache_dir else None
        if cached is not None and cached.exists():
            text = cached.read_text(encoding="utf-8")
            per_company[ticker] = _metrics_in(text)
            if not per_company[ticker] and "non-gaap" not in text.lower():
                no_nongaap.append(ticker)
            print(f"      （原文快取）{len(per_company[ticker])} 個候選指標")
            continue

        try:
            filing = _latest_earnings_filing(ticker)
        except Exception as exc:
            print(f"      清單失敗: {type(exc).__name__}", file=sys.stderr)
            failed.append(ticker)
            continue
        if filing is None:
            print("      找不到 Item 2.02 8-K")
            failed.append(ticker)
            continue

        text = _press_release_text(filing)
        if not text:
            print("      取不到新聞稿")
            failed.append(ticker)
            continue

        # 原文存檔：調整比對規則後可以重跑分析而不必重新下載
        if cache_dir is not None:
            cache_dir.mkdir(parents=True, exist_ok=True)
            (cache_dir / f"{ticker}.md").write_text(text, encoding="utf-8")

        if "non-gaap" not in text.lower() and "adjusted" not in text.lower():
            print("      新聞稿沒有 Non-GAAP 段落")
            no_nongaap.append(ticker)
            per_company[ticker] = []
            continue

        per_company[ticker] = _metrics_in(text)
        print(f"      {len(per_company[ticker])} 個候選指標")

    # 跨公司頻率：同一家只算一次
    freq: Counter = Counter()
    for names in per_company.values():
        for n in set(names):
            freq[n] += 1

    return {
        "per_company": per_company,
        "frequency": freq.most_common(),
        "no_nongaap": no_nongaap,
        "failed": failed,
        "n_companies": len([t for t in per_company if per_company[t]]),
    }


def main() -> None:
    cfg = config.load_config()
    set_identity(cfg["identity"])
    cache_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("nongaap_survey_text")
    result = survey(TICKERS, cache_dir=cache_dir)

    out_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("nongaap_survey.json")
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + "=" * 72)
    print(f"有 Non-GAAP 段落：{result['n_companies']} 家")
    print(f"沒有 Non-GAAP：{', '.join(result['no_nongaap']) or '無'}")
    print(f"抓取失敗：{', '.join(result['failed']) or '無'}")
    print("\n跨公司出現頻率（前 60）：")
    for name, count in result["frequency"][:60]:
        print(f"  {count:3d}  {name}")
    print(f"\n完整結果：{out_path}")


if __name__ == "__main__":
    main()
