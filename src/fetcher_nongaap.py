"""
fetcher_nongaap.py — Non-GAAP data extraction from 8-K press releases via AI.

Flow:
  8-K (Item 2.02) → edgartools eps_reconciliation + AI on EX-99.1 press release
  → nongaap_cache.json (per-ticker, incremental) → StatementTable list
"""

import json
import re
import sys
import time
import unicodedata
from pathlib import Path
from typing import Any

from edgar import Company, set_identity

import metric_rules
import nongaap_layout
from errsafe import _exc_status
from fetcher_gaap import StatementTable

CACHE_FILENAME = "nongaap_cache.json"


# ── Cache I/O ───────────────────────────────────────────────────────────────
#
# Cache format (new):  {ticker: {quarter_label: {filing_date, eps_recon, metrics}}}
# Cache format (old):  {quarter_label: {...}}  — single-ticker, no isolation
#
# Old-format detection: top-level key starts with "FY" (quarter labels look like FY2024Q1).
# On first load of an old file, the data is treated as belonging to the current ticker
# and will be written in new format on the next save.

def _load_cache(cache_path: Path, ticker: str) -> dict:
    """Return ticker's cached quarters from nongaap_cache.json.

    Returns {} if file is missing, malformed, or ticker has no cached data.
    Transparently handles old single-ticker format (no ticker key).
    """
    if not cache_path.exists():
        return {}
    try:
        with open(cache_path, encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}

    if not isinstance(data, dict) or not data:
        return {}

    # Old format: keys are quarter labels like "FY2024Q1"
    if next(iter(data)).startswith("FY"):
        return data  # treat as this ticker's data; migrated to new format on next save

    return data.get(ticker, {})


def _save_cache(cache_path: Path, ticker: str, ticker_data: dict) -> None:
    """Write ticker's data into nongaap_cache.json (multi-ticker format).

    Reads the existing file, updates this ticker's slice, and writes back.
    Old-format files are silently replaced with new format on first write
    (ticker_data already contains all previously loaded quarters).
    """
    if cache_path.exists():
        try:
            with open(cache_path, encoding="utf-8") as f:
                all_data = json.load(f)
            if not isinstance(all_data, dict):
                all_data = {}
            # Old-format file: discard raw dict — ticker_data already has the migrated content
            if all_data and next(iter(all_data)).startswith("FY"):
                all_data = {}
        except (json.JSONDecodeError, OSError):
            all_data = {}
    else:
        all_data = {}

    all_data[ticker] = ticker_data
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)


# ── Period helpers ───────────────────────────────────────────────────────────

def _period_to_quarter_label(period_of_report: str) -> str:
    """Convert '20240331' or '2024-03-31' to 'FY2024Q1'."""
    period = period_of_report.replace("-", "")
    year = period[:4]
    month = int(period[4:6])
    if month <= 3:
        suffix = "Q1"
    elif month <= 6:
        suffix = "Q2"
    elif month <= 9:
        suffix = "Q3"
    else:
        suffix = "Q4"
    return f"FY{year}{suffix}"


# ── StatementTable builders ──────────────────────────────────────────────────

def _build_eps_recon_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_EPS_Recon StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # Collect all EPS recon keys (union across quarters)
    all_keys: list[str] = []
    seen: set[str] = set()
    for q in sorted_qs:
        for key in cache[q].get("eps_recon", {}):
            if key not in seen:
                all_keys.append(key)
                seen.add(key)

    if not all_keys:
        return None

    values: list[list[Any]] = []
    for key in all_keys:
        values.append([cache[q].get("eps_recon", {}).get(key) for q in sorted_qs])

    return StatementTable(
        sheet_name="Data_EPS_Recon",
        quarter_labels=sorted_qs,
        filing_dates=filing_dates,
        concepts=all_keys,
        values=values,
        ticker=ticker,
        labels=[""] * len(all_keys),
    )


def _build_nongaap_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_NonGAAP StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # 正規化與中英對照都在「讀取」時做，不是寫入快取時做。理由有二：
    #   1. 既有 nongaap_cache.json 存的是舊規則的產物（中文、帶期間 token），
    #      在這裡重跑一次就能救回來，不必刪快取重呼叫 AI。
    #   2. 改 metric_rules.py 的規則表後重跑即生效，不必重抓。
    # _normalize_nongaap_metrics 具冪等性，對已經乾淨的名稱再跑一次不會變樣。
    per_quarter: dict[str, dict[str, Any]] = {}
    for q in sorted_qs:
        normalized = _normalize_nongaap_metrics(cache[q].get("metrics", {}) or {})
        per_quarter[q] = {
            _canonicalize_metric_name(name): val for name, val in normalized.items()
        }

    # 跨季合併：同一指標的不同寫法（中／英、大小寫漂移）併成同一個顯示名，
    # 首見的勝出。合併在這裡做完，版面模組拿到的就是已對齊的名稱。
    display_by_key: dict[str, str] = {}
    for q in sorted_qs:
        for name in per_quarter[q]:
            display_by_key.setdefault(_metric_merge_key(name), name)

    aligned: dict[str, dict[str, Any]] = {}
    for q in sorted_qs:
        aligned[q] = {
            display_by_key[_metric_merge_key(name)]: val
            for name, val in per_quarter[q].items()
        }

    # 版面（core / 調節 / overflow / 年度分區）交給 nongaap_layout。
    # 即使一格資料都沒有也要回一張骨架表——讀不到 sheet 與讀到空 sheet 是兩種
    # 訊號，前者無法區分「這家沒報 Non-GAAP」與「抓取失敗」。
    return nongaap_layout.build_nongaap_table(ticker, aligned, sorted_qs, filing_dates)


# ── Extraction functions ─────────────────────────────────────────────────────

def _extract_eps_recon(eight_k) -> dict[str, float]:
    """Extract EPS reconciliation using edgartools native support.

    Returns dict like {"GAAP EPS": 0.53, "SBC": -0.12, "Non-GAAP EPS": 0.65}.
    Returns {} if not available.
    """
    try:
        earnings = getattr(eight_k, "earnings", None)
        if earnings is None:
            return {}
        recon = getattr(earnings, "eps_reconciliation", None)
        if recon is None:
            return {}
        df = recon.dataframe
        if df is None or df.empty:
            return {}

        result: dict[str, float] = {}
        value_cols = [c for c in df.columns if c not in {"label", "concept", "description"}]
        if not value_cols:
            return {}
        val_col = value_cols[0]

        label_col = "label" if "label" in df.columns else df.columns[0]
        for _, row in df.iterrows():
            label = str(row.get(label_col, "") or "").strip()
            val = row.get(val_col)
            if label and val is not None:
                try:
                    result[label] = float(val)
                except (ValueError, TypeError):
                    pass
        return result
    except Exception as exc:
        print(f"[fetcher_nongaap] eps_recon warning: {exc!r}", file=sys.stderr)
        return {}


# ── Metric name normalisation ────────────────────────────────────────────────
#
# Press releases (especially NVDA) embed period tokens and comparison-period
# rows in their Non-GAAP tables, causing the AI to return names like:
#
#   "Non-GAAP Gross margin (Q4 FY26)"   ← trailing suffix — strip
#   "Non-GAAP Q4 FY26 Gross margin"     ← period in middle — strip
#   "Q2 FY26 Non-GAAP Revenue"          ← period as prefix — strip
#   "Non-GAAP Q3 FY26 Gross margin"     ← comparison quarter — dedup, discard
#   "Non-GAAP FY2026 Gross margin"      ← annual duplicate — discard if Q exists
#   "Expected Non-GAAP Gross margin (Q1 FY27)"  ← guidance — always discard
#
# Strategy:
#   1. Drop guidance / outlook entries.
#   2. Strip ALL period tokens (Q\d FY\d+ or FY\d+) and trailing noise labels.
#   3. Two-pass dedup: quarterly-tagged entries win over FY-only entries;
#      first occurrence wins within each bucket (press release shows current
#      quarter first, so comparison-period rows are skipped automatically).

_PERIOD_TOKEN_RE = re.compile(r"\b(?:Q[1-4]\s+)?FY\d{2,4}\b", re.IGNORECASE)
_QTR_TOKEN_RE    = re.compile(r"\bQ[1-4]\s+FY\d{2,4}\b", re.IGNORECASE)
_TRAILING_LABEL_RE = re.compile(
    r"\s*\(\s*(?:Table|Summary|Chart|Note|Detail|Details)\s*\)\s*$", re.IGNORECASE
)
_GUIDANCE_PREFIXES = metric_rules.GUIDANCE_PREFIXES_EN

# 中文樣式（規則表在 metric_rules.py，這裡只負責編譯）。
# 年度必須先於季度比對：「2025年全年度」若先跑季度樣式不會誤中，但把年度放前面
# 語意更清楚，也避免日後有人加了寬鬆的季度樣式而互相吃到。
_ZH_ANNUAL_RE  = re.compile("|".join(metric_rules.ZH_ANNUAL_PATTERNS))
_ZH_QUARTER_RE = re.compile("|".join(metric_rules.ZH_QUARTER_PATTERNS))


def _clean_metric_name(name: str) -> str:
    """Strip period tokens (EN + ZH) and trailing noise labels, normalise whitespace."""
    clean = _PERIOD_TOKEN_RE.sub("", name)
    clean = _ZH_ANNUAL_RE.sub("", clean)
    clean = _ZH_QUARTER_RE.sub("", clean)
    clean = _TRAILING_LABEL_RE.sub("", clean)
    clean = re.sub(r"\(\s*\)", "", clean)   # remove empty parens left by period removal
    return re.sub(r"\s+", " ", clean).strip()


# ── 中英對照 ────────────────────────────────────────────────────────────────
#
# 跑在 _clean_metric_name 之後、表格組裝的時候（不是寫入快取的時候）。
# 這個順序是刻意的：規則表改了不必重抓 8-K，重跑就套用新對照。
# 詞彙表長詞優先，確保「訂閱與服務毛利率」整段命中而不是被拆成「服務」+「毛利率」。

_ZH_TERMS_SORTED = sorted(metric_rules.ZH_TERMS.items(), key=lambda kv: -len(kv[0]))


def _canonicalize_metric_name(name: str) -> str:
    """中文詞彙替換 + 同義名合併。未收錄的名稱原樣回傳（不可吞資料）。"""
    out = name
    for zh, en in _ZH_TERMS_SORTED:
        if zh in out:
            out = out.replace(zh, f" {en} ")
    out = re.sub(r"\s+", " ", out).strip()

    # (FY) 標記先摘下來再查對照表，否則年度列永遠對不到表、跨季合併不起來
    suffix = metric_rules.FY_ROW_SUFFIX
    is_fy = out.endswith(suffix.strip()) or out.endswith(suffix)
    if is_fy:
        out = out[: out.rfind("(")].strip()

    out = metric_rules.METRIC_ALIASES.get(out.casefold(), out)
    return f"{out}{suffix}" if is_fy else out


def _metric_merge_key(name: str) -> str:
    """跨季合併用的比對鍵：忽略大小寫、空白與標點。

    對照表收錄不到的名稱也能靠這個把大小寫漂移的變體併起來。
    """
    return re.sub(r"[^0-9a-z一-鿿]+", "", name.casefold())


def _normalize_nongaap_metrics(raw: dict[str, Any]) -> dict[str, Any]:
    """Strip period tokens, drop guidance rows, deduplicate (quarterly > FY).

    Two-pass approach:
      Pass 1 — entries that had a Q# token (quarterly) or no period token at all.
      Pass 2 — entries that had only an FY token (annual totals).
    First occurrence of each clean name wins within each pass.
    """
    quarterly: dict[str, Any] = {}
    fy_only: dict[str, Any] = {}

    for name, val in raw.items():
        name_lower = name.lower().strip()
        # Drop guidance / outlook / forward-looking entries
        if any(name_lower.startswith(p) for p in _GUIDANCE_PREFIXES):
            continue
        if any(s in name_lower for s in metric_rules.GUIDANCE_SUBSTRINGS_EN):
            continue
        # 中文 guidance 詞常出現在名稱中間（「2026財年預期 Non-GAAP 營業利潤率上限」），
        # 所以用「包含」比對而非 startswith。
        if any(s in name for s in metric_rules.GUIDANCE_SUBSTRINGS_ZH):
            continue

        has_qtr  = bool(_QTR_TOKEN_RE.search(name)) or bool(_ZH_QUARTER_RE.search(name))
        has_period = (
            bool(_PERIOD_TOKEN_RE.search(name))
            or bool(_ZH_ANNUAL_RE.search(name))
            or has_qtr
        )
        is_fy_only = has_period and not has_qtr  # 年度 token 有、季度 token 無

        clean = _clean_metric_name(name)
        if not clean:
            continue

        if is_fy_only:
            if clean not in fy_only:
                fy_only[clean] = val
        else:
            if clean not in quarterly:
                quarterly[clean] = val

    # 年度值的處理見 metric_rules.FY_ONLY_HANDLING（預設 "label"：另成一列不佔季欄位）
    result = dict(quarterly)
    mode = metric_rules.FY_ONLY_HANDLING
    for clean, val in fy_only.items():
        if mode == "drop":
            continue
        if mode == "fill":
            if clean not in result:
                result[clean] = val
            continue
        labelled = f"{clean}{metric_rules.FY_ROW_SUFFIX}"
        if labelled not in result:
            result[labelled] = val
    return result


# ── AI extraction ────────────────────────────────────────────────────────────

# prompt 一律用英文寫、且明確要求英文指標名（2026-08-01 改）。
#
# 原本 prompt 是中文，AI 就回中文指標名——而且不穩定，同一個 ticker 內都會混
# （CRM FY2026Q2 回中文、FY2026Q1 回英文）。下游的期間剝除、guidance 過濾、
# Excel ÷1M 豁免三條規則當時全部只認英文，整張 Data_NonGAAP 因此不可用。
#
# 這裡是第一道防線（減少中文輸入）；第二道是 metric_rules.py 的中英對照層
# （AI 不聽話回中文時接住）。兩道都要，因為 prompt 遵從度不是保證。
_NONGAAP_PROMPT = """\
You are a financial analyst. Below is a company quarterly earnings press release
(Markdown format). Extract all Non-GAAP financial metrics and return JSON:

{{"Metric Name": value}}

Rules:
- Metric names MUST be in English. Never use Chinese or any other language.
- Use standard US financial terminology: "Non-GAAP Gross Margin",
  "Non-GAAP Diluted EPS", "Adjusted EBITDA", "Adjusted EBITDA Margin",
  "Free Cash Flow", "Non-GAAP Operating Income", "Non-GAAP Revenue".
- Include Non-GAAP / Adjusted / Excluding metrics.
- ALSO include the GAAP counterparts the release states alongside them, named with
  a "GAAP " prefix: "GAAP Revenue", "GAAP Gross Margin", "GAAP Operating Income",
  "GAAP Operating Margin", "GAAP Net Income", "GAAP Diluted EPS".
  Only if the release actually states them — never compute or infer them.
- ALSO include the individual reconciling items that bridge GAAP net income to
  Non-GAAP net income, using exactly these names when they apply:
  "Stock-Based Compensation", "Amortization of Intangibles",
  "Restructuring Charges", "Impairment Charges", "Litigation and Settlement",
  "Acquisition-Related Costs", "Tax Effect of Adjustments".
  Give each as the SIGNED amount ADDED to GAAP net income to reach Non-GAAP net
  income (so the tax effect is normally negative).
- Values must be plain numbers — no currency symbols, no commas.
- Convert amounts to absolute units: millions x 1000000, billions x 1000000000.
- Percentages as the bare number (17.6% -> 17.6).
- Do NOT include guidance, outlook, or forward-looking figures for future periods.
- Extract figures for the CURRENT reported quarter only. Do NOT include full-year,
  full year, year-to-date, LTM / trailing-twelve-month, or prior-period comparison
  figures, even when the release presents them alongside the quarterly numbers.
- If no Non-GAAP metrics are found, return empty JSON {{}}.
- Return JSON only, no explanation.

Press release:
{press_release_text}
"""


# ── AI 呼叫重試（2026-08-01 加）───────────────────────────────────────────
#
# 實跑撞到 Gemini 的 HTTP 429。兩種 429 要分開看：
#   每分鐘限流 → 等一下重試就會過，退避有效
#   每日配額用盡 → 重試一定失敗，只是白等
# 所以次數壓低、退避拉開，寧可少試也不要讓批次更新卡住。跑完會統計未取得的季數，
# 使用者換一把 key 或隔天再跑一次即可（失敗的季不寫快取，重跑會自動補）。
AI_MAX_ATTEMPTS         = 3          # 含第一次，總共打幾次
AI_RETRY_BACKOFF_SECONDS = (5, 15)   # 第 1、2 次重試前各等幾秒

# 送進 prompt 的新聞稿長度上限。
#
# 原本是 12,000，是舊 prompt 時代留下的保守值。2026-08-02 實跑 ARLO 才發現這是
# 調節表抓不到的**真正原因**：ARLO 新聞稿全長 53,569 字元，「Stock-based
# compensation」出現在 18,605 / 33,759 / 38,440 / 40,558、「Amortization」在
# 40,848——全部在截斷之後，AI 根本沒看到。重點條列都在文件最前面（所以毛利率、
# EPS 一直抓得到），但**調節表一律在文件尾端**。
#
# 200K 字元約 50K token，Gemini Flash 的 context 是 100 萬 token，綽綽有餘；
# 上限仍保留，避免異常長的文件把單次呼叫撐爆。
PROMPT_TEXT_LIMIT = 200_000


def _is_quota_error(exc: BaseException) -> bool:
    """是不是「額度用盡」型的失敗（HTTP 429）。

    Gemini 的額度**按請求次數**計算，所以撞到 429 之後再重試每一次都必敗、
    而且每一次都扣一次額度。實測 CRM 一趟燒掉 12 次呼叫換到 0 筆資料
    （4 季 × 3 次嘗試）。這個判斷讓重試只用在真正有機會成功的暫時性錯誤上。
    """
    return _exc_status(exc).strip().endswith("429")


def _ai_request(prompt: str, ai_config: dict) -> str | None:
    """實際打 AI provider，回傳原始文字。provider 設錯回 None，其餘失敗直接拋。

    抽成獨立函式是為了讓重試邏輯與測試有一個乾淨的接縫——測試不必碰三家 SDK。
    """
    provider = ai_config.get("provider", "google")
    model    = ai_config.get("model", "")
    api_key  = ai_config.get("api_key", "")

    if provider == "google":
        from google import genai
        client = genai.Client(api_key=api_key)
        return client.models.generate_content(model=model, contents=prompt).text
    if provider == "openai":
        from openai import OpenAI
        response = OpenAI(api_key=api_key).chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1024,
        )
        return response.choices[0].message.content
    if provider == "anthropic":
        import anthropic
        response = anthropic.Anthropic(api_key=api_key).messages.create(
            model=model, max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text
    return None      # provider 設錯：不是「沒有指標」，不可寫快取


def _call_ai(text: str, ai_config: dict) -> dict[str, Any] | None:
    """Call configured AI provider with press release text.

    回傳 dict = 呼叫成功（空 dict 代表新聞稿真的沒有 Non-GAAP 指標）。
    回傳 None = 呼叫失敗（例外、429、provider 設錯）。呼叫端必須據此**不寫快取**，
    否則一次暫時性失敗會把該季永久標記為「已抓過但沒有資料」。
    """
    prompt = _NONGAAP_PROMPT.format(press_release_text=text[:PROMPT_TEXT_LIMIT])

    for attempt in range(1, AI_MAX_ATTEMPTS + 1):
        try:
            raw = _ai_request(prompt, ai_config)
            if raw is None:
                return None      # provider 設錯，重試也沒用

            # Strip markdown code fences if present            raw = raw.strip()
            if raw.startswith("```"):
                raw = "\n".join(raw.split("\n")[1:])
            if raw.endswith("```"):
                raw = raw.rsplit("```", 1)[0]

            parsed = json.loads(raw.strip())
            result = {}
            for k, v in parsed.items():
                if v is None:
                    continue
                try:
                    result[k] = float(v)
                except (ValueError, TypeError):
                    pass
            return _normalize_nongaap_metrics(result)

        except Exception as exc:
            # 只印類型 + status code。不可用 {exc!r} / {exc}：三家 LLM SDK 的例外訊息
            # 天生挾帶 URL（google-genai 走 REST 時帶 ?key=），而 launcher.ps1
            # 刻意留著主控台視窗，等於把金鑰印在畫面上。
            print(
                f"[fetcher_nongaap] AI call failed: {type(exc).__name__}"
                f"{_exc_status(exc)} | 嘗試 {attempt}/{AI_MAX_ATTEMPTS}",
                file=sys.stderr,
            )
            if _is_quota_error(exc):
                # 額度用盡：重試必敗且每次都扣額度，直接放棄這一季
                return None
            if attempt < AI_MAX_ATTEMPTS:
                time.sleep(AI_RETRY_BACKOFF_SECONDS[attempt - 1])

    return None


def _extract_nongaap_metrics(eight_k, ai_config: dict) -> dict[str, Any] | None:
    """Get press release text and call AI to extract Non-GAAP metrics.

    回傳 dict = 成功（{} 代表新聞稿真的沒有 Non-GAAP 指標，可以寫快取）。
    回傳 None = 失敗（AI 呼叫掛掉、下載出錯），呼叫端必須跳過寫快取以便下次重試。
    """
    try:
        press_releases = getattr(eight_k, "press_releases", None)
        text = None

        if press_releases:
            for pr in press_releases:
                try:
                    text = pr.markdown() if hasattr(pr, "markdown") else pr.text()
                    if text:
                        break
                except Exception:
                    continue

        # Fallback: search attachments for EX-99
        if not text:
            try:
                attachments = getattr(eight_k, "_filing", None)
                if attachments:
                    attachments = attachments.attachments
                    for att in attachments:
                        doc_type = str(getattr(att, "document_type", "") or "")
                        if "EX-99" in doc_type.upper():
                            text = att.markdown() if hasattr(att, "markdown") else att.text()
                            if text:
                                break
            except Exception:
                pass

        if not text:
            return {}

        text = unicodedata.normalize("NFKC", text)
        return _call_ai(text, ai_config)

    except Exception as exc:
        # 在 AI 呼叫鏈上（下方 _call_ai）。目前 _call_ai 自己吞例外不會冒上來，
        # 但仍只印類型 + status：哪天內層改成 re-raise，這裡就會變成金鑰洩漏點。
        print(
            f"[fetcher_nongaap] metrics extraction failed: "
            f"{type(exc).__name__}{_exc_status(exc)}",
            file=sys.stderr,
        )
        return None


# ── 8-K discovery ────────────────────────────────────────────────────────────

def _filter_nongaap_by_year(
    filings: list[tuple],
    start_year: int | None,
    end_year: int | None,
) -> list[tuple]:
    """Filter (label, filing) tuples by year extracted from label (e.g. 'FY2021Q2' → 2021)."""
    if start_year is None and end_year is None:
        return filings
    result = []
    for item in filings:
        label = item[0]
        m = re.search(r'(\d{4})', label)
        if m is None:
            result.append(item)
            continue
        year = int(m.group(1))
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(item)
    return result


def _has_exhibit(filing) -> bool:
    """True when the filing declares Item 9.01 (Financial Statements and Exhibits).

    An earnings 8-K carries its press release as an EX-99 exhibit, which is what
    Item 9.01 announces. A 2.02 filing without it has no press release to parse.
    Read off listing metadata — no document is downloaded.
    """
    return "9.01" in str(getattr(filing, "items", "") or "")


def _dedupe_by_label(candidates: list[tuple[str, Any]]) -> list[tuple[str, Any]]:
    """Collapse same-label filings to one, keeping the best candidate.

    ``candidates`` arrives newest-first, as EDGAR lists filings. Ranking:

    1. **Has an exhibit** (Item 9.01) beats one without. WDC FY2025Q1 collided an
       Item 2.02+5.02 filing with no press release against the real earnings
       release 13 days later; keeping the oldest silently lost the whole quarter.
    2. **Newest** wins ties. A preliminary release always precedes the final one,
       so recency is what keeps the official numbers — QRVO FY2025Q4 kept a
       "Preliminary ... Results" filing under the previous rule.

    Amendments never reach here (``get_filings(amendments=False)``), so the old
    "oldest wins, a corrected re-filing must not displace the original" rationale
    no longer buys anything.
    """
    best: dict[str, tuple[bool, int, Any]] = {}
    for index, (label, filing) in enumerate(candidates):     # newest → oldest
        rank = (_has_exhibit(filing), -index)                # newest = smallest index
        if label not in best or rank > best[label][:2]:
            best[label] = (rank[0], rank[1], filing)

    order = list(dict.fromkeys(label for label, _ in candidates))   # newest → oldest
    return [(label, best[label][2]) for label in order]


def _list_earnings_filings(
    company,
    start_year: int | None = None,
    end_year: int | None = None,
    max_filings: int = 80,
) -> list[tuple[str, Any]]:
    """Return [(quarter_label, filing)] for earnings 8-Ks, newest first.

    Filters entirely on listing metadata (``items`` and ``period_of_report``),
    which EDGAR supplies with the filing index — no document is downloaded here.
    Callers download only the filings they actually need.

    Item 2.02 is "Results of Operations and Financial Condition", i.e. the
    earnings release. SEC adopted that numbering on 2004-08-23; earlier filings
    used Item 12 or Item 5 and are not matched. See the design doc for why that
    is acceptable (max_filings defaults to 80 quarters ≈ 20 years).
    """
    candidates: list[tuple[str, Any]] = []
    for filing in company.get_filings(form="8-K", amendments=False):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" not in items:
            continue
        period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
        if len(period) < 8:
            continue
        try:
            label = _period_to_quarter_label(period)
        except Exception as exc:
            print(
                f"[fetcher_nongaap] listing skip {period}: "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )
            continue
        candidates.append((label, filing))

    deduped = _dedupe_by_label(candidates)

    deduped = _filter_nongaap_by_year(deduped, start_year, end_year)
    return deduped[:max_filings]


def _quarter_ordinal(label: str) -> int | None:
    """Convert 'FY2024Q3' to a sortable integer (2024*4 + 2). None if unparseable."""
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", label.strip())
    if m is None:
        return None
    return int(m.group(1)) * 4 + (int(m.group(2)) - 1)


def _ordinal_to_quarter(ordinal: int) -> str:
    """Inverse of _quarter_ordinal."""
    return f"FY{ordinal // 4}Q{ordinal % 4 + 1}"


def _find_missing_quarters(labels: list[str]) -> list[str]:
    """Return quarter labels absent between the oldest and newest supplied label.

    A gap means the listing-stage filter missed an earnings release — usually an
    8-K that omitted Item 2.02. Nothing outside the supplied span counts as a gap:
    a company simply has no filings before its IPO or after its latest report.
    """
    ordinals = sorted(o for o in (_quarter_ordinal(x) for x in labels) if o is not None)
    if len(ordinals) < 2:
        return []
    present = set(ordinals)
    return [
        _ordinal_to_quarter(o)
        for o in range(ordinals[0], ordinals[-1] + 1)
        if o not in present
    ]


def _recover_missing_quarters(company, missing: list[str]) -> list[tuple[str, Any]]:
    """Deep-scan only the quarters the listing filter came up short on.

    Downloads a filing only when its period falls in a missing quarter and it was
    not already tagged Item 2.02 — typically a handful of filings, versus the
    hundreds a full historical scan would fetch.
    """
    if not missing:
        return []

    wanted = set(missing)
    found: dict[str, Any] = {}
    for filing in company.get_filings(form="8-K", amendments=False):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" in items:
            continue
        period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
        if len(period) < 8:
            continue
        try:
            label = _period_to_quarter_label(period)
        except Exception as exc:
            print(
                f"[fetcher_nongaap] gap scan skip {period}: "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )
            continue
        if _quarter_ordinal(label) is None:
            continue
        if label not in wanted or label in found:
            continue
        try:
            eight_k = filing.obj()
        except Exception as exc:
            print(
                f"[fetcher_nongaap] gap scan {label} -> "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )
            continue
        if getattr(eight_k, "has_earnings", False):
            found[label] = filing

    return [(label, found[label]) for label in sorted(found, key=_quarter_ordinal)]


# ── Public API ───────────────────────────────────────────────────────────────

def fetch_nongaap_statements(
    ticker: str,
    identity: str,
    ai_config: dict,
    output_dir: Path,
    progress_cb=None,
    max_filings: int = 80,
    start_year: int | None = None,
    end_year: int | None = None,
) -> list[StatementTable]:
    """Fetch Non-GAAP statements from 8-K filings for a ticker.

    Args:
        ticker:      Stock ticker, e.g. "AAPL"
        identity:    SEC EDGAR identity string
        ai_config:   {"provider": ..., "model": ..., "api_key": ...}
        output_dir:  Directory where nongaap_cache.json will be stored
        progress_cb: Optional callable(current, total, label) for progress updates
        max_filings: Max number of earnings quarters to process (newest first, default 80).
                     Applied after the year range narrows the pool.

    Returns:
        List of StatementTable: [Data_EPS_Recon, Data_NonGAAP] (omits None tables)
    """
    set_identity(identity)
    company = Company(ticker)
    cache_path = Path(output_dir) / CACHE_FILENAME

    cache = _load_cache(cache_path, ticker)
    filings = _list_earnings_filings(company, start_year, end_year, max_filings)

    # A hole in the quarter sequence means the listing filter missed a release —
    # scan just that range rather than every 8-K the company ever filed.
    missing = _find_missing_quarters([label for label, _ in filings])
    if missing:
        recovered = _recover_missing_quarters(company, missing)
        if recovered:
            filings = sorted(
                filings + recovered,
                key=lambda item: _quarter_ordinal(item[0]) or 0,
                reverse=True,
            )
            # 回補後重新裁切（TODO 第 4 項）。_list_earnings_filings() 已經套過
            # 一次 max_filings，但回補是在切片之後才把缺季加回來，不重切的話
            # 「要 4 季、保留區間有 2 個缺口」會實際處理 6 季——每多一季就多一次
            # AI 呼叫，直接吃配額。裁切保留最新的。
            filings = filings[:max_filings]
        still_missing = sorted(set(missing) - {label for label, _ in recovered})
        if still_missing:
            print(
                f"[fetcher_nongaap] {ticker} 無此季財報 8-K: {', '.join(still_missing)}",
                file=sys.stderr,
            )

    new_filings = [(lbl, f) for lbl, f in filings if lbl not in cache]
    total = len(new_filings)
    failed_quarters: list[str] = []

    for i, (quarter_label, filing) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eight_k = filing.obj()      # the only download in this loop
            eps_recon = _extract_eps_recon(eight_k)
            metrics   = _extract_nongaap_metrics(eight_k, ai_config)
            if metrics is None:
                # AI 呼叫失敗（例如 HTTP 429）。**不寫快取**——寫了的話下次執行
                # `lbl not in cache` 會命中，這一季就永遠不會再被抓。
                failed_quarters.append(quarter_label)
                continue
            cache[quarter_label] = {
                "filing_date": str(filing.filing_date),
                "eps_recon":   eps_recon,
                "metrics":     metrics,
            }
            # Save after each quarter (crash-safe incremental)
            _save_cache(cache_path, ticker, cache)
        except Exception as exc:
            # 這層同時包住 filing.obj() 下載與 _extract_nongaap_metrics -> _call_ai 的 AI 呼叫鏈。
            print(
                f"[fetcher_nongaap] {quarter_label} failed: "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )
            failed_quarters.append(quarter_label)
            if _is_quota_error(exc):
                # 額度已用盡，這趟剩下的季再打也只是白扣額度
                remaining = [lbl for lbl, _ in new_filings[i:]
                             if lbl not in cache and lbl not in failed_quarters]
                failed_quarters.extend(remaining)
                print(
                    f"[fetcher_nongaap] {ticker} AI 額度用盡，本趟停止"
                    f"（尚餘 {len(remaining)} 季未抓）",
                    file=sys.stderr,
                )
                break

    # 未取得的季度要講清楚。這些季**沒有寫進快取**，直接重跑就會補抓；
    # 若是 AI 配額用盡（HTTP 429），換一把 key 或隔天再跑即可。
    # 同時推給 progress_cb——GUI 使用者看不到 stderr。
    if failed_quarters:
        summary = (
            f"{ticker}：{len(failed_quarters)} 季未取得 Non-GAAP "
            f"（{', '.join(failed_quarters)}），未寫入快取，重跑即會補抓"
        )
        print(f"[fetcher_nongaap] {summary}", file=sys.stderr)
        if progress_cb:
            progress_cb(total, total, summary)

    tables: list[StatementTable] = []
    eps_tbl = _build_eps_recon_table(ticker, cache)
    ng_tbl  = _build_nongaap_table(ticker, cache)
    if eps_tbl:
        tables.append(eps_tbl)
    if ng_tbl:
        tables.append(ng_tbl)
    return tables
