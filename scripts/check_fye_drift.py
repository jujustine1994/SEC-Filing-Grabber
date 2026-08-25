"""check_fye_drift.py — 量「公司改過財年」對 8-K 零下載規則的風險（TODO B6）。

**不打網路**，吃 `output/_spike/facts_*.json`（201 家的 companyfacts 快取）。

## 為什麼要量這個

B5（2026-08-25）之後，Item 2.02 8-K 的季度標籤是用 EDGAR 的
`Company.fiscal_year_end`（完整 MMDD）回推名目季末算出來的。**EDGAR 只給「現在」
的值**——公司改過財年的話，拿現值去回推改制以前的申報會整段偏掉。

交接文件原本寫「加一道 0~70 天 sanity check 就接得住」，那是**恆真式**：候選季末
永遠相隔 89~92 天，規則取「不晚於 發布日+tol 的最新候選」，選中的必然落在
`[-tol, 91-tol)`，tol=21 時上界剛好 70。所以那道檢查一次都攔不到，FYE 漂移必須
另外量。

## 怎麼判定「改過財年」

從每份 10-K 的資產負債表時點事實取該年度的期末日（同一個 accession 取最大的
`end`），排序後跟**最新那年**的月日比。52/53 週制的期末日本來就會在月底前後
浮動（一年最多差 7 天，跨閏年再差 1 天），所以門檻放 `> 14 天`才算真的改過。

    ./venv/Scripts/python.exe scripts/check_fye_drift.py [門檻天數]
"""
import json
import pathlib
import sys
from collections import defaultdict
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.stdout.reconfigure(encoding="utf-8")

THRESHOLD = int(sys.argv[1]) if len(sys.argv) > 1 else 14
CACHE = ROOT / "output" / "_spike"

# 資產負債表的時點事實，幾乎每家每年都有；用它取「那個財年結束在哪一天」。
_ANCHORS = ("Assets", "StockholdersEquity", "LiabilitiesAndStockholdersEquity")


def _fy_ends(facts: dict) -> list[str]:
    """每份 10-K 的財年結束日（同一個 accession 取最大的 end），由舊到新。"""
    by_accn: dict[str, str] = {}
    us_gaap = facts.get("facts", {}).get("us-gaap", {})
    for anchor in _ANCHORS:
        for unit_rows in us_gaap.get(anchor, {}).get("units", {}).values():
            for row in unit_rows:
                if row.get("form") != "10-K" or row.get("fp") != "FY":
                    continue
                accn, end = row.get("accn", ""), row.get("end", "")
                if accn and end and end > by_accn.get(accn, ""):
                    by_accn[accn] = end
    return sorted(set(by_accn.values()))


def _circular_day_gap(a: str, b: str) -> int:
    """兩個日期的「月日」相差幾天（跨年取最短路徑，0~182）。"""
    ya, ma, da = (int(x) for x in a.split("-"))
    yb, mb, db = (int(x) for x in b.split("-"))
    # 一律放到同一個非閏年上比，避免 2/29 與閏年造成的 1 天位移被當成漂移
    ref = 2001
    pa = date(ref, ma, min(da, 28) if ma == 2 else da).timetuple().tm_yday
    pb = date(ref, mb, min(db, 28) if mb == 2 else db).timetuple().tm_yday
    diff = abs(pa - pb)
    return min(diff, 365 - diff)


def _quarter_ends(facts: dict) -> list[str]:
    """每份 10-Q 的期末日（同一個 accession 取最大的 `end`），由舊到新。"""
    by_accn: dict[str, str] = {}
    us_gaap = facts.get("facts", {}).get("us-gaap", {})
    for anchor in _ANCHORS:
        for unit_rows in us_gaap.get(anchor, {}).get("units", {}).values():
            for row in unit_rows:
                if row.get("form") != "10-Q":
                    continue
                accn, end = row.get("accn", ""), row.get("end", "")
                if accn and end and end > by_accn.get(accn, ""):
                    by_accn[accn] = end
    return sorted(set(by_accn.values()))


# 財報發布日落後期末日幾天。200 份 Item 2.02 8-K 實測是 4~58 天（見
# `docs/8k-period-off-by-one.md`），中位數約 28。companyfacts 裡沒有 8-K 的發布日，
# 所以拿它當發布日的替身——重點是「同一個發布日，用不同 FYE 會算出不同的季」，
# 替身抓得準不準不影響這個結論。
_TYPICAL_ANNOUNCE_LAG = 28


def _quantify(ticker: str, latest_end: str, offenders: list[str]) -> None:
    """改制**以前**那些季，用 EDGAR 現值回推會標錯幾季。"""
    from datetime import timedelta

    sys.path.insert(0, str(ROOT / "src"))
    from fiscal_input import quarter_label_from_announcement as label_of

    facts = json.loads((CACHE / f"facts_{ticker}.json").read_text(encoding="utf-8"))
    cur_mmdd = latest_end[5:7] + latest_end[8:10]
    old_mmdd = offenders[-1][5:7] + offenders[-1][8:10]     # 改制前最後一個財年
    cutoff = offenders[-1]

    wrong, same, samples = 0, 0, []
    for end in _quarter_ends(facts):
        if end >= cutoff:
            continue
        y, m, d = (int(x) for x in end.split("-"))
        announce = (date(y, m, d) + timedelta(days=_TYPICAL_ANNOUNCE_LAG)).strftime("%Y%m%d")
        now, then = label_of(announce, cur_mmdd), label_of(announce, old_mmdd)
        if now == then:
            same += 1
            continue
        wrong += 1
        if len(samples) < 3:
            samples.append(f"期末 {end} → 現行規則 {now}，當時 FYE 應為 {then}")

    total = wrong + same
    if not total:
        return
    print(f"\n  ▸ {ticker} 影響量化（發布日用「期末日 + {_TYPICAL_ANNOUNCE_LAG} 天」代入）：")
    print(f"      改制前有 {total} 季，用 EDGAR 現值（{cur_mmdd}）回推會標錯 "
          f"{wrong} 季（{wrong / total * 100:.0f}%）")
    for line in samples:
        print(f"      {line}")


def main() -> int:
    files = sorted(CACHE.glob("facts_*.json"))
    if not files:
        print(f"找不到 {CACHE}/facts_*.json——先跑 scripts/spike_derive_mapping.py")
        return 1

    drifted: list[tuple[str, str, list[str]]] = []
    skipped: list[str] = []
    hist = defaultdict(int)

    for path in files:
        ticker = path.stem.removeprefix("facts_")
        try:
            ends = _fy_ends(json.loads(path.read_text(encoding="utf-8")))
        except Exception as exc:                      # 快取壞掉不該拖垮整批
            skipped.append(f"{ticker}:{type(exc).__name__}")
            continue
        if len(ends) < 2:
            skipped.append(f"{ticker}:只有 {len(ends)} 個年度")
            continue
        latest = ends[-1]
        offenders = [e for e in ends if _circular_day_gap(e, latest) > THRESHOLD]
        worst = max((_circular_day_gap(e, latest) for e in ends), default=0)
        hist[min(worst // 5 * 5, 60)] += 1
        if offenders:
            drifted.append((ticker, latest, offenders))

    n = len(files) - len(skipped)
    print(f"樣本 {n} 家（{CACHE} 的 companyfacts 快取，零網路請求）")
    print(f"門檻：歷年 10-K 期末日的月日與最新那年相差 > {THRESHOLD} 天\n")
    print("最大偏移分布（天）：")
    for bucket in sorted(hist):
        print(f"  {bucket:2d}~{bucket + 4:2d}  {'#' * hist[bucket]} {hist[bucket]}")

    print(f"\n=== 判定為改過財年：{len(drifted)} 家 ===")
    for ticker, latest, offenders in drifted:
        gaps = ", ".join(f"{e}(差{_circular_day_gap(e, latest)}天)" for e in offenders)
        print(f"  {ticker:6s} 最新財年結束 {latest}；偏離的年度：{gaps}")

    for ticker, latest, offenders in drifted:
        _quantify(ticker, latest, offenders)

    if skipped:
        print(f"\n略過 {len(skipped)} 家：{', '.join(skipped[:12])}"
              f"{' …' if len(skipped) > 12 else ''}")

    print("\n【怎麼讀這個結果】")
    print("  有列在上面的公司，**改制以前**的 Item 2.02 8-K 會被 B5 的零下載規則標錯，")
    print("  因為 EDGAR 只給現在的 fiscal_year_end。改制之後的申報不受影響。")
    print("  沒有列出來的公司，這個風險在這批樣本上不存在。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
