"""Tests for i18n.py 與 locales/ — 多語言的三道防線。

1. 語言檔的 key 集合必須完全一致（漏翻當場紅燈）
2. placeholder 必須一致（譯文打錯 {name} 會讓 t() 靜默吐未格式化的字串）
3. src/ 底下不得再出現寫死的中日文字面（防止日後功能開發時悄悄退化）

第 3 條是**永久**的：它不是為了這次遷移，是為了讓下一次不會又回到 558 條
寫死字串的狀態。
"""

import ast
import re
from pathlib import Path

import pytest

import i18n

SRC = Path(__file__).resolve().parent.parent / "src"
CJK = re.compile(r"[一-鿿぀-ヿ]")
PLACEHOLDER = re.compile(r"\{[a-zA-Z_][a-zA-Z0-9_]*\}")

LANGS = [code for code, _, _ in i18n.LANGUAGES]


def _strings(lang: str) -> dict[str, str]:
    return i18n._strings(lang)


# ── 1. key 集合一致 ────────────────────────────────────────────────────────

def test_every_language_has_the_same_keys():
    """任一語言少一條就紅燈。

    新增語言時漏翻幾條是必然會發生的，靠人眼比對 341 條不可能可靠——
    這條測試就是那個「不可能可靠」的替代品。
    """
    base = set(_strings(i18n.FALLBACK_LANG))
    assert base, "繁中母表是空的，locale 載入壞了"
    for lang in LANGS:
        keys = set(_strings(lang))
        missing = sorted(base - keys)
        extra = sorted(keys - base)
        assert not missing, f"{lang} 少了 {len(missing)} 條：{missing[:10]}"
        assert not extra, f"{lang} 多了 {len(extra)} 條（母表沒有）：{extra[:10]}"


def test_no_language_table_is_empty():
    for lang in LANGS:
        assert _strings(lang), f"{lang} 的 STRINGS 是空的"


# ── 2. placeholder 一致 ───────────────────────────────────────────────────

def test_placeholders_match_across_languages():
    """譯文的 {name} 打錯或漏掉，t() 會 format 失敗並吐出未格式化的原字串——
    畫面上看到 `{ticker}` 這種殘留，不會 crash 所以特別容易漏掉。"""
    base = _strings(i18n.FALLBACK_LANG)
    for lang in LANGS:
        if lang == i18n.FALLBACK_LANG:
            continue
        table = _strings(lang)
        for key, zh in base.items():
            want = set(PLACEHOLDER.findall(zh))
            got = set(PLACEHOLDER.findall(table[key]))
            assert want == got, (
                f"{lang} / {key} 的 placeholder 不一致："
                f"母表 {sorted(want)}、譯文 {sorted(got)}"
            )


# ── 3. src/ 不得寫死中日文 ────────────────────────────────────────────────
#
# 豁免清單。每一條都要有理由——沒理由的豁免等於把這條測試關掉。
ALLOWLIST = {
    # 語言選單顯示名（「繁體中文」「日本語」）本來就該用各語言自稱，
    # 而且它們住在 i18n.py 自己身上，沒有更上層可以查。
    "i18n.py",
    # 2026-09-02 起 logs/app.log 的內容一律英文（見 docs/ARCHITECTURE.md
    # 「logs/app.log 的語言與格式」），所以 _write_log* 那些字面已經不是中文。
    # main.py 仍在名單上的原因只剩 UNCATEGORIZED（存進 config.json 的群組
    # 名稱，是資料不是介面文字）與少數維護者導向的主控台訊息。
    "main.py",
    # 給 skill 與維護者的開發者介面，不是產品 UI。輸出 Excel 的語言由
    # --lang 控制（見 cli.py 的 --lang 說明）。
    "cli.py",
    # 主控台診斷訊息，維護者導向，與 log 同一類。
    "fetcher_gaap.py",
    "fetcher_nongaap.py",
    # Data_NonGAAP 的版面。該功能停用中（main.NONGAAP_ENABLED = False），
    # 不產出任何 sheet，也沒有 golden 覆蓋。要遷移得連 A 欄機器鍵一起改，
    # 屬於另一次改動。見 docs/TODO.md E1。
    "nongaap_layout.py",
    "metric_rules.py",
}


def _hardcoded_cjk(path: Path) -> list[tuple[int, str]]:
    """回傳 (行號, 字串)。docstring 與註解不算——那些是寫給人看的說明。"""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    docs = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef,
                             ast.ClassDef)):
            body = node.body
            if (body and isinstance(body[0], ast.Expr)
                    and isinstance(body[0].value, ast.Constant)
                    and isinstance(body[0].value.value, str)):
                docs.add(id(body[0].value))
    hits = []
    for node in ast.walk(tree):
        if (isinstance(node, ast.Constant) and isinstance(node.value, str)
                and CJK.search(node.value) and id(node) not in docs):
            hits.append((node.lineno, node.value))
    return sorted(hits)


def _scannable() -> list[Path]:
    return [p for p in sorted(SRC.rglob("*.py"))
            if "locales" not in p.parts
            and "__pycache__" not in p.parts
            and p.name not in ALLOWLIST]


@pytest.mark.parametrize("path", _scannable(), ids=lambda p: p.name)
def test_no_hardcoded_cjk(path):
    """介面文字一律走 t()。

    這條擋的不是這次遷移，是**下一次**：新增功能時順手寫一個中文按鈕標籤
    最自然不過，沒有這條測試，三個月後就又回到全部寫死的狀態。

    真的需要豁免就加進 ALLOWLIST，但要寫清楚理由。
    """
    hits = _hardcoded_cjk(path)
    assert not hits, (
        f"{path.name} 有 {len(hits)} 條寫死的中日文字串，請改走 i18n.t()：\n"
        + "\n".join(f"  行 {ln}: {v[:60]!r}" for ln, v in hits[:10])
    )


# ── t() 的行為 ────────────────────────────────────────────────────────────

def test_unknown_key_returns_the_key_itself():
    """查不到不回空字串——空白按鈕看不見，key 看得見。"""
    i18n.set_lang("zh_tw")
    assert i18n.t("gui.btn.does_not_exist") == "gui.btn.does_not_exist"


def test_falls_back_to_traditional_chinese(monkeypatch):
    monkeypatch.setitem(i18n._cache, "ja", {})
    i18n.set_lang("ja")
    try:
        assert i18n.t("gui.btn.run") == _strings("zh_tw")["gui.btn.run"]
    finally:
        i18n._cache.pop("ja", None)
        i18n.set_lang("zh_tw")


def test_unknown_lang_falls_back_to_default():
    try:
        assert i18n.set_lang("kl_ingon") == i18n.DEFAULT_LANG
        assert i18n.set_lang(None) == i18n.DEFAULT_LANG
    finally:
        i18n.set_lang("zh_tw")


def test_format_failure_returns_the_unformatted_string():
    """譯文的 placeholder 打錯不該讓程式當掉。"""
    i18n.set_lang("zh_tw")
    got = i18n.t("gui.status.processing", ticker="AAPL")   # 少給 i / total
    assert "{i}" in got or "{total}" in got


def test_excel_font_follows_language():
    """微軟正黑體缺日文假名字形，日文必須換字型。"""
    try:
        assert i18n.excel_font("ja") != i18n.excel_font("zh_tw")
        assert i18n.excel_font("ja") == "Yu Gothic"
    finally:
        i18n.set_lang("zh_tw")


def test_language_menu_is_generated_from_the_registry():
    """下拉選單的選項不是寫死的——新增語言只改 LANGUAGES 一行。"""
    codes = [c for c, _ in i18n.available_languages()]
    assert codes == LANGS
    assert "zh_tw" in codes and "ja" in codes


# ── Excel 的機器鍵不得被翻譯 ─────────────────────────────────────────────

MACHINE_KEY_ROWS = [
    "Revenue", "Gross Profit", "Operating Income", "Net Income",
    "Cash", "Total Assets", "Operating Cash Flow", "Capex", "Free Cash Flow",
]


@pytest.mark.parametrize("lang", LANGS)
def test_machine_keys_are_never_translated(lang):
    """A 欄是下游跨檔案 MATCH 的生命線，任何語言下都必須是同一串英文。

    這裡驗的是「locale 檔裡沒有人不小心把 acct.Revenue 的 **key** 翻掉」——
    值可以翻，key 不行。
    """
    table = _strings(lang)
    for row in MACHINE_KEY_ROWS:
        assert f"acct.{row}" in table, f"{lang} 缺 acct.{row}"


@pytest.mark.parametrize("lang", LANGS)
def test_ratio_keys_keep_their_unit_suffix(lang):
    """單位後綴是 excel_formatter 判斷數字格式的依據，且優先於關鍵字判斷。

    後綴長在 key 上（＝Excel A 欄），不是值上——這條確認沒有人在整理
    locale 時把 key 的後綴順手拿掉。
    """
    for key in _strings(lang):
        if key.startswith("ratio."):
            assert key.endswith(("(%)", "(x)", "(days)", "($)")), key


# ── Watchlist 群組名稱：顯示與儲存必須分開 ────────────────────────────────
#
# 這是整個 i18n 改動裡唯一會**污染使用者資料**的地方，所以單獨釘住。
# 群組名稱寫進 config.json，程式到處在 `g["name"] == UNCATEGORIZED` 比對。
# 若顯示名直接被存回去，日文使用者新增公司時會建出一個叫「未分類」的
# 日文群組與既有的並存，股票分散在兩邊——而且畫面上兩個群組名稱長得一樣，
# 看起來只是「怎麼多了一個空群組」，不會有人聯想到是語言造成的。

def test_group_name_round_trips_in_every_language():
    import main

    for lang in LANGS:
        i18n.set_lang(lang)
        shown = main._group_display(main.UNCATEGORIZED)
        assert main._group_stored(shown) == main.UNCATEGORIZED, (
            f"{lang}: {shown!r} 換不回 {main.UNCATEGORIZED!r}，"
            f"存回 config.json 會建出重複群組"
        )
    i18n.set_lang("zh_tw")


def test_user_group_names_are_never_translated():
    """使用者自訂的群組名原樣進出——只有預設群組有譯名。"""
    import main

    for lang in LANGS:
        i18n.set_lang(lang)
        for name in ("半導體", "AI Infra", "watchlist-2026"):
            assert main._group_display(name) == name
            assert main._group_stored(name) == name
    i18n.set_lang("zh_tw")


def test_uncategorized_stored_value_is_still_the_original():
    """儲存值不可以改——改了等於讓所有既有 config.json 的群組失聯。"""
    import main

    assert main.UNCATEGORIZED == "未分類"


# ── 財年區間公式：譯文帶引號不可以把公式切碎 ──────────────────────────────

def test_fy_span_formula_survives_a_quote_in_the_translation(monkeypatch):
    """英文譯文出現撇號（`Company's FY`）完全可預期。

    舊寫法是 `repr()` 之後把單引號換成雙引號——字串內含單引號時 repr() 會
    自己改用雙引號包，那次 replace 就把公式切成 `="Fisc" l "&TEXT(...)`。
    Excel 開起來是 #NAME? 或乾脆拒絕開檔，而 Python 這邊一點錯都不會報。
    """
    import fiscal_input

    monkeypatch.setitem(i18n._cache, "zh_tw",
                        dict(_strings("zh_tw"),
                             **{"xls.fy_input.span_prefix": "Fisc'l \"FY\" "}))
    i18n.set_lang("zh_tw")
    try:
        f = fiscal_input._fy_span_formula()
        # Excel 的字串常值一律雙引號，內部的雙引號寫兩次 → 總數必為偶數
        assert f.count('"') % 2 == 0, f
        assert "'" not in f.split("&")[0] or f.startswith('=IF('), f
        # 撇號原樣留在字串裡，沒被當成語法
        assert "Fisc'l" in f
        # 內部的雙引號有被逸出成兩個
        assert '""FY""' in f
    finally:
        i18n._cache.pop("zh_tw", None)
        i18n.set_lang("zh_tw")


def test_fy_span_formula_uses_month_names_in_english():
    """英文要 `FY Oct – Sep`。`"m"` 只會給你 `FY 10 – 9`——月份格式碼必須
    跟著語言走，不能寫死在公式裡。"""
    import fiscal_input

    i18n.set_lang("en")
    try:
        assert '"mmm"' in fiscal_input._fy_span_formula()
    finally:
        i18n.set_lang("zh_tw")

    i18n.set_lang("zh_tw")
    assert '"m"' in fiscal_input._fy_span_formula()


@pytest.mark.parametrize("lang", LANGS)
def test_fy_span_formula_is_balanced_in_every_language(lang):
    import fiscal_input

    i18n.set_lang(lang)
    try:
        f = fiscal_input._fy_span_formula()
        assert f.startswith("=IF(")
        assert f.count("(") == f.count(")"), f
        assert f.count('"') % 2 == 0, f
    finally:
        i18n.set_lang("zh_tw")


# ── B 欄查表：模板路徑與非模板路徑必須一致 ────────────────────────────────

def test_template_path_uses_the_same_column_b_lookup():
    """設了 template_path 時走 `_write_sheet_template`，那條路徑一度還在直接
    呼叫 `zh_label()`——Data_Ratios / Data_Meta 的 A 欄鍵住在 ratio.* / meta.*，
    用 acct.* 查一定落空，B 欄整欄空白，而且只有設了模板的人才會遇到。"""
    import inspect

    import excel_writer

    src = inspect.getsource(excel_writer._write_sheet_template)
    assert "_col_b(" in src, "模板路徑沒走 _col_b，B 欄命名空間會查錯"
    assert "zh_label(" not in src, "模板路徑還在直接呼叫 zh_label"
