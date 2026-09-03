"""跨公司比較視窗的小工具函式（2026-09-02 CTH 回報的兩件事）。

1. 已選公司的 chip 只顯示 ticker——原本是 `NVDA NVIDIA CORP`，三家就把
   兩行塞爆，而公司全名在選的當下已經在建議清單看過了
2. `pack_wrapped_chips()` 的換行位置會歪掉：chip 建成**外層容器**的子元件、
   再用 `in_=rows[-1]` 塞進行 frame，父子關係與排版容器不一致，第二行以後
   的 chip 位置就跑掉（CTH 截圖：第三個 chip 跑到視窗右邊）

耗時格式化（log 顯示「這次跑了幾分幾秒」）也在這裡測——它是純函式。
"""
import time

import pytest

import main


# ── 耗時格式 ────────────────────────────────────────────────────────────────
#
# log 檔一律英文（2026-09-02 CTH 決定：log 是給 AI 與維護者除錯用的，
# 而且這個檔同時被 PowerShell 與 Python 寫，全英文可以整類避開 cp950 的
# 編碼地雷）。單位符號在中文介面上也讀得懂，所以畫面與 log 共用同一個格式，
# 不維護兩套。

@pytest.mark.parametrize("seconds, expected", [
    (0, "0s"),
    (7, "7s"),
    (59, "59s"),
    (60, "1m 00s"),
    (114, "1m 54s"),        # logs/app.log 實際出現過的 NVDA 那筆
    (599, "9m 59s"),
    (3599, "59m 59s"),
    (3600, "1h 00m 00s"),
    (3661, "1h 01m 01s"),
    (7325, "2h 02m 05s"),
])
def test_format_elapsed(seconds, expected):
    assert main.format_elapsed(seconds) == expected


def test_format_elapsed_accepts_floats_and_rounds_down():
    """`time.time()` 的差是浮點數，不要讓 log 出現 `1m 53.99999s`。"""
    assert main.format_elapsed(113.9) == "1m 53s"


def test_format_elapsed_never_returns_a_negative():
    """系統時間被調整過時寧可顯示 0s，也不要吐出 `-1m -3s` 這種東西。"""
    assert main.format_elapsed(-5) == "0s"


# ── 已選公司的 chip 文字 ────────────────────────────────────────────────────

def test_company_chip_entries_show_only_the_ticker():
    """CTH 2026-09-02：chip 顯示 `NVDA NVIDIA CORP` 太長，三家就換行。"""
    selected = [("NVDA", "NVIDIA CORP"), ("MSFT", "MICROSOFT CORP")]
    assert main.company_chip_entries(selected) == [("NVDA", "NVDA"), ("MSFT", "MSFT")]


def test_company_chip_entries_keeps_order_and_handles_missing_names():
    selected = [("INTC", ""), ("AMD", "ADVANCED MICRO DEVICES INC")]
    assert main.company_chip_entries(selected) == [("INTC", "INTC"), ("AMD", "AMD")]


# ── chip 換行 ───────────────────────────────────────────────────────────────

@pytest.fixture
def tk_root():
    """真的開一個 Tk root——`pack_wrapped_chips()` 要量 widget 實際寬度，
    沒有辦法用假物件測。沒有顯示裝置的環境（headless CI）直接跳過。

    ⚠ 重試三次不是防禦性程式碼：同一個 pytest session 內連續建/毀 Tk root，
    Windows 上偶爾會有一次 `TclError`（實測 2026-09-02，同一批測試跑兩次，
    skip 的是不同那條）。不重試的話會變成「隨機少跑一條測試」，而且因為它是
    skip 不是 fail，看起來完全正常——這種抖動比失敗更難發現。
    """
    tkinter = pytest.importorskip("tkinter")
    last_exc = None
    for _ in range(3):
        try:
            root = tkinter.Tk()
        except tkinter.TclError as exc:
            last_exc = exc
            time.sleep(0.2)
            continue
        root.withdraw()
        yield root
        root.destroy()
        return
    pytest.skip(f"no display after 3 attempts: {last_exc}")


def _chips(container):
    """容器底下所有 chip（chip 是 Frame，行也是 Frame——chip 有子元件，行沒有
    直接的 Label/Button，用「有沒有 Label 子元件」分辨）。"""
    found = []
    for row in container.winfo_children():
        for child in row.winfo_children():
            if any(w.winfo_class() == "TLabel" for w in child.winfo_children()):
                found.append(child)
    return found


def test_chips_belong_to_the_row_that_lays_them_out(tk_root):
    """每個 chip 的**父元件**必須就是它所在的那一行。

    原本 chip 的父元件是外層容器、只用 `in_=` 借行 frame 排版，第二行以後
    的位置會歪掉（CTH 截圖：`INTC` 那個 chip 跑到右邊）。
    """
    import tkinter.ttk as ttk

    container = ttk.Frame(tk_root)
    container.pack()
    entries = [(f"TICK{i}", f"TICK{i}") for i in range(8)]
    rows = main.pack_wrapped_chips(container, entries, lambda key: None, max_width=200)

    assert rows > 1, "測試前提：這組 chip 一定要撐到換行，否則測不到父子關係"
    chips = _chips(container)
    assert len(chips) == len(entries)
    row_frames = {str(r) for r in container.winfo_children()}
    for chip in chips:
        assert str(chip.master) in row_frames, (
            f"chip {chip} 的父元件是 {chip.master}，不是它所在的那一行")


def test_chips_wrap_onto_one_row_when_they_fit(tk_root):
    import tkinter.ttk as ttk

    container = ttk.Frame(tk_root)
    container.pack()
    rows = main.pack_wrapped_chips(container, [("AMD", "AMD")], lambda key: None,
                                   max_width=2000)
    assert rows == 1
    assert len(_chips(container)) == 1


def test_pack_wrapped_chips_clears_previous_chips(tk_root):
    """重畫時舊 chip 要清乾淨，不能疊上去。"""
    import tkinter.ttk as ttk

    container = ttk.Frame(tk_root)
    container.pack()
    main.pack_wrapped_chips(container, [("A", "A"), ("B", "B")], lambda key: None,
                            max_width=2000)
    main.pack_wrapped_chips(container, [("C", "C")], lambda key: None, max_width=2000)
    labels = [w.cget("text") for chip in _chips(container)
              for w in chip.winfo_children() if w.winfo_class() == "TLabel"]
    assert labels == ["C"]


# ── ticker 候選清單的位置與顯示時機 ────────────────────────────────────────
#
# 2026-09-02 CTH 連續回報兩次：① 沒輸入時它是一個常駐的大白框；② 改成動態顯示
# 之後，清單跑到整個視窗最底下（`pack()` 預設排到父容器最後面）。這條測試把
# 「打字→清單出現在輸入框正下方；清空→收起來」整段釘住。

@pytest.fixture
def compare_window(tk_root):
    """真的把「選擇比較內容」視窗建起來。tkinter 版面沒有辦法用假物件驗——
    這兩個 bug 都是版面問題，不開視窗就測不到。"""
    import main as main_mod

    app = main_mod.SECFetcherApp(tk_root)
    app.compare_selected_tickers = [("NVDA", "NVIDIA CORP")]
    app._open_compare_selection_window()
    tk_root.update_idletasks()
    import tkinter
    win = [w for w in tk_root.winfo_children() if isinstance(w, tkinter.Toplevel)][-1]
    yield win
    win.destroy()


def _walk(widget):
    for child in widget.winfo_children():
        yield child
        yield from _walk(child)


def test_ticker_suggestions_hidden_until_something_matches(tk_root, compare_window):
    listbox = next(w for w in _walk(compare_window) if w.winfo_class() == "Listbox")
    entry = next(w for w in _walk(compare_window) if w.winfo_class() == "TEntry")
    var = entry.cget("textvariable")

    assert not listbox.winfo_manager(), "沒輸入時候選清單不該占版面"

    tk_root.setvar(var, "INTEL")
    tk_root.update_idletasks()
    assert listbox.winfo_manager(), "打了字、有比中就該顯示"
    assert listbox.size() > 0

    tk_root.setvar(var, "")
    tk_root.update_idletasks()
    assert not listbox.winfo_manager(), "清空後要收起來"


def test_ticker_suggestions_appear_right_below_the_input(tk_root, compare_window):
    """`pack()` 預設排到父容器**最後面**——動態顯示時不指定 `after=`，
    清單會掉到整個視窗最底下（CTH 截圖：跑到「快照時間點」下面）。"""
    listbox = next(w for w in _walk(compare_window) if w.winfo_class() == "Listbox")
    entry = next(w for w in _walk(compare_window) if w.winfo_class() == "TEntry")

    tk_root.setvar(entry.cget("textvariable"), "INTEL")
    tk_root.update_idletasks()

    slaves = listbox.master.pack_slaves()
    ticker_row = next(w for w in slaves if entry in w.winfo_children())
    assert slaves.index(listbox) == slaves.index(ticker_row) + 1, (
        "候選清單必須緊跟在輸入框那一列後面，"
        f"實際排在第 {slaves.index(listbox)} 個、輸入框在第 {slaves.index(ticker_row)} 個")


# ── 快取命中數的 log 行（2026-09-03）────────────────────────────────────────
#
# 耗時變快也可能只是那天 SEC 比較順。沒有命中數字就無法判斷這次有沒有吃到快取。
# `logs/app.log` 一律英文（2026-09-02 起的既有規則）。

def test_cache_log_line_reports_hits_over_total():
    assert main.cache_log_line("NVDA", 24, 25) == "NVDA cache 24/25"


def test_cache_log_line_is_skipped_when_nothing_was_processed():
    """沒抓任何 filing 時不要在 log 留一行 `cache 0/0` 的雜訊。"""
    assert main.cache_log_line("NVDA", 0, 0) is None


def test_cache_log_line_is_english_only():
    """log 的讀者是維護者與 AI，而且這個檔同時被 PowerShell 寫，全英文
    可以整類避開 cp950 的編碼地雷。"""
    line = main.cache_log_line("NVDA", 0, 25)
    assert all(ord(ch) < 128 for ch in line)


# ── 快取容量顯示 ────────────────────────────────────────────────────────────
#
# CTH 選的是手動清、不做自動上限，所以「現在到底佔多少」是他做決定的依據，
# 這個數字必須一眼看得懂——不是 41104179 這種原始位元組數。

@pytest.mark.parametrize("num_bytes, expected", [
    (0, "0 KB"),
    (512, "0.5 KB"),
    (61234, "59.8 KB"),
    (18_400_000, "17.5 MB"),
    (1_073_741_824, "1.0 GB"),
])
def test_format_size(num_bytes, expected):
    assert main.format_size(num_bytes) == expected


def test_format_size_never_shows_a_negative():
    assert main.format_size(-1) == "0 KB"


# ── 抓取進行中不可以邊寫邊刪 ────────────────────────────────────────────────

def test_cache_buttons_are_locked_while_a_fetch_is_running():
    """Tab1／批次／跨公司比較任一個 worker thread 還在跑時，兩顆清除按鈕都要
    disable——不然會邊寫邊刪同一個 ticker 的資料夾。沿用專案既有的
    「執行中鎖住相關按鈕」慣例。"""
    assert main.cache_buttons_state(True) == "disabled"
    assert main.cache_buttons_state(False) == "normal"
