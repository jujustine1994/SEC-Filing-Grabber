"""跨公司比較視窗的小工具函式（2026-09-02 CTH 回報的兩件事）。

1. 已選公司的 chip 只顯示 ticker——原本是 `NVDA NVIDIA CORP`，三家就把
   兩行塞爆，而公司全名在選的當下已經在建議清單看過了
2. `pack_wrapped_chips()` 的換行位置會歪掉：chip 建成**外層容器**的子元件、
   再用 `in_=rows[-1]` 塞進行 frame，父子關係與排版容器不一致，第二行以後
   的 chip 位置就跑掉（CTH 截圖：第三個 chip 跑到視窗右邊）

耗時格式化（log 顯示「這次跑了幾分幾秒」）也在這裡測——它是純函式。
"""
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
    沒有辦法用假物件測。沒有顯示裝置的環境（headless CI）直接跳過。"""
    tkinter = pytest.importorskip("tkinter")
    try:
        root = tkinter.Tk()
    except tkinter.TclError as exc:            # 沒有 display
        pytest.skip(f"no display: {exc}")
    root.withdraw()
    yield root
    root.destroy()


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
