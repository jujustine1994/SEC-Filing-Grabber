"""視窗開啟位置（CTH 2026-08-17 回報）。

症狀：視窗很高，而且起始位置往下切到螢幕最下面，底部被工作列蓋掉。
成因：`__init__` 只呼叫 `geometry("700x650")` 沒給座標，位置完全交給
Windows 決定（會沿用上一個視窗的階梯式落點，越開越低）。

這裡測的是座標算式本身——真正會出錯的是「算出來的視窗有沒有跑出工作區」，
不是 tkinter 版面，所以不開視窗（沿用 test_first_run_language.py 的作法）。

工作區 = 螢幕扣掉工作列後的矩形，用 (left, top, right, bottom) 表示。
多螢幕時 left/top 可能是負的，算式不能假設從 0 開始。
"""

import pytest

import main

# (left, top, right, bottom)
FHD = (0, 0, 1920, 1040)          # 1920x1080 扣掉底部 40px 工作列
LAPTOP = (0, 0, 1366, 728)        # 1366x768 扣掉工作列
TINY = (0, 0, 800, 600)           # 比視窗還小
LEFT_MONITOR = (-1920, 0, 0, 1040)  # 副螢幕在主螢幕左邊，座標為負
TASKBAR_ON_TOP = (0, 48, 1920, 1080)


def _parse(geom: str) -> tuple[int, int, int, int]:
    """'900x680+510+126' -> (900, 680, 510, 126)，支援負座標。"""
    size, rest = geom.split("+", 1) if "+" in geom else (geom, "0+0")
    w, h = (int(v) for v in size.split("x"))
    # 負座標長這樣：'900x680+-1510+126'
    parts = rest.split("+")
    x, y = int(parts[0]), int(parts[1])
    return w, h, x, y


ALL_AREAS = [FHD, LAPTOP, TINY, LEFT_MONITOR, TASKBAR_ON_TOP]


# ── 核心保證：永遠不跑出工作區 ────────────────────────────────────────────

@pytest.mark.parametrize("area", ALL_AREAS)
@pytest.mark.parametrize("want", [(900, 680), (700, 650), (900, 900), (1200, 1200)])
def test_window_never_falls_outside_the_work_area(area, want):
    """這是 CTH 回報的症狀本身：視窗下緣掉到工作列底下看不到。
    不管螢幕多小、視窗要多大，四個邊都必須落在工作區內。"""
    left, top, right, bottom = area
    w, h, x, y = _parse(main.fit_geometry(want[0], want[1], area))
    assert x >= left, f"左緣 {x} 超出工作區 {left}"
    assert y >= top, f"上緣 {y} 超出工作區 {top}"
    assert x + w <= right, f"右緣 {x + w} 超出工作區 {right}"
    assert y + h <= bottom, f"下緣 {y + h} 超出工作區 {bottom}——這就是被工作列切掉"


@pytest.mark.parametrize("area", ALL_AREAS)
def test_oversized_window_is_shrunk_to_fit(area):
    """視窗比工作區還大時要縮小，不是硬擺著讓它出界。"""
    left, top, right, bottom = area
    w, h, _, _ = _parse(main.fit_geometry(5000, 5000, area))
    assert w == right - left
    assert h == bottom - top


# ── 置中 ──────────────────────────────────────────────────────────────────

def test_window_is_horizontally_centred():
    w, _, x, _ = _parse(main.fit_geometry(900, 680, FHD))
    left_gap, right_gap = x, 1920 - (x + w)
    assert abs(left_gap - right_gap) <= 1


def test_window_sits_slightly_above_vertical_centre():
    """正中央看起來偏低。略高於中線比較自然，但不可以貼齊上緣。"""
    _, h, _, y = _parse(main.fit_geometry(900, 680, FHD))
    centred_y = (1040 - h) // 2
    assert 0 < y < centred_y


def test_offsets_are_relative_to_the_work_area_not_the_screen():
    """工作列在上方時，視窗不能從 y=0 開始算，否則標題列被蓋住。"""
    _, _, _, y = _parse(main.fit_geometry(900, 680, TASKBAR_ON_TOP))
    assert y >= 48


def test_negative_monitor_coordinates_are_preserved():
    """副螢幕在主螢幕左側時工作區座標是負的，不可以被夾成 0。"""
    w, _, x, _ = _parse(main.fit_geometry(900, 680, LEFT_MONITOR))
    assert x < 0
    assert x + w <= 0


# ── 工作區取得 ────────────────────────────────────────────────────────────

def test_work_area_is_a_sane_rectangle():
    """實機取值：不論走 ctypes 還是退回 winfo_screen*，都要是有效矩形。"""
    left, top, right, bottom = main.work_area()
    assert right > left
    assert bottom > top
