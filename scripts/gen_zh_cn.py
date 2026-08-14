# -*- coding: utf-8 -*-
"""從 src/locales/zh_tw.py 重新產生 src/locales/zh_cn.py。

**一次性工具，平常不需要跑。** zh_cn.py 產完就是一個普通的 Python 檔，
要改用詞直接改它即可；只有在 zh_tw 大量新增條目、想一次補齊简中時才用這支。

需要 OpenCC（**不在 requirements.txt**，執行期用不到）：

    uv pip install opencc-python-reimplemented --python venv\\Scripts\\python.exe
    ./venv/Scripts/python.exe scripts/gen_zh_cn.py
    uv pip uninstall opencc-python-reimplemented --python venv\\Scripts\\python.exe

⚠ 會**覆蓋** zh_cn.py。手動改過的用詞會不見——先確認你要的都進了下面的
PHRASE 表再跑。

用 `tw2s`（只轉字形）而不是 `tw2sp`（連詞彙一起換）：tw2sp 的詞庫太激進，
「進階設定」會被換成「高端设置」、「執行」變「运行」，那是它猜的產品用語
不是翻譯。詞彙差異改由 PHRASE 表逐條指定，看得到也控制得住。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

try:
    import opencc
except ImportError:
    sys.exit("需要 OpenCC，安裝方式見本檔開頭的 docstring")

import locales.zh_tw as zh_tw       # noqa: E402

cc = opencc.OpenCC("tw2s")

# OpenCC 之後的專案專屬修正。只列真的會誤讀或明顯不地道的，
# 兩岸共通的財務科目（營業收入、毛利率…）不動。
PHRASE = {
    "资讯": "信息",
    "软体": "软件",
    "硬体": "硬件",
    "档案": "文件",
    "资料夹": "文件夹",
    "预设": "默认",
    "储存位置": "保存位置",
    "储存关闭": "保存关闭",
    "网路": "网络",
    "连线": "连接",
    "程式": "程序",
    "支股票": "只股票",
    "间公司": "家公司",
    "萤幕": "屏幕",
    "点选": "点击",
    "汇入": "导入",
    "汇出": "导出",
}

# {placeholder} 內若混到中文會被翻掉，先擋下來再放回去。
PLACEHOLDER = re.compile(r"\{[^}]*\}")


def convert(s: str) -> str:
    holes = PLACEHOLDER.findall(s)
    out = cc.convert(PLACEHOLDER.sub("\x00", s))
    for a, b in PHRASE.items():
        out = out.replace(a, b)
    for h in holes:
        out = out.replace("\x00", h, 1)
    return out


def main() -> None:
    rows = {k: convert(v) for k, v in zh_tw.STRINGS.items()}

    # placeholder 必須完全一致，否則 t() 會 format 失敗
    for k, v in zh_tw.STRINGS.items():
        assert PLACEHOLDER.findall(v) == PLACEHOLDER.findall(rows[k]), k

    header = '''"""locales/zh_cn.py — 简体中文

由 zh_tw.py 经 OpenCC (tw2s) 生成，见 scripts/gen_zh_cn.py。
生成后这就是一个普通的 Python 文件，直接改即可。

key 命名空间见 locales/__init__.py。改这里的译文不影响任何逻辑：
程序一律用 key 比对，Excel 一律用 A 栏英文机器键比对。
"""

from __future__ import annotations

STRINGS: dict[str, str] = {
'''
    body = "\n".join(f"    {k!r}: {v!r}," for k, v in rows.items())
    (ROOT / "src" / "locales" / "zh_cn.py").write_text(
        header + body + "\n}\n", encoding="utf-8", newline="")
    print(f"zh_cn.py: {len(rows)} keys")


if __name__ == "__main__":
    main()
