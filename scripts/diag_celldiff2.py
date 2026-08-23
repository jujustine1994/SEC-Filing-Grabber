"""diag_celldiff2.py — 兩份答案卷快取的逐格回歸比對。

用法：
    venv/Scripts/python.exe scripts/diag_celldiff2.py <改動前的目錄> <改動後的目錄>

改任何 concept 對照之前先把 `output/_spike` 複製一份當基準，改完重建再比。
驗收標準：**不能有任何一格從「有值」變成「不同的值」或「空」**。

⚠ **鍵一定要用 `(列名, 第幾次出現)`。** `Net Income` 與 `SBC` 在 IS 和 CF 模板
各有一列，用列名當字典鍵只會留最後一個，會變成拿 IS 那列去比 CF 那列——
2026-08-24 這樣憑空生出 3,659 個假異動、一度誤判成嚴重回歸。overflow 區的
`Other`／`Accrued expenses` 這類重複列名更嚴重。這支已經處理好了。
"""
import pickle, pathlib, sys
from collections import Counter
sys.stdout.reconfigure(encoding="utf-8")
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
TPL = {r[0] for T in (IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE) for r in T}
OLD, NEW = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])

def keyed(g):
    seen, out = Counter(), {}
    for i, n in enumerate(g["concepts"]):
        seen[n] += 1
        out[(n, seen[n])] = g["values"][i]
    return out

stat = Counter(); tpl_bad = []; gained = Counter(); by_co = Counter()
for p in sorted(NEW.glob("gaap_*.pkl")):
    tk = p.stem[5:]; o = OLD / p.name
    if not o.exists(): continue
    a, b = pickle.loads(o.read_bytes()), pickle.loads(p.read_bytes())
    if a["ends"] != b["ends"]: stat["期間軸不同"] += 1; continue
    ka, kb = keyed(a), keyed(b)
    for key, bv in kb.items():
        if key not in ka: continue
        kind = "模板" if key[0] in TPL else "overflow"
        for i, (x, y) in enumerate(zip(ka[key], bv)):
            hx, hy = isinstance(x,(int,float)), isinstance(y,(int,float))
            if hx and hy and x != y:
                stat[f"{kind}:值變了"] += 1; by_co[tk] += 1
                if kind == "模板": tpl_bad.append((tk, key[0], a["ends"][i], x, y))
            elif hx and not hy:
                stat[f"{kind}:變空"] += 1; by_co[tk] += 1
                if kind == "模板": tpl_bad.append((tk, key[0], a["ends"][i], x, None))
            elif hy and not hx:
                stat[f"{kind}:補上"] += 1
                if kind == "模板": gained[key[0]] += 1
for k, v in sorted(stat.items()): print(f"  {k:20} {v}")
print(f"\n模板列補上（空 -> 有值）：")
for n, c in gained.most_common(12): print(f"   {n:32} +{c}")
print(f"\n模板列的回歸（共 {len(tpl_bad)} 格）：")
for n, c in Counter(x[1] for x in tpl_bad).most_common(12): print(f"   {n:32} {c}")
print(f"\n異動最多的公司：{', '.join(f'{t}({n})' for t, n in by_co.most_common(8))}")
