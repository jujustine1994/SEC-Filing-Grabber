#!/usr/bin/env bash
# watchdog_h0_baseline.sh — H0 基線重建的無人值守看門狗（2026-09-04 夜間作業用）
#
# 為什麼要這支：201 家答案卷重建要跑好幾個小時，而 AI session 可能中途因為
# 額度用盡斷掉。這支是 detached 的 shell process，跟 AI 的額度無關——它負責
# 「抓取跑完 → 自動產生基線 → 把驗收數字撈出來寫成摘要檔」，確保成果落在
# 硬碟上，下一個 session 只要讀檔就能接手寫報告。
#
# 用法（背景執行，不要前景卡住）：
#   nohup bash scripts/watchdog_h0_baseline.sh > output/_spike/watchdog.log 2>&1 &
#
# 產出：
#   docs/template-coverage-baseline-<日期>.md   新基線
#   output/_spike/h0_summary.txt                晨間摘要要的關鍵數字
#   output/_spike/watchdog.log                  這支自己的 log

set -u
ROOT="C:/Users/CTH/Documents/Code/SEC Financial Tools"
cd "$ROOT" || exit 1
PY="./venv/Scripts/python.exe"
LOG="output/_spike/rebuild_20260904.log"
SUM="output/_spike/h0_summary.txt"

say() { echo "[watchdog $(date '+%m-%d %H:%M:%S')] $*"; }

# ── 1) 等抓取跑完 ─────────────────────────────────────────────────────────
# 完成判定：log 出現最後一行的「完整候選：」。
# 死亡判定：log 超過 20 分鐘沒被碰過（單家最慢實測約 2 分鐘，20 分鐘夠寬）。
say "開始等待 201 家重建"
while true; do
  n=$(ls output/_spike/gaap_*.pkl 2>/dev/null | wc -l)
  if grep -q "完整候選" "$LOG" 2>/dev/null; then
    say "重建完成，pkl=$n"
    break
  fi
  age=$(( $(date +%s) - $(stat -c %Y "$LOG") ))
  if [ "$age" -gt 1200 ]; then
    say "log 已 ${age}s 沒有更新（pkl=$n/201）——判定 process 已死，仍繼續產基線"
    break
  fi
  say "進行中 pkl=$n/201"
  sleep 300
done

# ── 2) D11 缺漏偵測：哪幾家抓取時噴了警告 ────────────────────────────────
# `_load_gaap()` 不管有沒有缺漏都會寫 pkl，壞掉的結果會被凍進去。
# 把可疑的 ticker 撈出來給人看，**不自動重跑**——要刪哪幾家的 pkl 由人決定。
say "掃描抓取警告"
grep -nE "warning|警告|Traceback|Error" "$LOG" > output/_spike/rebuild_warnings.txt 2>/dev/null
say "警告行數 $(wc -l < output/_spike/rebuild_warnings.txt)"

# ── 3) 產基線（不打網路，幾分鐘）─────────────────────────────────────────
say "產生基線"
PYTHONIOENCODING=utf-8 $PY scripts/gen_template_coverage_baseline.py
say "基線產生完畢（exit $?）"

# ── 4) 撈晨間摘要要的數字 ────────────────────────────────────────────────
NEW=$(ls -t docs/template-coverage-baseline-*.md | head -1)
OLD="docs/template-coverage-baseline-2026-08-24.md"
{
  echo "=== H0 基線重建摘要（$(date '+%Y-%m-%d %H:%M')）==="
  echo "新基線：$NEW"
  echo "pkl 家數：$(ls output/_spike/gaap_*.pkl 2>/dev/null | wc -l) / 201"
  echo
  echo "--- 達標列數 ---"
  echo "舊（08-24）：$(grep -o '的列：[0-9]* / [0-9]*' "$OLD")"
  echo "新（今天）：$(grep -o '的列：[0-9]* / [0-9]*' "$NEW")"
  echo
  echo "--- H1 驗收：from_ytd 列的 facts 填滿率 ---"
  grep -o 'facts 填滿率中位數：[0-9]*%' "$NEW"
  echo "（H1 記錄的原始症狀是約 25%）"
  echo
  echo "--- 三分類（我們抓到 / 真缺口 / 公司真的沒有）---"
  grep -A 6 '| 分類 | 格數 |' "$NEW"
  echo
  echo "--- 假警報（標紅家次）---"
  grep -A 6 '標紅：矛盾' "$NEW"
  echo
  echo "--- 兩份基線是否逐位元組相同（相同代表 pkl 沒重建成功）---"
  if cmp -s "$OLD" "$NEW"; then echo "⚠ 相同——有問題"; else echo "不同（正常）"; fi
} > "$SUM" 2>&1
say "摘要寫入 $SUM"
cat "$SUM"
say "結束"
