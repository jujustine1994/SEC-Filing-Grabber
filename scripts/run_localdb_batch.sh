#!/usr/bin/env bash
# run_localdb_batch.sh — 分段跑「更新本地庫」，每段一個獨立 process。
#
# 為什麼要分段：2026-09-06 實測，一個 process 連跑 67 家在第 16 家被系統
# 因記憶體不足中止（edgartools 的內部快取會跨公司累積，我們自己的
# `_parse_cache_scope()` 只涵蓋單次抓取）。每段跑完 process 結束，記憶體
# 整個還給系統，下一段重新開始。
#
# 中止不會白費：`save_filing()` 逐份即時落檔，已完成的公司下次會整家跳過。
#
#   bash scripts/run_localdb_batch.sh <輸出前綴> <每段幾家> <ticker...>
#
# 例：bash scripts/run_localdb_batch.sh output/_localdb/batch2 8 LVS MA MAR ...
set -u
PREFIX="$1"; SIZE="$2"; shift 2
PY="./venv/Scripts/python.exe"
export PYTHONIOENCODING=utf-8

ALL=("$@")
TOTAL=${#ALL[@]}
CHUNK=0
for ((i=0; i<TOTAL; i+=SIZE)); do
  CHUNK=$((CHUNK+1))
  PART=("${ALL[@]:i:SIZE}")
  echo "===== 第 $CHUNK 段（$((i+1))-$((i+${#PART[@]})) / $TOTAL）：${PART[*]} ====="
  "$PY" -u src/cli.py update-db "${PART[@]}" \
        --json "${PREFIX}_chunk$(printf '%02d' "$CHUNK").json" 2>&1 \
    | grep -vE "^(No XBRL attachments|Failed to resolve)"
  echo "----- 第 $CHUNK 段結束（exit ${PIPESTATUS[0]}）-----"
done
echo "全部 $CHUNK 段跑完"
