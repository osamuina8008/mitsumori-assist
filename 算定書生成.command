#!/bin/bash
# ダブルクリックで実行: ~/Downloads の最新 算定書_*.json を 0-2 テンプレに流し込み xlsx を生成
cd "$(dirname "$0")"
echo "── 見積算定書 生成 ──"
PY=$(command -v python3 || echo /usr/bin/python3)
"$PY" 算定書生成.py
echo ""
echo "完了。ウィンドウは閉じて構いません。"
