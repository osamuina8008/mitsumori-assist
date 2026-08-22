#!/bin/bash
# ダブルクリックで: 最新のツールをコピー→GitHubへpush→公開URLに反映
cd "$(dirname "$0")"
SRC="/Users/osamuinagaki/Library/CloudStorage/GoogleDrive-onextec.inagaki@gmail.com/マイドライブ/20_社内/20_見積・単価・原価/201_正式見積作成手順/新"
echo "── 最新版を取り込み中 ──"
cp "$SRC/見積算定書アシスト.html" index.html
cp "$SRC/0-1_炉組みアシスト.html" tool_0-1_炉組み.html
cp "$SRC/算定書生成.py" 算定書生成.py
cp "$SRC/算定書生成.command" 算定書生成.command
cp "$SRC/README_算定書アシスト.txt" README_使い方.txt
git add -A
if git diff --cached --quiet; then
  echo "変更なし（既に最新）。"
else
  git commit -m "更新 $(date +%Y-%m-%d_%H:%M)"
  git push && echo "✅ 公開URLに反映しました: https://osamuina8008.github.io/mitsumori-assist/"
fi
echo "ウィンドウは閉じて構いません。"
