#!/bin/bash
cd "$(dirname "$0")"

echo "================================"
echo "  報告書自動出力ツール"
echo "================================"
echo ""

DEBUG_PROFILE="$HOME/Library/Application Support/Google/ChromeDebug"

# リモートデバッグ付きChromeが既に起動しているかチェック
if curl -s http://localhost:9222/json/version > /dev/null 2>&1; then
    echo "✓ リモートデバッグ付きChromeを検出しました"
else
    echo "⚠️  既にChromeが開いている場合は先に閉じてください (Cmd+Q)"
    echo ""
    read -p "準備ができたらEnterキーを押してください..."
    echo ""

    # デバッグ用プロファイルがなければ既存プロファイルをコピー
    if [ ! -d "$DEBUG_PROFILE" ]; then
        echo "初回セットアップ: プロファイルをコピー中..."
        cp -R "$HOME/Library/Application Support/Google/Chrome" "$DEBUG_PROFILE"
        echo "コピー完了"
    fi

    echo "Chromeをリモートデバッグモードで起動します..."
    /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
        --remote-debugging-port=9222 \
        --user-data-dir="$DEBUG_PROFILE" \
        --profile-directory=Default &

    sleep 5
    echo "Chromeが起動しました。"
fi

echo ""
echo "=== 次の手順 ==="
echo "  1. Chromeでログインしてください"
echo "  2. 「介護報告書一括印刷」画面を開いてください"
echo "  3. 準備ができたらこのウィンドウに戻ってください"
echo ""

# -------------------------------------------------------
#  対象年月の入力
# -------------------------------------------------------
# デフォルト値: 先月
DEFAULT_YEAR=$(date -v-1m +%Y)
DEFAULT_MONTH=$(date -v-1m +%-m)

echo "=========================================="
echo "  対象年月を入力してください"
echo "=========================================="
echo ""

# 年の入力
while true; do
    read -p "  対象年 [${DEFAULT_YEAR}]: " INPUT_YEAR
    INPUT_YEAR=${INPUT_YEAR:-$DEFAULT_YEAR}

    # 数字チェック
    if [[ "$INPUT_YEAR" =~ ^[0-9]{4}$ ]]; then
        break
    else
        echo "  ⚠️  4桁の西暦で入力してください（例: 2026）"
    fi
done

# 月の入力
while true; do
    read -p "  対象月 [${DEFAULT_MONTH}]: " INPUT_MONTH
    INPUT_MONTH=${INPUT_MONTH:-$DEFAULT_MONTH}

    # 数字チェック & 範囲チェック
    if [[ "$INPUT_MONTH" =~ ^[0-9]{1,2}$ ]] && [ "$INPUT_MONTH" -ge 1 ] && [ "$INPUT_MONTH" -le 12 ]; then
        break
    else
        echo "  ⚠️  1〜12の数字で入力してください"
    fi
done

echo ""
echo "  → ${INPUT_YEAR}年${INPUT_MONTH}月分 の報告書を出力します"
echo ""
read -p "よろしければEnterキーを押してください（中止: Ctrl+C）..."

echo ""
echo "💡 途中で止めたい場合は Ctrl+C を押してください"
echo "   （進捗は自動保存され、次回続きから再開できます）"
echo ""
python3 main.py --year "$INPUT_YEAR" --month "$INPUT_MONTH"
echo ""
echo "処理が完了しました。何かキーを押して終了..."
read -n 1
