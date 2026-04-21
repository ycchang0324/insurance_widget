#!/bin/bash

# --- 設定變數 ---
CHROME_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
DEBUG_PORT=9222
USER_DATA_DIR="$HOME/ChromeProfile"

# Python 檔案名稱
LOGIN_FILE="fubon_login.py"
ENROLL_FILE="enrollment.py"
SURRENDER_FILE="surrender.py"
QUERY_FILE="query_today.py"

echo "=================================================="
echo "🚀 正在初始化自動化環境..."

# 0. 先強行關閉可能殘留的 Chrome (避免 Port 被佔用)
killall "Google Chrome" > /dev/null 2>&1
sleep 1

# 1. 啟動 Google Chrome
echo "📂 使用設定檔: $USER_DATA_DIR"
"$CHROME_PATH" --remote-debugging-port=$DEBUG_PORT --user-data-dir="$USER_DATA_DIR" --no-first-run > /dev/null 2>&1 &

# 2. 迴圈檢查 Port 9222 是否開啟
echo "⏳ 等待 Chrome 偵測埠位 (Port: $DEBUG_PORT)..."
MAX_RETRIES=10
COUNT=0
while ! lsof -Pi :$DEBUG_PORT -sTCP:LISTEN -t >/dev/null; do
    sleep 1
    COUNT=$((COUNT+1))
    if [ $COUNT -ge $MAX_RETRIES ]; then
        echo "❌ 錯誤: 無法啟動 Chrome 除錯模式，請檢查路徑或權限。"
        exit 1
    fi
done

echo "✅ Chrome 遠端除錯模式已就緒！"
echo "⚠️  請手動清理不相關的分頁，只留下一個空白頁。"
read -p "👉 準備好了請按 [Enter] 開始登入..."

# 3. 執行登入並進入功能選單
if python3 "$LOGIN_FILE"; then
    while true; do
        echo ""
        echo "--------------------------------------------------"
        echo "🔑 登入狀態中，請選擇任務："
        echo " [e] 執行加保 (enrollment.py)"
        echo " [s] 執行退保 (surrender.py)"
        echo " [q] 查詢今日異動 (query_today.py)"
        echo " [x] 離開"
        echo "--------------------------------------------------"
        read -p "👉 輸入指令: " choice

        case $choice in
            [eE] ) python3 "$ENROLL_FILE" ;;
            [sS] ) python3 "$SURRENDER_FILE" ;;
            [qQ] ) python3 "$QUERY_FILE" ;;
            [xX] ) echo "👋 腳本結束"; break ;;
            * ) echo "❓ 無效指令" ;;
        esac
    done
else
    echo "❌ 登入過程發生錯誤。"
fi