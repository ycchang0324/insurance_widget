#!/bin/bash

# --- 設定變數 ---
CHROME_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
DEBUG_PORT=9222
USER_DATA_DIR="$HOME/ChromeProfile"
LOGIN_FILE="fubon_login.py"
ENROLL_FILE="enrollment.py"
SURRENDER_FILE="surrender.py"

echo "=================================================="
echo "🚀 準備啟動帶有除錯模式的 Chrome..."
echo "📂 使用者設定檔路徑: $USER_DATA_DIR"
echo "=================================================="

# 1. 啟動 Google Chrome
"$CHROME_PATH" --remote-debugging-port=$DEBUG_PORT --user-data-dir="$USER_DATA_DIR" > /dev/null 2>&1 &

echo "⏳ 等待 Chrome 啟動中 (1秒)..."
sleep 1

# 2. 檢查 Chrome 是否成功開啟埠位
if lsof -Pi :$DEBUG_PORT -sTCP:LISTEN -t >/dev/null ; then
    echo "✅ Chrome 遠端除錯模式已就緒 (Port: $DEBUG_PORT)"
    
    echo "🐍 準備執行登入程式: $LOGIN_FILE..."
    
    # 執行登入 Python 程式，並判斷其結束狀態
    if python3 "$LOGIN_FILE"; then
        # 如果 python 傳回 exit 0 (成功)，則繼續執行
        echo ""
        echo "--------------------------------------------------"
        echo "🔑 登入成功，請選擇後續要執行的自動化任務："
        echo " [e] 執行增員 (enrollment.py)"
        echo " [s] 執行退保 (surrender.py)"
        echo " [any] 按其他任意鍵結束腳本"
        echo "--------------------------------------------------"
        read -p "👉 請輸入指令並按 Enter: " user_choice

        case $user_choice in
            [eE] )
                echo "🚀 正在啟動 增員自動化 ($ENROLL_FILE)..."
                python3 "$ENROLL_FILE"
                ;;
            [sS] )
                echo "🚀 正在啟動 退保自動化 ($SURRENDER_FILE)..."
                python3 "$SURRENDER_FILE"
                ;;
            * )
                echo "👋 已選擇結束，腳本關閉中。"
                ;;
        esac
    else
        # 如果 python 傳回 exit 1 (失敗)，則終止腳本
        echo ""
        echo "⚠️  登入失敗或偵測超時，將不執行後續自動化選單。"
        echo "請檢查網頁狀態或重新執行腳本。"
    fi

else
    echo "❌ 錯誤: 無法偵測到 Chrome 的除錯埠位，請確認 Chrome 是否正確開啟。"
    exit 1
fi

echo "=================================================="
echo "✨ 腳本執行流程結束"
echo "=================================================="