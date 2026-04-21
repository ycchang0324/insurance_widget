#!/bin/bash

# --- 設定變數 ---
CHROME_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
DEBUG_PORT=9222
USER_DATA_DIR="$HOME/ChromeProfile"
LOGIN_FILE="fubon_login.py"
ENROLL_FILE="enrollment.py"
SURRENDER_FILE="surrender.py"
QUERY_FILE="query_today.py"

echo "=================================================="
echo "🚀 準備啟動帶有除錯模式的 Chrome..."
echo "📂 使用者設定檔路徑: $USER_DATA_DIR"
echo "=================================================="

# 1. 啟動 Google Chrome
"$CHROME_PATH" --remote-debugging-port=$DEBUG_PORT --user-data-dir="$USER_DATA_DIR" > /dev/null 2>&1 &

echo "⏳ 等待 Chrome 啟動中 (3秒)..."
sleep 3

# 2. 檢查 Chrome 是否成功開啟埠位
if lsof -Pi :$DEBUG_PORT -sTCP:LISTEN -t >/dev/null ; then
    echo "✅ Chrome 遠端除錯模式已就緒 (Port: $DEBUG_PORT)"
    
    echo ""
    echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    echo "⚠️  重要提醒：請先處理 Chrome 分頁"
    echo " 1. 請將該除錯視窗內『所有不相關的分頁』關閉。"
    echo " 2. 確保只留下一個『新分頁』即可。"
    echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    read -p "👉 確認分頁已清理完成，請按『Enter』開始導向登入頁面..." 

    echo "🐍 準備執行登入程式: $LOGIN_FILE..."
    
    # 執行登入 Python 程式
    if python3 "$LOGIN_FILE"; then
        
        # --- 新增：無限迴圈選單 ---
        while true; do
            echo ""
            echo "--------------------------------------------------"
            echo "🔑 目前處於登入狀態，請選擇要執行的任務："
            echo " [e] 執行加保 (enrollment.py)"
            echo " [s] 執行退保 (surrender.py)"
            echo " [q] 查詢當日加退保情況 (query_today.py)"
            echo " [x] 結束並關閉腳本"
            echo "--------------------------------------------------"
            read -p "👉 請輸入指令並按 Enter: " user_choice

            case $user_choice in
                [eE] )
                    echo "🚀 正在啟動 加保自動化 ($ENROLL_FILE)..."
                    python3 "$ENROLL_FILE"
                    ;;
                [sS] )
                    echo "🚀 正在啟動 退保自動化 ($SURRENDER_FILE)..."
                    python3 "$SURRENDER_FILE"
                    ;;
                [qQ] )
                    echo "🚀 正在啟動 今日異動查詢與下載 ($QUERY_FILE)..."
                    python3 "$QUERY_FILE"
                    ;;
                [xX] )
                    echo "👋 已選擇結束，腳本關閉中。"
                    break # 跳出 while 迴圈
                    ;;
                * )
                    echo "❓ 無效指令，請重新輸入。"
                    ;;
            esac
            
            echo ""
            echo "✅ 任務執行完畢，返回主選單..."
            sleep 1
        done
        # ----------------------------
        
    else
        echo ""
        echo "⚠️  登入失敗或偵測超時，將不執行後續自動化選單。"
    fi

else
    echo "❌ 錯誤: 無法偵測到 Chrome 的除錯埠位，請確認 Chrome 是否正確開啟。"
    exit 1
fi

echo "=================================================="
echo "✨ 腳本執行流程結束"
echo "=================================================="