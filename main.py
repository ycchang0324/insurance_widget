import os
import subprocess
import time
import platform
import socket
import sys
from datetime import datetime

# --- 基礎路徑設定 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOG_DIR = os.path.join(BASE_DIR, "log")
DOWNLOAD_DIR = os.path.join(DATA_DIR, "downloads")
PROTECTED_FILE = os.path.join(DATA_DIR, "protected.xlsx")

DEBUG_PORT = 9222

def init_folders():
    """初始化必要資料夾"""
    for folder in [LOG_DIR, DOWNLOAD_DIR]:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"📁 已建立資料夾: {folder}")

def check_protected_file():
    """檢查保護名單更新狀態"""
    print("\n🔍 檢查資源狀態...")
    if not os.path.exists(PROTECTED_FILE):
        print("⚠️  警告: 找不到 data/protected.xlsx，請確認檔案路徑！")
        return

    mtime = os.path.getmtime(PROTECTED_FILE)
    last_mod = datetime.fromtimestamp(mtime)
    diff_days = (datetime.now() - last_mod).days

    print(f"📄 保護名單最後修改: {last_mod.strftime('%Y-%m-%d %H:%M')}")
    
    if diff_days > 7:
        print("=" * 60)
        print(f"🚨 【提醒】保護名單已超過 {diff_days} 天未更新！")
        print("👉 請確認是否需要更新最新名單後再執行比對任務。")
        print("=" * 60)
    else:
        print("✅ 保護名單狀態：近期已更新。")

def get_config():
    sys_platform = platform.system()
    if sys_platform == "Darwin":  # macOS
        return {
            "chrome_path": "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "user_data_dir": "/tmp/chrome_automation",
            "kill_cmd": ["killall", "Google Chrome"]
        }
    elif sys_platform == "Windows":
        return {
            "chrome_path": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "user_data_dir": os.path.join(os.environ.get("LOCALAPPDATA", "C:"), "Temp", "chrome_automation"),
            "kill_cmd": ["taskkill", "/F", "/IM", "chrome.exe"]
        }
    else:
        print(f"❌ 不支援的作業系統: {sys_platform}")
        sys.exit(1)

def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

def run_task(file_name):
    print(f"▶️ 正在執行 {file_name}...")
    # 確保使用當前的 Python 環境執行子程式
    result = subprocess.run([sys.executable, os.path.join(BASE_DIR, file_name)])
    return result.returncode == 0

def main():
    init_folders()
    check_protected_file()
    
    config = get_config()
    print("\n" + "=" * 60)
    print("🚀 正在初始化跨平台自動化環境...")

    # 關閉殘留 Chrome
    try:
        subprocess.run(config["kill_cmd"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1)
    except:
        pass

    # 啟動 Chrome
    subprocess.Popen([
        config["chrome_path"],
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={config['user_data_dir']}",
        "--no-first-run"
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # 檢查 Port
    ready = False
    for _ in range(10):
        if is_port_open(DEBUG_PORT):
            ready = True
            break
        time.sleep(1)
    
    if not ready:
        print("❌ 錯誤: 無法啟動 Chrome 除錯模式，請檢查 Chrome 路徑或權限。")
        return

    print("✅ Chrome 遠端除錯模式已就緒！")
    input("👉 請確保瀏覽器已開啟且只有一個分頁，準備好了請按 [Enter] 開始登入...")

    if run_task("fubon_login.py"):
        while True:
            # 獲取最新修改日期顯示在選單上方
            mod_date = datetime.fromtimestamp(os.path.getmtime(PROTECTED_FILE)).strftime('%m/%d')
            print("\n" + "-" * 60)
            print(f"🔑 登入成功 | 保護名單日期: {mod_date}")
            print("請選擇任務：")
            print(" [e] 執行加保 (enrollment.py)")
            print(" [s] 執行退保 (surrender.py)")
            print(" [q] 查詢今日異動 (query_today.py)")
            print(" [x] 離開")
            print("-" * 60)
            
            choice = input("👉 指令: ").lower()
            if choice == 'e': run_task("enrollment.py")
            elif choice == 's': run_task("surrender.py")
            elif choice == 'q': run_task("query_today.py")
            elif choice == 'x': 
                print("👋 腳本結束"); break
            else: print("❓ 無效指令")
    else:
        print("❌ 登入過程發生錯誤。")

if __name__ == "__main__":
    main()