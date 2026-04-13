import logging
import os
import sys
import socket
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 載入環境變數
load_dotenv()

# 匯入工具函式
try:
    from src.utility import wait_for_spinner_to_disappear
except ImportError:
    def wait_for_spinner_to_disappear(driver):
        pass

# --- 1. 設定 Log 紀錄配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("./log/execution_login.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- 2. 輔助函式：檢查 Port 是否有在監聽 ---
def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(('127.0.0.1', port)) == 0

# --- 3. 登入/狀態檢查子程式 ---
def check_and_login(driver, policy_no, account, password):
    """
    檢查是否已登入，若未登入則執行填表流程
    """
    wait = WebDriverWait(driver, 10)
    logout_btn_xpath = "//button[contains(., '登出')] | //span[contains(text(), '登出')]"
    
    # === A. 預檢機制：檢查是否已經在登入狀態 ===
    logger.info("檢查瀏覽器當前登入狀態...")
    try:
        # 使用極短的等待時間 (3秒) 檢查是否有登出按鈕
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, logout_btn_xpath)))
        logger.info("🎉 偵測到已處於登入狀態，跳過填表流程。")
        return True
    except:
        logger.info("未偵測到登入狀態，準備執行自動填表...")

    # === B. 執行自動填表流程 ===
    login_url = "https://gis.fubonlife.com.tw/gis-co-web/login"
    driver.get(login_url)
    
    try:
        # 1. 填寫保單號碼-序號
        policy_xpath = "//div[label[contains(text(), '保單號碼-序號')]]//input[@class='ant-input']"
        policy_field = wait.until(EC.presence_of_element_located((By.XPATH, policy_xpath)))
        policy_field.clear()
        policy_field.send_keys(policy_no)
        
        # 2. 填寫使用者帳號 (身分證字號)
        account_xpath = "//div[label[contains(text(), '身分證字號')]]//input"
        try:
            account_field = driver.find_element(By.XPATH, account_xpath)
        except:
            account_field = driver.find_element(By.XPATH, "//input[@placeholder='A168168888']")
        
        account_field.clear()
        account_field.send_keys(account)
        
        # 3. 填寫密碼
        password_field = driver.find_element(By.XPATH, "//input[@type='password']")
        password_field.clear()
        password_field.send_keys(password)
        
        logger.info("✅ 帳號資訊已自動填寫完成。")
        print("\n" + "="*50)
        print(" ✍️   請在瀏覽器輸入驗證碼並點擊「登入」")
        print(" ⏳   程式將自動偵測登入狀態 (限時 60 秒)...")
        print("="*50)

        # 4. 等待登入成功 (偵測「登出」按鈕出現)
        success_wait = WebDriverWait(driver, 60)
        success_wait.until(EC.presence_of_element_located((By.XPATH, logout_btn_xpath)))
        wait_for_spinner_to_disappear(driver)
        
        logger.info("🎉 偵測到登出按鈕，判定登入成功！")
        return True

    except Exception as e:
        logger.error(f"❌ 填表或登入偵測過程中發生錯誤: {e}")
        return False

# --- 4. 主流程 ---
if __name__ == "__main__":
    POLICY_NO = os.getenv("FUBON_POLICY_NO", "")
    USER_ACCOUNT = os.getenv("FUBON_USER_ACCOUNT", "")
    USER_PASSWORD = os.getenv("FUBON_USER_PASSWORD", "")
    DEBUG_PORT = 9222

    if not all([POLICY_NO, USER_ACCOUNT, USER_PASSWORD]):
        logger.error("❌ 錯誤：.env 檔案資訊不完整。")
        sys.exit(1)

    # 🛑 核心檢查：如果 Port 沒開，絕對不允許啟動，避免 Selenium 自己開新視窗
    if not is_port_in_use(DEBUG_PORT):
        print("\n" + "!"*60)
        print(f"❌ 嚴重錯誤：偵測不到在 Port {DEBUG_PORT} 執行的除錯 Chrome！")
        print("這就是為什麼它會一直開啟『新視窗』的原因。")
        print("\n💡 解決辦法：")
        print(f"1. 請先『完全退出』(Command + Q) 所有的 Google Chrome。")
        print(f"2. 再次執行指令啟動：")
        print(f"   /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222 --user-data-dir=\"$HOME/ChromeProfile\" --remote-allow-origins=*")
        print("!"*60 + "\n")
        sys.exit(1)

    chrome_options = Options()
    # 關鍵：強制指定連線埠位，並允許所有來源連線
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
    chrome_options.add_argument("--remote-allow-origins=*")
    
    try:
        # 當 debuggerAddress 有效且 Port 已開啟時，Selenium 只會「連接」
        driver = webdriver.Chrome(options=chrome_options)
        
        # 再次檢查確保沒有產生新的 Session (如果 window_handles 多於原本的，表示可能連接失敗)
        if check_and_login(driver, POLICY_NO, USER_ACCOUNT, USER_PASSWORD):
            sys.exit(0)
        else:
            sys.exit(1)
    except Exception as e:
        logger.error(f"連線至現有 Chrome 失敗: {e}")
        sys.exit(1)