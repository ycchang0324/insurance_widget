import logging
import os
import sys
import socket
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 載入環境變數
load_dotenv()

# --- 1. 設定 Log ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("./log/execution_login.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(('127.0.0.1', port)) == 0

def fill_and_wait_login(driver, policy_no, account, password):
    wait = WebDriverWait(driver, 15)
    login_url = "https://gis.fubonlife.com.tw/gis-co-web/login"
    
    logger.info(f"🚀 嘗試跳轉至: {login_url}")
    driver.get(login_url)
    time.sleep(3)

    try:
        # --- 核心修改：使用 JS 填寫以繞過遮罩攔截 ---
        def js_fill(element, value):
            # 直接賦值並觸發 input/change 事件，確保網頁框架(React/Vue)有收到值
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].blur();
            """, element, value)

        # A. 填寫保單號碼
        policy_xpath = "//div[label[contains(text(), '保單號碼-序號')]]//input[@class='ant-input']"
        policy_field = wait.until(EC.presence_of_element_located((By.XPATH, policy_xpath)))
        js_fill(policy_field, policy_no)
        logger.info("✅ 保單號碼填寫完成 (JS)")

        # B. 身分證字號
        account_xpath = "//div[label[contains(text(), '身分證字號')]]//input"
        account_field = wait.until(EC.presence_of_element_located((By.XPATH, account_xpath)))
        js_fill(account_field, account)
        logger.info("✅ 身分證字號填寫完成 (JS)")

        # C. 密碼
        password_xpath = "//input[@type='password']"
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, password_xpath)))
        js_fill(password_field, password)
        logger.info("✅ 密碼填寫完成 (JS)")
        
        # --- 提示使用者 ---
        print("\n" + "="*50)
        print(" ✍️   遮罩已繞過，請手動輸入驗證碼並點擊「登入」")
        print("="*50)

        # D. 偵測登入成功
        success_wait = WebDriverWait(driver, 60)
        logout_btn_xpath = "//button[contains(., '登出')] | //span[contains(text(), '登出')]"
        success_wait.until(EC.presence_of_element_located((By.XPATH, logout_btn_xpath)))
        
        logger.info("🎉 登入成功！")
        return True

    except Exception as e:
        logger.error(f"❌ 填表偵測失敗: {e}")
        return False

if __name__ == "__main__":
    POLICY_NO = os.getenv("FUBON_POLICY_NO", "")
    USER_ACCOUNT = os.getenv("FUBON_USER_ACCOUNT", "")
    USER_PASSWORD = os.getenv("FUBON_USER_PASSWORD", "")
    DEBUG_PORT = 9222

    if not is_port_in_use(DEBUG_PORT):
        logger.error("❌ Chrome 除錯埠位未開啟，請先執行腳本啟動 Chrome")
        sys.exit(1)

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(15) # 避免頁面卡死

        # --- 核心邏輯：尋找正確的網頁分頁 ---
        logger.info("🔍 正在尋找可控制的網頁分頁...")
        target_handle = None
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            curr_url = driver.current_url
            # 過濾掉 Chrome 內部的特殊分頁
            if not any(x in curr_url for x in ["chrome-extension://", "devtools://", "about:blank"]):
                target_handle = handle
                break
        
        # 如果沒找到任何分頁，或者都在空白頁，就開一個新的
        if not target_handle:
            logger.info("🆕 未發現活動網頁，正在開啟新分頁...")
            driver.execute_script("window.open('about:blank', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
        else:
            driver.switch_to.window(target_handle)
            
        logger.info(f"🎯 已鎖定分頁，目前網址: {driver.current_url}")

        if fill_and_wait_login(driver, POLICY_NO, USER_ACCOUNT, USER_PASSWORD):
            sys.exit(0)
        else:
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"❌ 程式執行異常: {e}")
        sys.exit(1)