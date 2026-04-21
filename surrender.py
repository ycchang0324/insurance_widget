import sys

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import logging

# 匯入工具函式
from src.utility import convert_to_roc_date, wait_for_spinner_to_disappear

# --- 1. 設定 Log 紀錄配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("./log/execution_surrender.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- 3. 填寫表格子程式 ---
def fill_fubon_surrender(driver, data):
    wait = WebDriverWait(driver, 10)
    
    # A. 選擇身分證 Radio (使用 JS 避免被遮罩)
    radio_xpath = "//input[@type='radio' and @value='2']"
    radio_btn = wait.until(EC.presence_of_element_located((By.XPATH, radio_xpath)))
    driver.execute_script("arguments[0].click();", radio_btn)
    
    # B. 填寫身分證
    id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@maxlength='10']")))
    target_id = str(data.get('身分證字號', '')).strip()
    
    # 清除舊內容並輸入
    id_field.clear() 
    driver.execute_script("arguments[0].value = arguments[1];", id_field, target_id)
    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", id_field)
    id_field.send_keys(Keys.ENTER)
    
    # 重要：等待身分證檢核完成，且下拉選單變為可點擊
    wait_for_spinner_to_disappear(driver)
    
    # C. 選擇變更項目 (改用更穩定的 XPATH)
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'ant-select-selection')]//div[@title='請選擇']")))
    driver.execute_script("arguments[0].click();", dropdown)
    
    # 等待選項出現並點擊
    option_xpath = "//li[contains(text(), '員工及眷屬退保')]"
    target_option = wait.until(EC.visibility_of_element_located((By.XPATH, option_xpath)))
    driver.execute_script("arguments[0].click();", target_option)

    # D. 下一步
    next_btn_xpath = "//button[contains(., '下一步') and contains(@class, 'primary')]"
    next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, next_btn_xpath)))
    driver.execute_script("arguments[0].click();", next_btn)
    wait_for_spinner_to_disappear(driver)

    # E. 填寫日期
    date_xpath = "//div[contains(., '退保/離職日期')]/following::input[@class='mx-input']"
    date_field = wait.until(EC.presence_of_element_located((By.XPATH, date_xpath)))
    target_date = convert_to_roc_date(data.get('退保日期', ''))
    
    # 填寫日期並觸發網頁事件
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, date_field, target_date)
    date_field.send_keys(Keys.ENTER) # 多補一個 Enter 嘗試收合日曆

    # F. 勾選方格 (JS 點擊最穩)
    checkbox_xpath = "//label[contains(@class, 'ant-checkbox-wrapper')]"
    checkbox = wait.until(EC.presence_of_element_located((By.XPATH, checkbox_xpath)))
    driver.execute_script("arguments[0].click();", checkbox)
    
    # G. 確保捲動到確定按鈕
    confirm_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '確定')]")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", confirm_btn)
    
    return confirm_btn

# --- 4. 主流程 ---
if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/empFamilyPolicyChangeQuery"
    
    try:
        df = pd.read_excel("./data/surrender.xlsx")
        try:
            protected_df = pd.read_excel("./data/protected.xlsx")
            protected_ids = set(protected_df['身分證字號'].astype(str).str.strip().tolist())
            logger.info(f"🛡️ 成功加載保護名單，共有 {len(protected_ids)} 位受保護人員。")
        except Exception as e:
            logger.warning(f"⚠️ 未找到保護名單，將不進行過濾。")
            protected_ids = set()
    except Exception as e:
        logger.error(f"❌ Excel 讀取失敗: {e}"); exit()

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)
    
    # 【關鍵補強 1】：確保 Selenium 盯著正確的 Page 分頁
    try:
        # 尋找現有的富邦分頁，如果找不到就用最後一個
        handles = driver.window_handles
        fubon_found = False
        for h in handles:
            driver.switch_to.window(h)
            if "fubonlife.com.tw" in driver.current_url:
                fubon_found = True
                break
        
        if not fubon_found:
            driver.switch_to.window(handles[-1]) # 切換到最後一個可用的標籤頁
            
        # 【關鍵補強 2】：使用 JavaScript 強制導航，繞過 driver.get 可能的掛起
        logger.info(f"正在前往: {target_url}")
        driver.execute_script(f"window.location.href = '{target_url}';")
        
    except Exception as e:
        logger.error(f"連線或導航失敗: {e}")
        sys.exit(1)

    input("\n🌐 網頁加載後，請在此按『Enter』開始執行退保作業...")

    # --- 初始化計數器 (移到迴圈外) ---
    success_count = 0
    failure_count = 0
    protected_count = 0 
    empty_count = 0 

    # --- 單層迴圈處理 ---
    for index, row in df.iterrows():
        emp_name = str(row.get('員工姓名', '')).strip()
        emp_id = str(row.get('身分證字號', '')).strip()

        # A. 檢查是否為空白列
        if not emp_name or emp_name.lower() == 'nan':
            empty_count += 1
            logger.info(f"💨 [跳過] 第 {index+1} 列為空白資料。")
            continue

        # B. 保護名單檢查
        if emp_id in protected_ids:
            protected_count += 1
            logger.info(f"⏭️  [跳過] {emp_name} ({emp_id}) 位列保護名單。")
            continue
        
        # C. 正常執行區塊
        try:
            logger.info(f"--- 串列 [{index+1}/{len(df)}] 開始處理: {emp_name} ---")
            
            # 【關鍵補強 3】：迴圈內回到初始頁面也要改用 JS 或檢查
            if target_url not in driver.current_url:
                driver.execute_script(f"window.location.href = '{target_url}';")
                wait_for_spinner_to_disappear(driver)
            
            confirm_btn = fill_fubon_surrender(driver, row)
            
            print(f"\n✅【自動填寫完成】人員：{emp_name}")
            input("👉 檢查資料後，請在此處按『Enter』正式送出...")
            
            driver.execute_script("arguments[0].click();", confirm_btn)
            logger.info(f"🎉 成功送出: {emp_name}")
            success_count += 1

        except Exception as e:
            failure_count += 1
            logger.error(f"❌ 流程出錯: {emp_name}，原因: {e}")
            input("👉 發生錯誤，請手動調整網頁至首頁後，按『Enter』繼續下一筆...")

    # 任務總結 (此處縮排需在迴圈外)
    logger.info(f"""
{'='*30}
任務執行結束統計：
✅ 成功送出: {success_count} 筆
🛡️ 保護跳過: {protected_count} 筆
💨 空白跳過: {empty_count} 筆
❌ 執行出錯: {failure_count} 筆
{'='*30}
    """)