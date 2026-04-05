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
    
    # A. 選擇身分證 Radio
    radio_xpath = "//input[@type='radio' and @value='2']"
    radio_btn = wait.until(EC.presence_of_element_located((By.XPATH, radio_xpath)))
    driver.execute_script("arguments[0].click();", radio_btn)
    
    # B. 填寫身分證
    id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@maxlength='10']")))
    target_id = str(data.get('身分證字號', '')).strip()
    driver.execute_script("arguments[0].value = arguments[1];", id_field, target_id)
    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", id_field)
    id_field.send_keys(Keys.ENTER)
    wait_for_spinner_to_disappear(driver)

    # C. 選擇變更項目
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='請選擇']")))
    driver.execute_script("arguments[0].click();", dropdown)
    
    option_xpath = "//li[contains(text(), '員工及眷屬退保')]"
    target_option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
    driver.execute_script("arguments[0].click();", target_option)

    # D. 下一步
    next_btn_xpath = "//button[contains(., '下一步') and contains(@class, 'primary')]"
    next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, next_btn_xpath)))
    driver.execute_script("arguments[0].click();", next_btn)
    wait_for_spinner_to_disappear(driver)

    # E. 填寫日期 (使用 utility 轉換)
    date_xpath = "//div[contains(., '退保/離職日期')]/following::input[@class='mx-input']"
    date_field = wait.until(EC.element_to_be_clickable((By.XPATH, date_xpath)))
    target_date = convert_to_roc_date(data.get('退保日期', ''))
    
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
    """, date_field, target_date)
    
    # F. 勾選方格與收合日曆
    checkbox_xpath = "//label[contains(@class, 'user__check') and contains(@class, 'ant-checkbox-wrapper')]"
    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
    driver.execute_script("arguments[0].click();", checkbox)
    
    info_txt = driver.find_element(By.XPATH, "//p[contains(@class, 'info__txt')]")
    actions = ActionChains(driver)
    actions.move_to_element(info_txt).click().perform()
    
    # G. 捲動到確定按鈕
    confirm_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '確定')]")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", confirm_btn)
    return confirm_btn

# --- 4. 主流程 ---
if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/empFamilyPolicyChangeQuery"
    
    try:
        # 讀取 Excel
        df = pd.read_excel("./data/surrender.xlsx")
        # 讀取保護名單
        try:
            protected_df = pd.read_excel("./data/protected.xlsx")
            protected_ids = set(protected_df['身分證字號'].astype(str).str.strip().tolist())
            logger.info(f"成功加載保護名單，共有 {len(protected_ids)} 位受保護人員。")
        except Exception as e:
            logger.warning(f"未找到保護名單，將不進行過濾。")
            protected_ids = set()
    except Exception as e:
        logger.error(f"Excel 讀取失敗: {e}"); exit()

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)
    
    # 前置確認
    driver.get(target_url)
    input("\n🌐 網頁加載後，請在此按『Enter』開始...")

    success_count, failure_count = 0, 0

    for index, row in df.iterrows():
        emp_name = str(row.get('員工姓名', '未知'))
        emp_id = str(row.get('身分證字號', '')).strip()
        
        # 保護名單檢查
        if emp_id in protected_ids:
            failure_count += 1
            logger.warning(f"🚫 攔截成功！{emp_name} 位列保護名單，已跳過。")
            continue
        
        try:
            if driver.current_url != target_url:
                driver.get(target_url)
            
            confirm_btn = fill_fubon_surrender(driver, row)
            
            print(f"\n✅【自動填寫完成】人員：{emp_name}")
            input("👉 檢查資料後，請在此處按『Enter』正式送出...")
            
            driver.execute_script("arguments[0].click();", confirm_btn)
            logger.info(f"🎉 成功送出: {emp_name}")
            success_count += 1

        except Exception as e:
            failure_count += 1
            logger.error(f"❌ 流程出錯: {emp_name}，原因: {e}")
            input("👉 發生錯誤，請手動調整後按『Enter』繼續...")

    logger.info(f"\n任務結束: 成功 {success_count}, 失敗 {failure_count}")