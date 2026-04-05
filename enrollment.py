import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 匯入剛移過去的工具函式
from src.utility import convert_to_roc_date, wait_for_spinner_to_disappear

# --- 1. 設定 Log 紀錄配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("./log/execution_enrollment.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# --- 2. 填寫表格主函式 ---
def fill_fubon_enrollment(driver, data):
    wait = WebDriverWait(driver, 10)
    try:
        logger.info(f">>> 正在填寫: {data['員工姓名']}")
        
        # 確保進入頁面時遮罩已消失
        wait_for_spinner_to_disappear(driver)

        # A. 員工姓名
        name_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '員工姓名')]/ancestor::div[contains(@class, 'ant-row')]//textarea")))
        name_field.clear()
        name_field.send_keys(str(data['員工姓名']))

        # B. 護照英文名字
        en_name = str(data.get('護照英文名字', '')).strip()
        if en_name and en_name.lower() != 'nan':
            en_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '英文姓名')]/ancestor::div[contains(@class, 'ant-form-item')]//textarea")))
            en_field.clear()
            en_field.send_keys(en_name)
        
        # C. 身分證字號
        id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '身分證')]/ancestor::div[contains(@class, 'ant-form-item')]//input")))
        id_field.clear()
        id_field.send_keys(str(data['身分證字號']))

        # 確保進入頁面時遮罩已消失
        logger.info(f"身分證字號填寫完成，等待遮罩結束...")
        # 4. 💡 依照您的要求，執行完後等待遮罩結束
        wait_for_spinner_to_disappear(driver)

        # D. 生日 (使用 JS 點擊避開遮罩)
        b_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '生日')]/ancestor::div[contains(@class, 'ant-form-item')]//input[@name='date']")))
        driver.execute_script("arguments[0].click();", b_field)
        b_field.send_keys(Keys.COMMAND + "a", Keys.BACKSPACE)
        b_field.send_keys(str(data['生日']), Keys.ENTER)

        # --- E. 國籍填寫 (單次 Enter 穩定版) ---
        nationality_display_xpath = "//label[contains(., '國籍')]/ancestor::div[contains(@class, 'ant-form-item')]//div[@class='ant-select-selection-selected-value']"
        
        target_nation = str(data['國籍']).strip()
        skip_list = ["中華民國", "台灣", "臺灣", "ROC", "R.O.C", "nan", ""]

        # 取得網頁目前的值
        current_web_nation = wait.until(EC.presence_of_element_located((By.XPATH, nationality_display_xpath))).text.strip()
        
        if not (target_nation in skip_list and current_web_nation == "中華民國"):
            logger.info(f"正在修改國籍為: {target_nation}")
            
            # 1. 點開下拉選單 (使用 JS 點擊避開攔截)
            trigger = wait.until(EC.presence_of_element_located((By.XPATH, nationality_display_xpath)))
            driver.execute_script("arguments[0].click();", trigger)

            # 2. 定位輸入框並輸入文字
            inp = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@class='ant-select-search__field']")))
            
            # 清除舊內容
            inp.send_keys(Keys.COMMAND + "a") # Windows 請改 Keys.CONTROL
            inp.send_keys(Keys.BACKSPACE)
            
            # 填入國籍
            inp.send_keys(target_nation)
            logger.info(f"已輸入國籍文字: {target_nation}，等待過濾...")
            inp.send_keys(Keys.ENTER)
            
            # 4. 💡 依照您的要求，執行完後等待遮罩結束
            logger.info(f"國籍填寫完成，等待遮罩結束...")
            # 確保進入頁面時遮罩已消失
            wait_for_spinner_to_disappear(driver)

        # F. 性別
        gender_val = "M" if "男" in str(data['性別']) else "F"
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//input[@value='{gender_val}']/ancestor::label"))).click()

        # G. 受僱日期 (使用 JS 點擊避開遮罩)
        h_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '受僱日期')]/ancestor::div[contains(@class, 'ant-form-item')]//input[@name='date']")))
        driver.execute_script("arguments[0].click();", h_field)
        h_field.send_keys(Keys.COMMAND + "a", Keys.BACKSPACE)
        h_field.send_keys(str(data['受僱日期']), Keys.ENTER)

        # H. 工作內容
        job_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '工作內容')]/ancestor::div[contains(@class, 'ant-form-item')]//input")))
        job_field.clear()
        job_field.send_keys(str(data['工作內容']))

        # I. 下一步 (第一頁)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '下一步')]"))).click()
        wait_for_spinner_to_disappear(driver) # 跳轉後的等待

        # J. 第二頁保險金額
        products = ["GADD", "GMR"]
        for prod in products:
            if prod in data:
                xpath = f"//label[contains(., '{prod}')]/following::input[1]"
                field = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                field.clear()
                field.send_keys(str(data[prod]), Keys.TAB)

        # K. 第二頁下一步 (JS 點擊)
        next_p2 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '下一步')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_p2)
        driver.execute_script("arguments[0].click();", next_p2)
        wait_for_spinner_to_disappear(driver)

        # L. 最終確認
        confirm_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '確定')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", confirm_btn)
        
        print(f"\n--- 📋 待確認人員：{data['員工姓名']} ---")
        input("👉 檢查資料後，請在此處按『Enter』正式送出...")
        
        driver.execute_script("arguments[0].click();", confirm_btn)
        logger.info(f"✅ 手動送出成功: {data['員工姓名']}")

    except Exception as e:
        logger.error(f"❌ 填寫出錯: {data.get('員工姓名', '未知')}，原因: {e}")
        raise e

# --- 3. 主流程 ---
if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/employeeEnrollment/employeeEnrollmentApplication"
    
    try:
        # 讀取 Excel 並預處理日期
        df = pd.read_excel("./data/enrollment.xlsx")
        df['生日'] = df['生日'].apply(convert_to_roc_date)
        df['受僱日期'] = df['受僱日期'].apply(convert_to_roc_date)
        logger.info(f"Excel 讀取完成，共 {len(df)} 筆。")
    except Exception as e:
        logger.error(f"Excel 讀取失敗: {e}")
        exit()

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    success_count, failure_count = 0, 0

    # 啟動前置檢查
    logger.info("正在加載初始頁面...")
    driver.get(target_url)
    input("👉 網頁就緒後，請在此按『Enter』開始執行自動化...")

    for index, row in df.iterrows():
        emp_name = str(row.get('員工姓名', '未知'))
        logger.info(f"--- 串列 [{index+1}/{len(df)}] 開始處理: {emp_name} ---")
        
        try:
            # 每一筆開始前回到初始頁面
            driver.get(target_url)
            
            # 等待關鍵元素出現，確認頁面加載成功
            name_xpath = "//label[contains(., '員工姓名')]/ancestor::div[contains(@class, 'ant-row')]//textarea"
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, name_xpath)))
            
            # 執行填寫 (內部已有成功後的 Enter 確認點)
            fill_fubon_enrollment(driver, row)
            success_count += 1
            logger.info(f"🎉 處理完成: {emp_name}")

        except Exception as e:
            # 💡 關鍵改動：失敗時的處理邏輯
            failure_count += 1
            logger.error(f"❌ 填寫失敗: {emp_name}")
            print(f"\n⚠️ 【錯誤發生】人員：{emp_name}")
            print(f"原因: {str(e)[:200]}") # 顯示前 200 字的錯誤詳情
            
            # 停下來等你檢查網頁
            print("-" * 30)
            input("👉 請檢查網頁報錯原因，手動處理後在此按『Enter』跳往下一筆...")
            print("-" * 30)
        
        if index < len(df) - 1:
            logger.info("⏳ 準備切換下一位人員...")

    # 任務總結
    logger.info(f"\n{'='*30}\n加保任務結束\n✅ 成功: {success_count} 筆\n❌ 失敗: {failure_count} 筆\n{'='*30}")