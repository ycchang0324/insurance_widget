import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import platform

# 自動判定作業系統使用 Command 或 Control
CMD_KEY = Keys.COMMAND if platform.system() == "Darwin" else Keys.CONTROL

# 匯入工具函式
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

        logger.info(f"身分證字號填寫完成，等待遮罩結束...")
        wait_for_spinner_to_disappear(driver)

        # D. 生日 (JS 點擊)
        b_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '生日')]/ancestor::div[contains(@class, 'ant-form-item')]//input[@name='date']")))
        driver.execute_script("arguments[0].click();", b_field)
        b_field.send_keys(CMD_KEY + "a", Keys.BACKSPACE)
        b_field.send_keys(str(data['生日']), Keys.ENTER)

        # --- E. 國籍填寫 ---
        nationality_display_xpath = "//label[contains(., '國籍')]/ancestor::div[contains(@class, 'ant-form-item')]//div[@class='ant-select-selection-selected-value']"
        target_nation = str(data['國籍']).strip()
        skip_list = ["中華民國", "台灣", "臺灣", "ROC", "R.O.C", "nan", ""]

        current_web_nation = wait.until(EC.presence_of_element_located((By.XPATH, nationality_display_xpath))).text.strip()
        
        if not (target_nation in skip_list and current_web_nation == "中華民國"):
            logger.info(f"正在修改國籍為: {target_nation}")
            trigger = wait.until(EC.presence_of_element_located((By.XPATH, nationality_display_xpath)))
            driver.execute_script("arguments[0].click();", trigger)

            inp = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@class='ant-select-search__field']")))
            inp.send_keys(CMD_KEY + "a", Keys.BACKSPACE)
            inp.send_keys(target_nation)
            inp.send_keys(Keys.ENTER)
            wait_for_spinner_to_disappear(driver)

        # F. 性別 (嚴格校驗 + JS 點擊)
        raw_gender = str(data['性別']).strip().upper()
        if raw_gender in ["男", "男性", "M"] or (len(raw_gender) > 0 and raw_gender.startswith('M')):
            gender_val = "M"
        elif raw_gender in ["女", "女性", "F"] or (len(raw_gender) > 0 and raw_gender.startswith('F')):
            gender_val = "F"
        else:
            raise ValueError(f"性別欄位格式錯誤或空白：'{raw_gender}'")

        gender_radio_xpath = f"//input[@value='{gender_val}']/ancestor::label"
        gender_element = wait.until(EC.presence_of_element_located((By.XPATH, gender_radio_xpath)))
        driver.execute_script("arguments[0].click();", gender_element)

        # G. 受僱日期 (JS 點擊)
        h_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '受僱日期')]/ancestor::div[contains(@class, 'ant-form-item')]//input[@name='date']")))
        driver.execute_script("arguments[0].click();", h_field)
        h_field.send_keys(CMD_KEY + "a", Keys.BACKSPACE)
        h_field.send_keys(str(data['受僱日期']), Keys.ENTER)

        # H. 工作內容
        job_field = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(., '工作內容')]/ancestor::div[contains(@class, 'ant-form-item')]//input")))
        job_field.clear()
        job_field.send_keys(str(data['工作內容']))

        wait_for_spinner_to_disappear(driver)

        # I. 下一步 (第一頁 - 改用 JS 點擊)
        next_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., '下一步')]")))
        driver.execute_script("arguments[0].click();", next_btn)
        wait_for_spinner_to_disappear(driver) 

        # J. 第二頁保險金額 (防呆 nan)
        products = ["GADD", "GMR"]
        for prod in products:
            val = str(data.get(prod, '')).strip()
            if val and val.lower() != 'nan':
                xpath = f"//label[contains(., '{prod}')]/following::input[1]"
                field = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                field.clear()
                field.send_keys(val, Keys.TAB)

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
        logger.info(f"✅ 送出成功: {data['員工姓名']}")

    except Exception as e:
        logger.error(f"❌ 填寫出錯: {data.get('員工姓名', '未知')}，原因: {e}")
        raise e

# --- 3. 主流程 ---
if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/employeeEnrollment/employeeEnrollmentApplication"
    
    try:
        df = pd.read_excel("./data/enrollment.xlsx")
        # 不在這裡 dropna，讓空白列進迴圈計算
        
        try:
            protected_df = pd.read_excel("./data/protected.xlsx")
            protected_ids = set(protected_df['身分證字號'].astype(str).str.strip().tolist())
            logger.info(f"🛡️ 保護名單讀取成功，共 {len(protected_ids)} 筆。")
        except:
            protected_ids = set()
            logger.warning("⚠️ 未找到保護名單。")

        df['生日'] = df['生日'].apply(convert_to_roc_date)
        df['受僱日期'] = df['受僱日期'].apply(convert_to_roc_date)
        logger.info(f"Excel 讀取完成，共 {len(df)} 筆。")
        
    except Exception as e:
        logger.error(f"Excel 讀取失敗: {e}"); exit()

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    # 初始化計數器
    success_count = 0
    failure_count = 0
    protected_count = 0
    empty_count = 0

    driver.get(target_url)
    input("👉 網頁就緒後，請在此按『Enter』開始執行...")

    for index, row in df.iterrows():
        emp_name = str(row.get('員工姓名', '')).strip()
        emp_id = str(row.get('身分證字號', '')).strip()

        # 1. 空白檢查
        if not emp_name or emp_name.lower() == 'nan':
            empty_count += 1
            continue

        # 2. 保護名單檢查
        if emp_id in protected_ids:
            protected_count += 1
            logger.info(f"⏭️  [跳過] {emp_name} 位列保護名單。")
            continue

        logger.info(f"--- 串列 [{index+1}/{len(df)}] 處理中: {emp_name} ---")
        
        try:
            driver.get(target_url)
            name_xpath = "//label[contains(., '員工姓名')]/ancestor::div[contains(@class, 'ant-row')]//textarea"
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, name_xpath)))
            
            fill_fubon_enrollment(driver, row)
            success_count += 1
            logger.info(f"🎉 處理完成: {emp_name}")

        except Exception as e:
            failure_count += 1
            logger.error(f"❌ 填寫失敗: {emp_name}")
            print(f"\n⚠️ 【錯誤發生】人員：{emp_name}\n原因: {str(e)[:150]}")
            input("👉 請手動處理後按『Enter』跳往下一筆...")
        
        if index < len(df) - 1:
            logger.info("⏳ 準備切換下一位人員...")

    # 任務總結
    logger.info(f"""
{'='*30}
加保任務結束統計：
✅ 成功: {success_count} 筆
🛡️ 保護跳過: {protected_count} 筆
💨 空白跳過: {empty_count} 筆
❌ 失敗: {failure_count} 筆
{'='*30}
    """)