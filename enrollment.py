import time

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

        # J. 第二頁保險金額 (處理 100.0 -> 100 並繞過遮罩)
        products = ["GADD", "GMR"]
        for prod in products:
            raw_val = data.get(prod, '')
            
            # 1. 檢查是否為 NaN 或空值
            if str(raw_val).lower() == 'nan' or raw_val is None or str(raw_val).strip() == '':
                continue
                
            # 2. 格式化數值：將 100.0 轉為 100
            try:
                # 透過 float -> int -> str 的轉換路徑去掉小數點
                clean_val = str(int(float(raw_val)))
            except (ValueError, TypeError):
                # 如果是純文字則保持原樣
                clean_val = str(raw_val).strip()

            # 3. 執行填寫
            try:
                xpath = f"//label[contains(., '{prod}')]/following::input[1]"
                # 使用 presence 確保抓到元素即可，不必管它有沒有被遮罩擋住
                field = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                
                # 使用 JavaScript 直接塞值並觸發網頁事件
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, field, clean_val)
                
                # 補一個 TAB 鍵確保觸發頁面原本的計算邏輯（如果有）
                field.send_keys(Keys.TAB)
                
                logger.info(f"✅ {prod} 填寫成功: {clean_val}")
                
            except Exception as e:
                logger.error(f"❌ 填寫 {prod} 時發生錯誤: {e}")

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
    
    # 1. 讀取資料 (保持不變)
    try:
        df = pd.read_excel("./data/enrollment.xlsx")
        # 強化保護名單讀取：增加標準化處理
        try:
            protected_df = pd.read_excel("./data/protected.xlsx").dropna(subset=['身分證字號'])
            protected_ids = set(
                protected_df['身分證字號']
                .astype(str)
                .str.strip()
                .str.upper() # 確保大小寫一致
                .tolist()
            )
            logger.info(f"🛡️ 保護名單讀取成功，共 {len(protected_ids)} 筆。")
        except:
            protected_ids = set()
            logger.warning("⚠️ 未找到保護名單。")

        # --- 新增：啟動前的視覺化提醒 ---
        # 檢查 Excel 裡的人有哪些在保護名單
        to_be_protected = df[df['身分證字號'].astype(str).str.strip().str.upper().isin(protected_ids)]
        
        if not to_be_protected.empty:
            print("\n" + "!"*60)
            logger.warning(f"🚨 注意：偵測到加保名單中有 {len(to_be_protected)} 位人員處於【保護名單】中：")
            for _, p_row in to_be_protected.iterrows():
                print(f"   - [受保護] {p_row['員工姓名']} ({p_row['身分證字號']})")
            logger.warning("👉 以上人員將會被自動跳過，請確保您已完成他們的手動加保。")
            print("!"*60 + "\n")
            time.sleep(2) # 留一點時間讓你閱讀

        df['生日'] = df['生日'].apply(convert_to_roc_date)
        df['受僱日期'] = df['受僱日期'].apply(convert_to_roc_date)
        logger.info(f"Excel 讀取完成，共 {len(df)} 筆。")
    except Exception as e:
        logger.error(f"Excel 讀取失敗: {e}"); exit()

    # 2. 初始化 Driver
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    # 3. 初始化計數器
    success_count = 0
    failure_count = 0
    protected_count = 0
    empty_count = 0

    # 4. 【補強導航】只寫一次即可
    try:
        handles = driver.window_handles
        fubon_found = False
        for h in handles:
            driver.switch_to.window(h)
            if "gis.fubonlife.com.tw" in driver.current_url:
                fubon_found = True
                break
        
        if not fubon_found:
            driver.switch_to.window(handles[-1])

        logger.info(f"正在導航至: {target_url}")
        driver.execute_script(f"window.location.href = '{target_url}';")

    except Exception as e:
        logger.error(f"連線或切換分頁失敗: {e}"); exit()

    # 只保留這一個確認點
    input("👉 網頁就緒後（確認看到填寫畫面），請在此按『Enter』正式開始執行...")

    # 5. 開始迴圈處理
    for index, row in df.iterrows():
        emp_name = str(row.get('員工姓名', '')).strip()
        # 修正：比對時也轉大寫
        emp_id = str(row.get('身分證字號', '')).strip().upper()

        if not emp_name or emp_name.lower() == 'nan':
            empty_count += 1
            continue

        if emp_id in protected_ids:
            protected_count += 1
            # 改為 warning 讓顏色在終端機更顯眼
            logger.warning(f"⏭️  [跳過保護人員] {emp_name} ({emp_id})")
            continue

        logger.info(f"--- 串列 [{index+1}/{len(df)}] 處理中: {emp_name} ---")
        
        try:
            # 確保頁面在正確位置
            if target_url not in driver.current_url:
                driver.execute_script(f"window.location.href = '{target_url}';")
            
            name_xpath = "//label[contains(., '員工姓名')]/ancestor::div[contains(@class, 'ant-row')]//textarea"
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, name_xpath)))
            
            fill_fubon_enrollment(driver, row)
            success_count += 1
            logger.info(f"🎉 處理完成: {emp_name}")

        except Exception as e:
            failure_count += 1
            logger.error(f"❌ 填寫失敗: {emp_name}")
            print(f"\n⚠️ 【錯誤發生】人員：{emp_name}\n原因: {str(e)[:150]}")
            input("👉 請在網頁手動點回「加保申請頁面」後，在此按『Enter』繼續下一筆...")
        
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