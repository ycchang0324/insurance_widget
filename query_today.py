import os
import time
import logging
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 載入環境變數
load_dotenv()

# --- 1. 路徑與配置 ---
current_dir = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(current_dir, "data", "downloads")
LOG_DIR = os.path.join(current_dir, "log")

# 確保資料夾存在
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# 檔案路徑
TARGET_FILENAME = "query.xlsx"
FULL_TARGET_PATH = os.path.join(DOWNLOAD_DIR, TARGET_FILENAME)
ENROLL_FILE = os.path.join(current_dir, "data", "enrollment.xlsx")
PROTECTED_FILE = os.path.join(current_dir, "data", "protected.xlsx")
SURRENDER_FILE = os.path.join(current_dir, "data", "surrender.xlsx")

# --- 設定 Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "execution_query.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)
# 🚨 關鍵點：在這裡定義全域 logger，這樣下面的 click_and_download 等函式才抓得到它
logger = logging.getLogger(__name__)

# --- 2. 輔助工具 ---

def mask_id(id_no):
    """將身分證字號轉為 A11*****111 形式 (前三後三)"""
    s_id = str(id_no).strip().upper()
    if len(s_id) < 6:
        return s_id
    return f"{s_id[:3]}*****{s_id[-3:]}"

def check_id_match(raw_id, masked_id):
    """比較原始 ID 與遮罩 ID 是否匹配 (前二後三)"""
    if pd.isna(raw_id) or pd.isna(masked_id):
        return False
    raw = str(raw_id).strip().upper()
    masked = str(masked_id).strip().upper()
    if len(raw) < 5 or len(masked) < 5:
        return False
    return raw[:2] == masked[:2] and raw[-3:] == masked[-3:]

# --- 3. 執行下載與更名 ---

def click_and_download(driver):
    wait = WebDriverWait(driver, 20)
    try:
        if os.path.exists(FULL_TARGET_PATH):
            os.remove(FULL_TARGET_PATH)
            logger.info(f"🗑️ 已清理舊檔案: {TARGET_FILENAME}")

        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": DOWNLOAD_DIR
        })
        
        download_xpath = "//a[img[contains(@src, 'button_download')]]"
        logger.info("🔍 正在尋找下載按鈕...")
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, download_xpath)))
        driver.execute_script("arguments[0].click();", btn)
        logger.info("✅ 下載指令已送出！")

        logger.info("⏳ 等待下載完成並重新命名中...")
        for i in range(30):
            time.sleep(1)
            current_files = os.listdir(DOWNLOAD_DIR)
            if any(f.endswith('.crdownload') for f in current_files):
                continue
            
            new_xlsx = [f for f in current_files if f.endswith('.xlsx') and f != TARGET_FILENAME]
            if new_xlsx:
                newest = max([os.path.join(DOWNLOAD_DIR, f) for f in new_xlsx], key=os.path.getctime)
                try:
                    os.rename(newest, FULL_TARGET_PATH)
                    logger.info(f"💾 檔案已重新命名為: {TARGET_FILENAME}")
                    return True
                except (PermissionError, OSError):
                    continue
        return False
    except Exception as e:
        logger.error(f"❌ 下載流程出錯: {e}")
        return False

# --- 4. 交叉比對邏輯 ---
def run_comparison(mode):
    today_str = datetime.now().strftime("%Y-%m-%d")
    logger.info(f"📊 開始執行【{mode}】名單比對 - 日期: {today_str}")
    
    try:
        # 1. 讀取主要來源與 Query 結果
        CURRENT_SOURCE = ENROLL_FILE if mode == "加保" else SURRENDER_FILE
        df_source = pd.read_excel(CURRENT_SOURCE).dropna(subset=['身分證字號'])
        df_query = pd.read_excel(FULL_TARGET_PATH).dropna(subset=['身分證字號/居留證號碼'])

        # 2. 讀取保護名單
        try:
            df_prot = pd.read_excel(PROTECTED_FILE).dropna(subset=['身分證字號'])
            prot_ids = set(df_prot['身分證字號'].astype(str).str.strip().str.upper().tolist())
            logger.info(f"🛡️  載入保護名單共 {len(prot_ids)} 筆紀錄")
        except Exception as e:
            logger.warning(f"⚠️ 無法讀取保護名單: {e}")
            prot_ids = set()

        # 3. 整理 Query 資料庫
        query_db = []
        for _, q_row in df_query.iterrows():
            if mode in str(q_row.get('作業別', '')):
                query_db.append({
                    'id': str(q_row.get('身分證字號/居留證號碼', '')).strip().upper(),
                    'name': str(q_row.get('被保險人姓名', '')).strip(),
                    'action': str(q_row.get('作業別', '')),
                    'matched': False 
                })

        p_list, s_list, f_list, unknown_list = [], [], [], []

        for _, row in df_source.iterrows():
            name = str(row.get('員工姓名', '未知')).strip()
            id_no = str(row.get('身分證字號', '')).strip().upper()
            display_info = f"{name} ({mask_id(id_no)})"

            if id_no in prot_ids:
                p_list.append(display_info)
                continue 

            matches = [q for q in query_db if check_id_match(id_no, q['id'])]
            if not matches:
                f_list.append(display_info)
            elif len(matches) == 1:
                matches[0]['matched'] = True
                s_list.append(display_info)
            else:
                name_match = next((q for q in matches if q['name'] == name), None)
                if name_match:
                    name_match['matched'] = True
                    s_list.append(display_info)
                else:
                    f_list.append(f"{display_info} (ID重複且姓名不符)")

        for q in query_db:
            if not q['matched']:
                unknown_list.append(f"{q['name']} ({q['id']}) [{q['action']}]")

        # --- 顯示報表 ---
        logger.info(f"\n{'='*20} {today_str} 【{mode}】比對報表 {'='*20}")
        if p_list:
            logger.warning(f"⚠️  以下 {len(p_list)} 位人員在保護名單中，請手動確認：")
            for p in p_list: logger.warning(f"   - [待確認] {p}")
        
        if unknown_list:
            logger.error(f"🚨 異常名單 (系統多出的紀錄): {len(unknown_list)} 位")
            for u in unknown_list: logger.error(f"   - {u}")
        
        logger.info(f"✅ 自動比對成功: {len(s_list)} 位")
        if f_list:
            logger.info(f"❌ 自動比對失敗: {len(f_list)} 位")
            for f in f_list: logger.info(f"   - [未完成] {f}")
        
    except Exception as e:
        logger.error(f"❌ 比對出錯: {e}")

# --- 5. 主流程 ---
if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/employeeFamilyToday/employeeFamilyTodayResultTable"
    
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        handles = driver.window_handles
        found = False
        for h in handles:
            driver.switch_to.window(h)
            if "fubonlife.com.tw" in driver.current_url:
                found = True
                break
        if not found: driver.switch_to.window(handles[-1])

        logger.info(f"🚀 正在強制跳轉至: {target_url}")
        driver.execute_script(f"window.location.href = '{target_url}';")

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "query__table-wrap"))
        )

        print(f"\n📊 今日日期: {datetime.now().strftime('%Y-%m-%d')}")
        print("="*60)
        print(" 📢 【 今日異動查詢比對工具 】")
        print("="*60)

        input("\n👉 確認網頁已出現資料後，按『Enter』執行下載並比對...")

        if click_and_download(driver):
            choice = input("\n👉 請選擇項目 [1] 加保  [2] 退保: ")
            if choice == "1":
                run_comparison("加保")
            elif choice == "2":
                run_comparison("退保")
            else:
                print("👋 取消動作。")
        else:
            print("❌ 下載或更名失敗。")

    except Exception as e:
        # 這裡原本報錯是因為 logger 沒定義，現在定義在上面了
        logger.error(f"❌ 發生錯誤: {e}")