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
DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "FubonDownload")
DOWNLOAD_DIR = os.getenv("FUBON_DOWNLOAD_PATH", DEFAULT_PATH)
TARGET_FILENAME = "query.xlsx"
FULL_TARGET_PATH = os.path.join(DOWNLOAD_DIR, TARGET_FILENAME)

ENROLL_FILE = "./data/enrollment.xlsx"
PROTECTED_FILE = "./data/protected.xlsx"

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# Log 配置 (確保會寫入檔案與終端機)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("./log/execution_query.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
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
        # 讀取並排除 nan
        df_source = pd.read_excel(ENROLL_FILE).dropna(subset=['員工姓名', '身分證字號'], how='any')
        df_query = pd.read_excel(FULL_TARGET_PATH).dropna(subset=['被保險人姓名', '身分證字號/居留證號碼'], how='any')
        
        try:
            df_prot = pd.read_excel(PROTECTED_FILE).dropna(subset=['身分證字號'])
            prot_ids = set(df_prot['身分證字號'].astype(str).str.strip().tolist())
        except:
            prot_ids = set()

        p_list, s_list, f_list = [], [], []

        for _, row in df_source.iterrows():
            name = str(row.get('員工姓名', '')).strip()
            id_no = str(row.get('身分證字號', '')).strip()
            if not name or name.lower() == 'nan': continue

            # 遮罩後的身分證 (用於顯示)
            display_id = mask_id(id_no)

            # 1. 保護名單檢查
            if id_no in prot_ids:
                p_list.append(f"{name} ({display_id})")
                continue

            # 2. 異動結果比對
            match_type = "加保" if mode == "加保" else "退保"
            query_matches = df_query[df_query['作業別'].str.contains(match_type, na=False)]
            
            is_success = False
            for _, q_row in query_matches.iterrows():
                q_name = str(q_row.get('被保險人姓名', '')).strip()
                q_id_masked = str(q_row.get('身分證字號/居留證號碼', '')).strip()
                if q_name.lower() == 'nan': continue

                if name == q_name and check_id_match(id_no, q_id_masked):
                    is_success = True
                    break
            
            if is_success:
                s_list.append(f"{name} ({display_id})")
            else:
                f_list.append(f"{name} ({display_id})")

        # --- 顯示與 Log 記錄結果 ---
        summary_title = f"\n{'='*20} {today_str} 【{mode}】比對報表 {'='*20}"
        logger.info(summary_title)
        
        prot_msg = f"🛡️  保護名單 (跳過): {len(p_list)} 位"
        logger.info(prot_msg)
        for p in p_list: logger.info(f"   - [保護] {p}")
        
        succ_msg = f"✅ 成功完成{mode}: {len(s_list)} 位"
        logger.info(succ_msg)
        for s in s_list: logger.info(f"   - [成功] {s}")
        
        fail_msg = f"❌ 未成功{mode} (查無記錄): {len(f_list)} 位"
        logger.info(fail_msg)
        for f in f_list: logger.info(f"   - [失敗] {f}")
        
        logger.info("="*60 + "\n")

    except Exception as e:
        logger.error(f"❌ 比對出錯: {e}")

# --- 5. 主流程 ---

if __name__ == "__main__":
    target_url = "https://gis.fubonlife.com.tw/gis-co-web/employeeFamilyToday/employeeFamilyTodayResultTable"
    
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    try:
        handles = driver.window_handles
        found = False
        for h in handles:
            driver.switch_to.window(h)
            if "fubonlife.com.tw" in driver.current_url:
                found = True; break
        if not found: driver.switch_to.window(handles[-1])

        logger.info(f"🚀 正在強制跳轉至: {target_url}")
        driver.execute_script(f"window.location.href = '{target_url}';")

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "query__table-wrap"))
        )
    except Exception as e:
        logger.error(f"❌ 導航失敗: {e}")

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