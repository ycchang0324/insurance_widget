import pandas as pd
import logging
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

logger = logging.getLogger(__name__)

def convert_to_roc_date(date_val):
    """
    終極日期轉換器：支援西元物件、民國字串、Excel 數值
    輸出格式：XX/MM/DD 或 XXX/MM/DD
    """
    if pd.isna(date_val) or str(date_val).strip() == "" or str(date_val).lower() == 'nan':
        return ""
    
    try:
        # 1. 統一轉換為 pandas datetime 物件 (處理西元年最穩定的方法)
        # errors='coerce' 會把無法轉換的變成 NaT
        d_obj = pd.to_datetime(date_val, errors='coerce')
        
        if pd.notna(d_obj):
            # 情況一：成功解析為西元日期 (物件化)
            year = d_obj.year
            month = d_obj.month
            day = d_obj.day
            
            # 西元轉民國
            roc_year = year - 1911
            return f"{roc_year}/{str(month).zfill(2)}/{str(day).zfill(2)}"
        
        # 情況二：如果 pd.to_datetime 失敗 (可能是已經是民國字串 "115/04/11")
        date_str = str(date_val).strip()
        parts = []
        if '/' in date_str:
            parts = date_str.split('/')
        elif '-' in date_str:
            parts = date_str.split('-')
            
        if len(parts) == 3:
            year = int(parts[0])
            # 如果年份 < 200，視為已經是民國
            roc_year = year if year < 200 else year - 1911
            return f"{roc_year}/{str(parts[1]).zfill(2)}/{str(parts[2]).zfill(2)}"
            
        return date_str
        
    except Exception as e:
        logger.error(f"日期轉換發生異常: {date_val}, 錯誤: {e}")
        return str(date_val)

def wait_for_spinner_to_disappear(driver, timeout=15):
    """確保網頁內部的 Loading 遮罩消失"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class, 'custom__spin')]"))
        )
    except:
        pass