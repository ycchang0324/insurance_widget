# 📋 富邦人壽 GIS 系統自動加保退保工具

本專案是一套基於 **Python** 與 **Selenium** 開發的自動化腳本，旨在協助處理富邦人壽 GIS 系統中的「員工加保」與「員工退保」流程。

> 💡 **開發備註**：本工具由 **Gemini (Google AI)** 與使用者協作開發與重整。針對富邦系統特殊的 Ant Design 與 Vue 元件特性，透過 AI 輔助優化了「日期自動轉換」及「雙重人工確認機制」，確保自動化流程的穩定性與準確性。

---
## 解決痛點

在輸入保險資料時，很常需要手動把西元出生年月日改成民國出生年月日，另外也要手動把身分證字號複製到保險網頁，有很高機率造成誤植。這套程式將把輸入保險資料變成是一直按 enter 鍵確認，不僅可節省將近 5 倍的時間，更大大減少出錯機率。

## ✨ 功能亮點

* **民國年自動轉換**：內建工具函式，將 Excel 西元日期自動轉為系統要求的 `115/04/03` 格式。
* **全流程手動確認**：
    * **啟動檢查**：程式啟動後會暫停，待人工確認網頁加載完成後按 `Enter` 才會開始。
    * **成功送出**：每筆資料填完後會停在確定按鈕前，由人工核對後按 `Enter` 才正式送出。
    * **失敗防護**：若執行中出錯，程式會即時停下並顯示原因，待人工排除後按 `Enter` 續行下一筆。

---

## 🚀 環境設定

### 1. 必要套件
請確保已安裝 Python 3.8+ 以及 Chrome 瀏覽器，並執行以下指令安裝依賴庫：
```bash
pip3 install requirements.txt
```

### 2. Chrome 遠端除錯模式
本腳本使用「接管現有瀏覽器」模式，先手動創造瀏覽器後手動登入，完成後再讓腳本進行加保退保動作。
1. 關閉所有 Chrome 視窗。
2. 透過終端機啟動具備偵測埠的 Chrome：
   * **macOS**:
       ```bash
       /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir="~/ChromeProfile"
       ```
   * **Windows**:
       ```cmd
       chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\AutomationProfile"
       ```

---

## 📂 資料夾架構 (Project Structure)

```bash
insurance/
├── 📁 data/                        # 資料存放區 (Excel 資料庫)
│   ├── 📄 enrollment.xlsx                 # [🔒 隱私] 實際加保人員名單
│   ├── 📄 enrollment_example.xlsx         # 💡 加保 Excel 範本 (供參考格式)
│   ├── 📄 surrender.xlsx              # [🔒 隱私] 實際退保人員名單
│   ├── 📄 surrender_example.xlsx      # 💡 退保 Excel 範本 (供參考格式)
│   ├── 📄 protected.xlsx          # [🔒 隱私] 保護名單 (名單內成員不可退保)
│   └── 📄 protected_example.xlsx  # 💡 保護名單範本
├── 📁 log/                         # 系統日誌紀錄
│   └── 📄 execution_enrollment.log            # 自動生成的加保執行過程與錯誤紀錄
│   └── 📄 execution_surrender.log            # 自動生成的退保執行過程與錯誤紀錄
├── 📁 src/                        # 工具程式資料夾
│   ├── 🐍 utility.py                 # 工具程式（西元民國日期轉換、等待遮罩）
├── 🐍 enrollment.py                       # 核心腳本：員工加保自動化流程
├── 🐍 surrender.py                    # 核心腳本：員工退保自動化流程
├── 📝 README.md                    # 專案說明文件與使用指南
├── 📋 requirements.txt             # 專案依賴套件清單 (pip install)
└── 🚫 .gitignore                   # Git 忽略設定 (確保敏感個資不外流)
```

## 📖 使用說明

### 第一步：準備資料
根據任務準備對應的 Excel 檔案：
* **加保**：使用 `enrollment.xlsx`（包含員工姓名、身分證字號、生日、國籍、性別、護照英文名字、受僱日期、工作內容、GADD、GMR）。

* **退保**：使用 `surrender.xlsx`（包含員工姓名、身分證字號、退保日期）。另外`protected.xlsx`（包含員工姓名、身分證字號）是受保護名單，不會被退保。 

#### 填寫規範與注意事項 (Guidelines)

🛡️ 保險金額預設：

GADD (團體意外傷害險)：預設填寫 100。

GMR (意外醫療附加條款)：預設填寫 2。

🌏 國籍與身分驗證：

本國籍：若國籍為「中華民國」，可免填「護照英文名字」。

外籍人士：若無身分證字號或居留證，系統無法受理自動化，請採取手寫保單處理。

### 第二步：啟動自動化
1. 在 Chrome 中手動登入富邦 GIS 目標功能頁面。
2. 執行腳本：

#### 加保
   ```bash
   python3 enrollment.py
   ```

#### 退保
   ```bash
   python3 surrender.py
   ```
3. 當終端機顯示 `👉 網頁就緒後，請在此按『Enter』開始執行...` 時，切換到網頁確認加載完成後按下 Enter。

### 第三步：核對與確認
* 腳本會自動填寫所有欄位並捲動到「確定」按鈕。
* **人工核對**網頁顯示資料是否與 Excel 一致。
* 在終端機按 **Enter**，程式將點擊送出並進入下一位人員。

---

### 📋 授權與免責
本腳本僅供內部行政效率提升使用。自動化執行期間請務必全程留守，並對最終送出的資料準確性負責。