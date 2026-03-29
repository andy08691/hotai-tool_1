# 名單整理工具

這是一個可在 **Windows / macOS** 執行的本機端 GUI 工具，專為汽車潛在客戶名單整理與精準行銷流程設計。

---

## 下載程式（Windows）

1. 前往 GitHub repo 的 **Actions** 頁籤
2. 點選最新一次成功的 workflow run
3. 往下捲至 **Artifacts**，下載 `lead_list_tool_windows.zip`
4. 解壓縮後直接執行 `lead_list_tool.exe`，不需安裝 Python

---

## 完整操作流程

### 前置準備

- 將各優先群組的名單檔案命名為 `G1_*.xlsx`、`G2_*.xlsx`、`G3_*.xlsx`...，數字越小優先權越高，放在同一個資料夾
- 執行工具前，**請先關閉 Excel 中的輸出檔案**，否則儲存時可能顯示舊版本

---

### 步驟一：合併檔案

**按鈕：「1. 合併檔案」**

1. 點擊按鈕後，選擇輸出檔案的儲存位置（新建或選取現有 .xlsx）
2. 選擇 G 檔案所在的資料夾
3. 選擇不聯繫過濾模式：
   - **不過濾**：全部保留
   - **電話不聯繫**：電銷欄有不聯繫標記則排除
   - **簡訊不聯繫**：SMS 欄有不聯繫標記則排除
   - **任一不聯繫**：任一欄有標記則排除
4. 程式會自動合併、去重，並輸出以下工作表：
   - `Working_List`：最終名單
   - `Merged_Master`：合併後完整版本
   - `Dropped_Duplicates`：被捨棄的重複資料
   - `Filtered_DNC`：因不聯繫標記被過濾的資料

**去重邏輯：** 只要兩筆記錄有任一相同的手機號碼、ONE ID 或 LINE ID，就視為同一人，G 數字小的優先，低優先權的資料僅用來補齊空白欄位。

**自動排除：** 電話與 LINE ID 皆為空的記錄不進入任何名單。

---

### 步驟二：去除近年訂過車名單

**按鈕：「2. 去除近年訂過車名單」**

1. 選擇含有受訂名單的資料夾（支援 .xlsx 與 .csv）
2. 程式依手機號碼比對，將命中的客戶從 `Working_List` 移至 `Removed_RecentOrders` 工作表

---

### 步驟三：產生手機號碼 Template

**按鈕：「3. 產生手機號碼 Template」**

1. 選擇 Template CSV 的儲存位置
2. 程式從 `Working_List` 匯出 SMS 手機號碼（`手機(CR)SMS` 優先，次選 `手機(和泰會員)SMS`），自動去重
3. 輸出格式：
   ```
   Phone
   0987654321
   0912345678
   ```
4. 將此 CSV 上傳至短網址平台，平台會為每支手機產生專屬短網址

---

### 步驟四：匯入短網址結果

**按鈕：「4. 匯入短網址結果」**

平台回傳含短網址的 CSV（格式：`No, Phone, Url, Count`）後：

1. 輸入欄位名稱（例如：`3月簡訊短網址`）
2. 選擇平台回傳的 CSV 檔案
3. 程式自動在 `Working_List` 新增兩欄：
   - `{名稱}`：短網址（已有值則不覆蓋，同一活動 URL 固定不變）
   - `{名稱}_次數`：點擊次數（每次都更新，可追蹤成長）
4. 若要在不同時間點紀錄成長，只需輸入不同的欄位名稱（例如 `3月短網址_0312`、`3月短網址_0320`），每次都會新增新欄位，舊欄位不受影響

配對紀錄會累積寫入 `ShortURL_Log` 工作表，不會覆蓋舊紀錄。

---

### 步驟五：比對留名單

**按鈕：「5. 比對留名單」**

1. 選擇從系統匯出的留名單 .xlsx（欄位需含「聯絡電話」與「備註」）
2. 程式比對 `Working_List` 並標記：
   - **簡訊留名單**：手機號碼有比對到且備註含「精準行銷」→ `V`，否則 `X`
   - **LINE留名單**：LINE ID 有比對到（備註含 `精準行銷` 且 `LINE ID_U...`）→ `V`，否則 `X`
3. 每次執行會整批覆蓋這兩欄，確保與最新名單同步

---

## 輸出工作表說明

| 工作表 | 說明 |
|---|---|
| `Working_List` | 最終對外使用的名單，所有後續步驟都在此更新 |
| `Merged_Master` | 合併後的完整版本（含被去重前的資料） |
| `Dropped_Duplicates` | 因重複被捨棄的記錄，含保留來源說明 |
| `Filtered_DNC` | 因不聯繫標記被過濾的記錄 |
| `Removed_RecentOrders` | 因近年受訂被排除的記錄 |
| `ShortURL_Log` | 每次短網址匯入的配對紀錄（累積不清除） |
| `Manifest` | 所有來源檔案的 SHA1 與時間戳，供稽核用 |

---

## 注意事項

- 執行任何步驟前，請**先關閉 Excel 中的輸出檔案**
- 輸入檔案格式以 `.xlsx` 為主，若來源是 `.xls` 請先另存為 `.xlsx`
- 建議每次都重新選擇完整的 G 檔案資料夾重跑合併，避免使用到未更新的舊版本

---

## 開發者：從原始碼執行

```bash
python -m venv .venv
source .venv/bin/activate   # Windows 改用 .venv\Scripts\activate
pip install -r requirements.txt
python lead_list_tool.py
```

### 打包成執行檔

```bash
# Windows
pyinstaller --noconsole --onefile lead_list_tool.py

# macOS
pyinstaller --windowed --onefile lead_list_tool.py
```
