# 名單整理工具

這是一個可在 **Windows / macOS** 執行的本機端 Python GUI 工具。

## 功能
- 合併多份 `G1_*.xlsx / G2_*.xlsx ...` 名單
- 依 `G1 > G2 > G3 ...` 優先權去重
- 若同一人重複出現在不同群組，會採用 **高優先權主檔**，但可用低優先權資料 **補齊空白欄位**
- 依近年受訂名單排除潛在客戶
- 產生手機號碼 Template CSV，供上傳至短網址平台
- 匯入平台回傳的短網址結果，支援多次匯入追蹤點擊次數成長
- 所有處理都在本機端完成

## 目前支援的主要欄位
- 編號
- 姓名
- ONE ID 序號
- LINEID
- 手機(和泰會員)電銷
- 手機(CR)電銷
- 手機(和泰會員)SMS
- 手機(CR)SMS
- 訂單
- 現保有車款_T
- 現保有車交車年份_T
- 現保有車款_L
- 現保有車交車年份_L

## 輸出檔案工作表
- `Working_List`：目前對外使用的名單
- `Merged_Master`：合併後原始版本
- `Dropped_Duplicates`：被捨棄的重複資料
- `Filtered_DNC`：因標示不聯繫而被過濾的資料
- `Removed_RecentOrders`：因近年受訂被排除的資料
- `ShortURL_Log`：短網址配對紀錄
- `Manifest`：本次使用到的來源檔案資訊

## 安裝
```bash
python -m venv .venv
source .venv/bin/activate   # Windows 改用 .venv\Scripts\activate
pip install -r requirements.txt
python lead_list_tool.py
```

## 打包
### Windows
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile lead_list_tool.py
```

### macOS
```bash
pip install pyinstaller
pyinstaller --windowed --onefile lead_list_tool.py
```

## 建議使用方式

### 1. 合併檔案
建議每次都重新選擇 **完整的 G 檔案資料夾** 重新計算，這樣最安全。

原因：
- 避免有人更新過舊檔但檔名沒變
- 避免重跑同一批資料時重複累加
- 可以穩定維持 idempotent 行為

合併時可選擇不聯繫過濾模式（不過濾 / 電話 / SMS / 任一）。電話與 LINE ID 皆為空的記錄會直接排除。

### 2. 去除近年訂過車名單
選擇含有受訂名單的資料夾，系統會依手機號碼比對並移除。

### 3. 產生手機號碼 Template
從 Working_List 匯出 SMS 手機號碼（`手機(CR)SMS` 優先）為 CSV，格式如下：
```
Phone
0987654321
0912345678
```
上傳至平台後，平台會回傳含短網址的 CSV。

### 4. 匯入短網址結果
選擇平台回傳的 CSV（格式：`No, Phone, Url, Count`），輸入欄位名稱後匯入。
- 系統自動產生 `{名稱}` 和 `{名稱}_次數` 兩欄
- URL 欄已有值則不覆蓋（同一活動 URL 固定）
- 次數欄每次更新（追蹤點擊成長）

## 去重規則
同一人若在不同名單重複出現，會用以下識別資料交叉比對：
- 手機(CR)SMS
- 手機(和泰會員)SMS
- 手機(CR)電銷
- 手機(和泰會員)電銷
- ONE ID 序號
- LINEID

只要任一識別資料相同，就會視為同一人，保留最高優先權那筆。

## 注意
- 目前以 `.xlsx` 為主，若來源是 `.xls`，建議先另存成 `.xlsx`
- 若欄位名稱之後有更多變形，可以再補 alias 規則
