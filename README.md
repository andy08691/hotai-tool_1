# 名單整理工具

這是一個可在 **Windows / macOS** 執行的本機端 Python GUI 工具。

## 功能
- 合併多份 `G1_*.xlsx / G2_*.xlsx ...` 名單
- 依 `G1 > G2 > G3 ...` 優先權去重
- 若同一人重複出現在不同群組，會採用 **高優先權主檔**，但可用低優先權資料 **補齊空白欄位**
- 依近年受訂名單排除潛在客戶
- 依簡訊 / LINE 短網址來源檔回填到整理名單
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

## 輸出檔案工作表
- `Working_List`：目前對外使用的名單
- `Merged_Master`：合併後原始版本
- `Dropped_Duplicates`：被捨棄的重複資料
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

### 2. 更新檔案
更新模式是更新既有輸出 xlsx。

但在執行「合併檔案」時，仍建議重新選擇完整來源資料夾，讓系統重新整理整份名單，而不是只追加。

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
