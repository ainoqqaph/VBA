# VBA

Invoice / Packing List VBA Generator 是一個 Streamlit 小工具，用來協助報關 invoice / packing list 清洗流程。

目前版本可以：

- 上傳 OP 以前手打好的最終輸出格式 Excel/CSV
- 從同一份最終格式範本中，分別選擇 `TINV` 與 `TPKG` 的表頭列
- 自動偵測 `TINV` / `TPKG` 各自的欄位名稱與輸出欄位位置
- 上傳 Excel、CSV、PDF、Word `.docx`
- 預覽讀到的 sheet、PDF 表格、Word 表格
- 分別指定 Invoice 與 Packing List 的表頭列與資料開始列
- 分別將客戶欄位對應成 `TINV` 與 `TPKG` 欄位
- 原始資料沒有分類標題時，可固定補入例如 `HAND TOOL`
- TPKG 可設定 `NW` / `GW` 是單箱重量或總重量推算
- 客戶規則可依直接客戶、終端客戶、規則分類與標籤保存
- 產生 OP 既有程式可接收的 `INVcollection` / `PKGcollection`
- 依固定順序輸出 `tempArray`
- 也可切換成直接輸出 `Tinv` / `Tpkg` 工作表的 VBA
- 下載客戶規則 JSON，下次同格式可直接匯入重用

> 舊版 `.doc` 請先轉成 `.docx`。掃描圖片型 PDF 之後可再加 OCR。

## 安裝

公司筆電建議先看：

- `DEPLOYMENT_FOR_COMPANY_LAPTOP.md`
- `scripts/setup_windows.ps1`
- `scripts/run_streamlit.ps1`
- `run_streamlit.bat`

最快方式：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_windows.ps1
.\scripts\run_streamlit.ps1
```

一般手動安裝方式：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 啟動

```powershell
streamlit run app.py
```

啟動後在瀏覽器打開 Streamlit 顯示的網址。

## 建議使用方式

1. 先上傳 OP 手打過的最終輸出格式檔，裡面可以同時有 `TINV` 和 `TPKG` 工作表。
2. 分別選擇 `TINV` 最終格式與 `TPKG` 最終格式。
3. 再上傳客戶原始 invoice / packing list。
4. 在 `TINV` 分頁選 invoice 來源，在 `TPKG` 分頁選 packing list 來源。
5. 分別指定來源表頭列、來源資料開始列、欄位對應。
6. 下載同一份 VBA 和客戶規則 JSON。

客戶很多時，不建議把格式寫死在 Python 裡。每個客戶或每種版型應下載一份規則 JSON；下次遇到同格式時直接匯入 JSON，就能重用欄位位置與來源對應。

## OP tempArray 輸出

預設 VBA 輸出模式是 `OP tempArray / Collection 格式`。產生的 VBA 會建立：

- `INVcollection`
- `PKGcollection`
- `Main`
- `TINV`
- `CollectDataTinv`
- `TPKG`
- `CollectDataTpkg`
- `ExportOPTempArrayPreview`

Invoice 每筆資料會固定打包成：

```vb
tempArray = Array(i, Itemno, Po, QtyArray, Upce, Amt, DscArray, Title, HS, BRAND, CusItemNo)
```

Packing List 每筆資料會固定打包成：

```vb
tempArray = Array(i, Itemno, Po, QtyArray, NW, GW, DscArray, MS)
```

TPKG 的 `NW` / `GW` 可依客戶規則設定：
- 來源是單箱/單件重量：直接放進 `tempArray`
- 來源是總淨重/總毛重：先除以箱數，再放進 `tempArray`

規則 JSON 也會保存分類資料，方便 400 多個客戶再依「終端客戶 / 買方 / 品牌」、「規則分類」、「搜尋標籤」整理。

OP 既有的後段寫入程式可以沿用這兩個 Collection。

如果 OP 想先檢查抓到的資料，可以執行 `ExportOPTempArrayPreview`，系統會輸出：

- `Debug_INVcollection`
- `Debug_PKGcollection`

這兩張表會把巢狀陣列攤平成可讀文字，方便非 VBA 人員檢查。

範例 VBA 可參考 `vba_snippets/op_temp_array_framework_generated.bas`。實際客戶仍建議由 Streamlit 對欄位後下載，才會帶入正確欄位位置。

## 後續要補的正式規則

等你提供現有 VBA 範例後，可以把以下邏輯加進 `invoice_packing_cleaner/vba_generator.py`：

- 固定輸出欄位順序
- Invoice 與 packing list 分開時的合併鍵值
- 重量、件數、材積、幣別、總金額的特殊解析
- 不同客戶格式的規則模板
- PDF/Word 轉 Excel 後的中繼表輸出
