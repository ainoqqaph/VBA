# Invoice / Packing List VBA Generator

這是一個 Streamlit 小工具，用來協助報關 invoice / packing list 清洗流程。

目前版本可以：

- 上傳 Excel、CSV、PDF、Word `.docx`
- 預覽讀到的 sheet、PDF 表格、Word 表格
- 指定表頭列與資料開始列
- 將客戶欄位對應成你需要的最終欄位
- 原始資料沒有分類標題時，可固定補入例如 `HAND TOOL`
- 產生可貼進 Excel VBA 的清洗巨集

> 舊版 `.doc` 請先轉成 `.docx`。掃描圖片型 PDF 之後可再加 OCR。

## 安裝

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

## 後續要補的正式規則

等你提供現有 VBA 範例後，可以把以下邏輯加進 `invoice_packing_cleaner/vba_generator.py`：

- 固定輸出欄位順序
- Invoice 與 packing list 分開時的合併鍵值
- 重量、件數、材積、幣別、總金額的特殊解析
- 不同客戶格式的規則模板
- PDF/Word 轉 Excel 後的中繼表輸出
