# 公司筆電部署指南

這份文件是給你、公司 IT，或另一個 AI 助手直接照著部署用的。目標是在公司 Windows 筆電上把本專案跑成地端 Streamlit 工具，不需要把客戶 invoice / packing list 上傳到外部服務。

## 一句話目標

部署一個本機網頁工具：打開 `http://localhost:8501` 後，可以上傳 OP 最終格式範本與客戶 Invoice / Packing List，產生 OP 可維護的 VBA 巨集與客戶規則 JSON。

## 需要的環境

- Windows 10 或 Windows 11
- Python 3.10 以上
- Git，選用；沒有 Git 也可以從 GitHub 下載 ZIP
- 可連線到 Python 套件來源安裝 requirements；若公司有 proxy，請 IT 協助設定 pip proxy
- Microsoft Excel，最後貼上與執行 VBA 時需要

## GitHub 位置

Repo: https://github.com/ainoqqaph/VBA

## 最快部署方式

在公司筆電開 PowerShell，執行：

```powershell
git clone https://github.com/ainoqqaph/VBA.git
cd VBA
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_windows.ps1
.\scripts\run_streamlit.ps1
```

啟動成功後，瀏覽器開：

```text
http://localhost:8501
```

## 沒有 Git 的部署方式

1. 到 https://github.com/ainoqqaph/VBA
2. 按 `Code` -> `Download ZIP`
3. 解壓縮到例如 `C:\Users\<你的帳號>\Documents\VBA`
4. 在該資料夾按右鍵開 PowerShell
5. 執行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_windows.ps1
.\scripts\run_streamlit.ps1
```

## 日常啟動方式

安裝完成後，下次只需要：

```powershell
.\scripts\run_streamlit.ps1
```

也可以直接雙擊 repo 根目錄的：

```text
run_streamlit.bat
```

## 給 AI 助手的部署 Prompt

可以把下面這段直接貼給公司筆電上的 AI 助手：

```text
請幫我在這台 Windows 公司筆電部署這個 Streamlit 專案：
https://github.com/ainoqqaph/VBA

需求：
1. 不要改動公司全域 Python，請在專案內建立 .venv。
2. 優先執行 scripts/setup_windows.ps1 安裝依賴。
3. 安裝後執行 scripts/run_streamlit.ps1。
4. 確認 http://localhost:8501 可以開啟。
5. 若公司電腦不能執行 PowerShell script，請用同等指令手動執行：
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   .\.venv\Scripts\python.exe -m streamlit run app.py --server.address localhost --server.port 8501
6. 不要把客戶檔案上傳到外部網站；這個工具只需要地端 localhost。
```

## 驗證清單

安裝完成後請確認：

- PowerShell 顯示 `Environment check passed`
- `http://localhost:8501` 可以打開
- Streamlit 畫面有 OP 範本上傳區
- 可以看到 TINV / TPKG 分頁
- TPKG 有 `NW / GW 重量推算` 選項
- 側邊欄有 `客戶規則分類`

## 常見問題

### Python 找不到

請安裝 Python 3.10 以上，或請 IT 確認 `python` / `py` 指令可以在 PowerShell 使用。

### PowerShell 不允許執行 script

先在同一個 PowerShell 視窗執行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

這只影響目前視窗，不會永久改公司電腦政策。

### pip 被公司網路擋住

請 IT 提供公司 proxy 或內部 PyPI mirror。安裝指令可改成：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt --proxy http://proxy.example.com:8080
```

### 8501 port 被占用

改用其他 port：

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py --server.address localhost --server.port 8502
```

然後打開 `http://localhost:8502`。

### `.doc` 讀不到

舊版 Word `.doc` 請先另存成 `.docx`。掃描圖片 PDF 目前仍需要先 OCR 或轉成可讀表格。

## 資料安全說明

這個部署方式跑在公司筆電本機：

- Streamlit 網址是 `localhost`
- 客戶檔案只在瀏覽器與本機 Python 程式中處理
- 預設不需要外部 API
- 只有安裝 Python 套件時需要連網
