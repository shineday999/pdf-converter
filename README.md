# Sabre PDF Converter

本機安全 PDF 多格式轉換工具，透過 Google Apps Script (GAS) 部署。

## 功能
- 全瀏覽器端轉換，**零上傳原始 PDF**
- 支援格式：TXT / PNG / JPG / DOCX / XLSX / HTML / CSV
- 圖片可選解析度：300 DPI / 150 DPI
- 支援頁碼範圍自訂
- 多檔批次轉換，各自獨立輸出
- 轉換完成自動存入 Google Drive + Email 寄送

## GAS Script Properties 設定（必填）

在 GAS 專案 → 專案設定 → Script 屬性中新增：

| Key | 說明 | 必填 |
|-----|------|------|
| `LOG_SHEET_ID` | 記錄用 Google Sheet 的 ID | ✅ |
| `OUTPUT_FOLDER_ID` | 指定 Drive 輸出資料夾 ID（不填則自動建立） | 選填 |
| `DEFAULT_RECIPIENT` | 預設收件人 Email | 選填 |

## clasp 部署

```bash
# 登入
clasp login

# 建立新 GAS 專案（第一次）
clasp create --type webapp --title "Sabre PDF Converter"

# 推送程式碼
clasp push

# 部署為 Web App
clasp deploy --description "v1.0.0"
```

## 開發本機測試

直接用瀏覽器開啟 `HTML.html`（功能有限，需在 GAS 環境才能使用 Drive/Email）。

## 版本

v1.0.0 | 2026-04-09

## 版權

Developed by Anderson | © 2026 All Rights Reserved
