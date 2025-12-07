# CRM System (Google Apps Script)

這是一個基於 Google Apps Script (GAS) 開發的客戶關係管理 (CRM) 系統。本專案包含前端 HTML 介面與後端 GAS 邏輯，可部署為 Web App 使用。

## 專案結構

- **核心邏輯**
  - `程式碼.js`: 後端主要邏輯與 API 接口。
  - `appsscript.json`: GAS 專案配置檔。

- **前端頁面**
  - `index.html`: 主頁面入口。
  - `Dashboard.html`: 儀表板頁面。
  - `MemberList.html`: 會員列表管理頁面。
  - `AdminUsers.html`: 管理員使用者管理頁面。
  - `Reconciliation.html`: 對帳相關頁面。

- **資源與組件**
  - `Stylesheet.html`: CSS 樣式定義。
  - `JavaScript.html`: 前端 JavaScript 邏輯。

## 安裝與使用說明

本專案建議使用 Google 官方的命令行工具 `clasp` 進行管理與部署。

### 1. 環境準備

確保您的電腦已安裝 [Node.js](https://nodejs.org/)。

然後安裝 `clasp`：

```bash
npm install -g @google/clasp
```

登入您的 Google 帳號：

```bash
clasp login
```

### 2. 初始化專案

您可以選擇建立一個新專案或將代碼推送到現有專案。

**選項 A：建立新專案**

在專案目錄下執行：

```bash
clasp create --type webapp --title "CRM System"
# 選擇 scriptID 對應的父層 Sheet (如果是綁定腳本) 或直接建立獨立腳本
```

**選項 B：連接現有專案**

如果您已經有一個 GAS 專案，請取得 Script ID (從 專案設定 -> Script ID 複製)，然後執行：

```bash
clasp clone <Script ID>
```

### 3. 推送代碼

將本地端的代碼推送到 Google Apps Script 雲端：

```bash
clasp push
```

### 4. 部署 Web App

代碼推送成功後，您需要將其部署為 Web App 才能看到前端介面。

1. 到 [Apps Script Dashboard](https://script.google.com/) 開啟您的專案。
2. 點擊右上角的「部署」 (Deploy) -> 「新增部署」 (New deployment)。
3. 選擇類型為「網頁應用程式」 (Web app)。
4. 設定以下權限 (依需求調整)：
   - **執行身分 (Execute as)**: Me (您的帳號)
   - **誰可以存取 (Who has access)**: 根據需求選擇 (例如：Anyone with Google account 或 Anyone)
5. 點擊「部署」 (Deploy) 並複製生成的網址 (Web App URL)。

## 注意事項

- 請確保 `appsscript.json` 中的時區與依賴庫設定正確。
- 修改代碼後，請記得再次執行 `clasp push` 更新雲端代碼。
- 若更動到 Web App 的邏輯，通常需要建立新的部署版本 (Manage deployments -> Edit -> New Version) 才能生效。
