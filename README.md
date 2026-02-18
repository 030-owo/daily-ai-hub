# LINE Bot + Google Sheet 冰箱庫存系統（中文）

本專案提供一個 Google Apps Script 範本，讓 LINE Bot 透過 Google Sheet 管理冰箱庫存。

## 功能

- 新增：`新增 品名 數量 [單位] [YYYY-MM-DD]`
- 取出：`取出 品名 數量`
- 清單：`清單`
- 快過期：`快過期 [天數]`
- 設定提醒：`設定提醒 天數`
- 資料儲存在 Google Sheet（`Users` / `Items` / `Logs`）

## 專案檔案

- `Code.gs`：LINE Webhook、指令解析、Sheet 操作、提醒推播
- `appsscript.json`：Apps Script Manifest（時區、權限、Web App 設定）
- `web/index.html`：管理介面（儲存 Properties、初始化 Sheets）

## Script Properties（金鑰與設定）

> 所有敏感資料都透過 `PropertiesService.getScriptProperties()` 儲存，不會硬編碼。

需設定以下鍵值：

- `LINE_CHANNEL_ACCESS_TOKEN`
- `LINE_CHANNEL_SECRET`
- `SPREADSHEET_ID`

## Sheet 設計

### Users

- `userId`
- `displayName`
- `defaultReminderDays`
- `createdAt`
- `updatedAt`

### Items

- `itemId`
- `userId`
- `itemName`
- `quantity`
- `unit`
- `expiryDate`
- `createdAt`
- `updatedAt`
- `status`（`active` / `empty`）

### Logs

- `logId`
- `timestamp`
- `userId`
- `action`
- `itemName`
- `quantityChange`
- `note`
- `rawMessage`

> 可執行 `initializeSheets()` 自動建立/修正三個分頁與表頭。

## 部署清單（Deployment Checklist）

1. 建立 Google Sheet，複製 Spreadsheet ID。
2. 建立 Apps Script 專案，放入本 repo 三個檔案。
3. 在 Apps Script 介面「專案設定 > Script properties」或管理頁面填入：
   - `LINE_CHANNEL_ACCESS_TOKEN`
   - `LINE_CHANNEL_SECRET`
   - `SPREADSHEET_ID`
4. 執行 `initializeSheets()`（首次需授權）。
5. 部署為 Web App：
   - Execute as: `User deploying the app`
   - Who has access: `Anyone`
6. 取得 Web App URL，填入 LINE Developers Webhook URL。
7. 在 LINE Developers 開啟 Webhook，測試訊息：
   - `新增 牛奶 2 瓶 2026-02-20`
   - `清單`
   - `快過期 7`
8. （選用）建立時間觸發器執行 `checkAndPushExpiryReminders()`，例如每日 09:00。

## 注意事項

- Apps Script `doPost(e)` 已驗證 `X-Line-Signature`。
- 若更換 Spreadsheet，請更新 `SPREADSHEET_ID` 後重跑 `initializeSheets()`。
