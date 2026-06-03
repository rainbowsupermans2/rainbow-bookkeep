# Apps Script 備份

這個資料夾備份 Google Apps Script 的程式碼，**僅作版本對照與安全備份用**。

## ⚠️ 重要：不會自動部署

Apps Script 真實運作的版本住在 **Google 伺服器（script.google.com）**，這個 repo 只是**靜態備份**。

每次更新流程：
1. 我提供新版 code（在 chat 訊息）
2. 你**手動**複製貼到 Apps Script 編輯器
3. 你**手動**部署新版本
4. 同時 push 一份到這裡作備份對照

## 檔案

- `Code.gs` — 完整 Apps Script 程式碼（最新版）

## 部署流程

1. 打開 Google Sheets → 擴充功能 → Apps Script
2. 全選編輯器內容（Ctrl+A）→ 刪除
3. 把這裡的 `Code.gs` 內容整段複製貼上
4. Ctrl+S 存檔
5. 部署 → 管理部署 → 編輯 → 新版本 → 部署

## 版本歷史

| 版本 | 日期 | 主要更新 |
|---|---|---|
| v3.22 | 2026/06/03 | 旅遊明細「信用卡 payment」動態納入負債（避免 CSV paste 繞過 form 不同步）|
| v3.21 | 2026/05/07 | 玉山日幣 / 富邦美金 / 旅遊基金 多 cash row 支援；旅遊頁付款自動扣現金 |
| v3.20 | 2026/05/04 | 富邦美金加入 cashAccounts；dailyAssetUpdate 含分期未到期 |
| v3.19 | 2026/05/03 | 分期負債 8 欄 schema；房貸 amortization；自動 trigger |
| v3.18 | 2026/05/02 | 現金（活存）即時加減；dailyAssetUpdate force 參數 |
| v3.17 | 2026/05/02 | 活存合計 / 4 欄資產設定 |
| v3.16 | 2026/05/02 | 月度結算 K 欄還卡費；BUDGET_DEF 對齊前端 |
| v3.14 | 2026/05/01 | 負債管理動態化；信用卡視為負債 |
| v3.13 | 2026/04/28 | 薪資切齊法；月底結算 |
