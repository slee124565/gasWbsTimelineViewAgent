# gasWbsTimelineViewAgent - WBS 自動化工具

這是一個 Google Apps Script 專案，旨在自動化 Google Sheets 中的工作分解結構 (WBS) 管理，簡化專案規劃與任務追蹤。

此專案的核心是**從一個預先設定好的 Google Sheets 範本開始**，該範本已包含 `wbs` 資料表、`TW_HOLIDAYS` 假日表，以及兩個預設的**時間軸 (Timeline)** 檢視：`Gantt-Chart` 與 `Resource-Chart`。

## 功能特色

*   **範本驅動**: 無需手動建立工作表，直接從功能齊全的範本開始。
*   **智慧欄位自動化**:
    *   **截止日期 (`DueDate`)**: 根據 `StartDate` 和 `WorkDays` 自動計算，並會排除 `TW_HOLIDAYS` 中定義的假日。
    *   **任務描述 (`TaskDescription-2`)**: 根據 `TaskDescription-1` 和 `Resource` 自動合併生成。
    *   **完成日期 (`DoneDate`)**: 當 `TaskStatus` 設為 `Done` 時，自動填入當天日期。
*   **自動儲存格顏色標記**:
    *   當 `Object` (A欄) 或 `Resource` (F欄) 的內容變更時，腳本會自動為該儲存格套用獨特的背景顏色，方便視覺化分組。

## 視覺化時間軸 (甘特圖)

WBS 的本質是將專案進行工作分解。透過範本中預設的 **時間軸** 功能，您可以將 WBS 內容無縫轉換為視覺化的甘特圖。

![Work Breakdown Structure](wbs.png)

### 專案甘特圖 (Project Gantt View)
使用 `Gantt-Chart` 工作表，您可以利用 `Object` 欄位來追蹤不同目標（例如：不同專案）之間的時程重疊情況。

![Project Gantt View](wbs-gantt-view.png)

### 資源檢視圖 (Resource View)
使用 `Resource-Chart` 工作表，您可以利用 `Resource` 欄位來追蹤同一個資源（例如：團隊成員）是否被指派了過多的任務。

![Resource View](wbs-resource-view.png)

---

## 安裝與使用說明

1.  **複製範本檔案**:
    *   點擊此連結來複製一份 WBS 範本到你自己的 Google Drive：
    *   [**點此複製範本**](https://docs.google.com/spreadsheets/d/1o6I1fYqk0xD9SdPk_qnZDObbcgyKA6gajaSGauc-VlM/copy)

2.  **開啟 Apps Script 編輯器**:
    *   在你複製的檔案中，點擊頂部選單的 **擴充功能 > Apps Script**。

3.  **貼上程式碼**:
    *   將本專案 `code.js` 的程式碼複製並貼到 Apps Script 編輯器中。
    *   你可以從此連結取得最新程式碼：[code.js](https://raw.githubusercontent.com/slee124565/gasWbsTimelineViewAgent/refs/heads/main/code.js)

4.  **儲存並重新整理**:
    *   儲存專案，然後重新整理你的 Google Sheet。你將會在頂部看到一個新的 **"🚀 WBS 自動化工具"** 選單。

    ![Google Sheet Custom Menu](google-sheet-cust-menu.png)

---

## 選單功能詳解

### 1. 重設任務內容與公式 (保留首欄)
*   **功能**: 清除 `wbs` 工作表中除了第一欄 (`Object`) 以外的所有任務內容，並重設自動化公式。適用於重複使用 WBS 範本。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 重設任務內容與公式 (保留首欄)**。

### 2. 套用 Object 顏色標記
*   **功能**: 手動為 `Object` 欄（A欄）中所有擁有相同值的儲存格套用相同的背景顏色。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 套用 Object 顏色標記**。

### 3. 套用 Resource 顏色標記
*   **功能**: 手動為 `Resource` 欄（F欄）中所有擁有相同值的儲存格套用相同的背景顏色。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 套用 Resource 顏色標記**。

---

## 手動測試說明

你可以依據以下步驟，驗證所有功能是否正常運作。

### 測試 1: `DueDate` 自動計算
1.  在 `TW_HOLIDAYS` 的 `A2` 填入一個日期 (例如 `2026-02-03`)。
2.  在 `wbs` 的 `D2` (`StartDate`) 填入 `2026-02-01`，`E2` (`WorkDays`) 填入 `5`。
3.  **驗證**: `I2` (`DueDate`) 應自動計算出 `2026-02-09`。

### 測試 2: `TaskDescription-2` 自動生成
1.  在 `C2` (`TaskDescription-1`) 輸入 `開發新功能`。
2.  **驗證**: `J2` 應顯示 `[未指派]-開發新功能`。
3.  在 `F2` (`Resource`) 輸入 `工程師A`。
4.  **驗證**: `J2` 應自動更新為 `[工程師A]-開發新功能`。

### 測試 3: `DoneDate` 自動更新
1.  在 `G3` (`TaskStatus`) 的下拉選單中選擇 `Done`。
2.  **驗證**: `H3` (`DoneDate`) 應自動填入今天的日期。
3.  將 `G3` 的狀態改為 `InProgress`。
4.  **驗證**: `H3` 的日期應被自動清除。

### 測試 4: `Object` 儲存格顏色標記
1.  在 `A2` 和 `A3` 輸入 `專案A`，在 `A4` 輸入 `專案B`。
2.  **驗證 (自動觸發)**:
    *   `A2` 和 `A3` 儲存格應自動變為相同的背景顏色。
    *   `A4` 儲存格應有不同的背景顏色。
3.  你也可以透過選單 **"套用 Object 顏色標記"** 來手動觸發。

### 測試 5: `Resource` 儲存格顏色標記
1.  在 `F2` 和 `F4` 輸入 `工程師A`，在 `F3` 輸入 `工程師B`。
2.  **驗證 (自動觸發)**:
    *   `F2` 和 `F4` 儲存格應自動變為相同的背景顏色。
    *   `F3` 儲存格應有不同的背景顏色。
3.  你也可以透過選單 **"套用 Resource 顏色標記"** 來手動觸發。