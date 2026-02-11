# gasWbsTimelineViewAgent - WBS 自動化工具

這是一個 Google Apps Script 專案，旨在自動化 Google Sheets 中的工作分解結構 (WBS) 管理，簡化專案規劃與任務追蹤。

此專案的核心是**透過腳本動態建立與管理 WBS 工作表**。您不再需要從特定範本開始，而是在任何 Google Sheet 中透過此工具提供的自訂選單，即可一鍵生成所需的工作環境。

## 功能特色

*   **腳本驅動初始化**: 無需預設範本，在任何 Google Sheet 中透過選單即可建立標準化的 `wbs` 工作表。
*   **動態工作表管理**: 如果 `wbs` 工作表已存在，腳本會自動建立 `wbs-1`, `wbs-2` 等新工作表，方便您在同一個檔案中管理多個專案。
*   **智慧欄位自動化**:
    *   **截止日期 (`DueDate`)**: 根據 `StartDate` 和 `WorkDays` 自動計算，並會排除在 `TW_HOLIDAYS` 工作表中定義的假日。
    *   **任務描述 (`TaskDescription-2`)**: 根據 `TaskDescription-1` 和 `Resource` 自動合併生成。
    *   **完成日期 (`DoneDate`)**: 當 `TaskStatus` 設為 `Done` 時，自動填入當天日期。
*   **資料驗證**:
    *   `TaskStatus` 欄位提供下拉選單 (`NotStarted`, `InProgress`, `Done`, `Blocked`)，確保資料一致性。
*   **自動儲存格顏色標記**:
    *   當 `Object` (A欄) 或 `Resource` (F欄) 的內容變更時，腳本會自動為該儲存格套用獨特的背景顏色，方便視覺化分組。

## 延伸設計建議：視覺化時間軸

WBS 的核心在於工作分解。雖然本工具目前著重於 WBS 資料的自動化管理與建立，但這些資料為未來實現視覺化時間軸提供了堅實的基礎。透過 Google Sheets 內建的圖表功能或其他擴充工具，您可以將 WBS 內容無縫轉換為視覺化的甘特圖或資源分配圖。

![Work Breakdown Structure](wbs.png)

### 專案甘特圖 (Project Gantt View) 潛力

利用 `Object` 欄位，可以進一步開發功能，以視覺化方式追蹤不同目標（例如：不同專案）之間的時程重疊情況，從而在多專案環境中進行更宏觀的規劃。

![Project Gantt View](wbs-gantt-view.png)

### 資源檢視圖 (Resource View) 潛力

利用 `Resource` 欄位，可以進一步開發功能，以視覺化方式追蹤同一個資源（例如：團隊成員）是否被指派了過多的任務，從而優化資源分配與負載平衡。

![Resource View](wbs-resource-view.png)

---

## 安裝與使用說明
1.  **建立新的 Google Sheet**:
    *   在你自己的 Google Drive 中建立一份新的、空白的 Google Sheet。

2.  **開啟 Apps Script 編輯器**:
    *   在新檔案中，點擊頂部選單的 **擴充功能 > Apps Script**。

3.  **貼上程式碼**:
    *   將本專案 `code.js` 的程式碼複製並貼到 Apps Script 編輯器中。
    *   你可以從此連結取得最新程式碼：[code.js](https://raw.githubusercontent.com/slee124565/gasWbsTimelineViewAgent/refs/heads/main/code.js)

4.  **儲存並重新整理**:
    *   儲存專案，然後重新整理你的 Google Sheet。你將會在頂部看到一個新的 **"🚀 WBS 自動化工具 (v1.2.0)"** 選單。

    ![Google Sheet Custom Menu](google-sheet-cust-menu.png)

5.  **初始化 WBS 工作表**:
    *   點擊 **🚀 WBS 自動化工具 > 1. 建立新 WBS 工作表**。
    *   腳本將會自動建立 `wbs` 和 `TW_HOLIDAYS` 兩個工作表，並設定好所有欄位、格式與公式。

---

## 選單功能詳解

### 1. 建立新 WBS 工作表
*   **功能**: 這是啟動所有功能的**第一步**。它會建立一個新的 WBS 工作表 (例如 `wbs`, `wbs-1` 等) 和一個 `TW_HOLIDAYS` 假日表（如果尚不存在）。所有必要的欄位、格式、資料驗證和自動化公式都會在此步驟中設定完成。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 1. 建立新 WBS 工作表**。

### 2. 重設任務內容與公式
*   **功能**: 清除**當前啟用**的 `wbs` 工作表中**所有**的任務內容（從 A2 儲存格開始的所有資料），並重新套用自動化公式。適用於清空並重複使用某個 WBS 工作表。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 2. 重設任務內容與公式**。

### 3. 套用 Object 顏色標記
*   **功能**: 手動為**當前啟用**的 `wbs` 工作表的 `Object` 欄（A欄）中所有擁有相同值的儲存格套用相同的背景顏色。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 3. 套用 Object 顏色標記**。

### 4. 套用 Resource 顏色標記
*   **功能**: 手動為**當前啟用**的 `wbs` 工作表的 `Resource` 欄（F欄）中所有擁有相同值的儲存格套用相同的背景顏色。
*   **操作**: 點擊 **🚀 WBS 自動化工具 > 4. 套用 Resource 顏色標記**。

---

## GAS Web API 介面 (可選)

此腳本包含一個可選的 API 功能，允許外部應用程式透過 HTTP GET 請求讀取 WBS 工作表的資料。

### 如何啟用 API
1.  **部署為 Web App**:
    *   在 Apps Script 編輯器中，點擊右上角的 **部署 > 新增部署**。
    *   選擇 **"網頁應用程式"** 作為部署類型。
    *   在 **"誰可以存取"** 欄位中，根據您的安全需求選擇存取權限 (例如，若要公開存取，請選擇 `任何人`)。
    *   點擊 **"部署"**，並複製產生的 **Web App URL**。此 URL 即為您的 API 端點。

### API 端點: `doGet(e)`
*   **功能**: 讀取指定 WBS 工作表的資料，並以 JSON 格式回傳。此 API 只會回傳 `Object` 欄位（A欄）包含文字內容的資料列，自動過濾掉空行。
*   **URL 參數**:
    *   `sheetName` (可選): 指定要讀取的工作表名稱。若未提供，則預設讀取 `wbs`。
*   **使用範例**:
    *   讀取 `wbs` 工作表: `YOUR_WEB_APP_URL`
    *   讀取 `wbs-2` 工作表: `YOUR_WEB_APP_URL?sheetName=wbs-2`

### 回應格式
*   **成功**:
    ```json
    {
      "status": "success",
      "sheet": "wbs",
      "data": [
        {
          "Object": "專案A",
          "TaskTitle": "任務1",
          // ... 其他欄位
        }
      ]
    }
    ```
*   **失敗 (找不到工作表)**:
    ```json
    {
      "status": "error",
      "message": "Sheet \"wbs-not-found\" not found."
    }
    ```

---

## 手動測試說明

你可以依據以下步驟，驗證所有功能是否正常運作。

**前置步驟: 初始化工作表**
1.  確保你已依照上方「安裝與使用說明」完成設定。
2.  點擊選單 **🚀 WBS 自動化工具 > 1. 建立新 WBS 工作表**。
3.  驗證 `wbs` 與 `TW_HOLIDAYS` 兩個工作表是否已成功建立。

### 測試 1: `DueDate` 自動計算
1.  在 `TW_HOLIDAYS` 工作表的 `A2` 填入一個日期 (例如 `2026-02-03`)。
2.  在 `wbs` 工作表的 `D2` (`StartDate`) 填入 `2026-02-01`，`E2` (`WorkDays`) 填入 `5`。
3.  **驗證**: `I2` (`DueDate`) 應自動計算出 `2026-02-09`。

### 測試 2: `TaskDescription-2` 自動生成
1.  在 `wbs` 的 `C2` (`TaskDescription-1`) 輸入 `開發新功能`。
2.  **驗證**: `J2` 應顯示 `[未指派]-開發新功能`。
3.  在 `F2` (`Resource`) 輸入 `工程師A`。
4.  **驗證**: `J2` 應自動更新為 `[工程師A]-開發新功能`。

### 測試 3: `DoneDate` 自動更新
1.  在 `wbs` 的 `G2` (`TaskStatus`) 的下拉選單中選擇 `Done`。
2.  **驗證**: `H2` (`DoneDate`) 應自動填入今天的日期。
3.  將 `G2` 的狀態改為 `InProgress`。
4.  **驗證**: `H2` 的日期應被自動清除。

### 測試 4: `Object` 儲存格顏色標記
1.  在 `wbs` 的 `A2` 和 `A3` 輸入 `專案A`，在 `A4` 輸入 `專案B`。
2.  **驗證 (自動觸發)**:
    *   `A2` 和 `A3` 儲存格應自動變為相同的背景顏色。
    *   `A4` 儲存格應有不同的背景顏色。
3.  你也可以透過選單 **"3. 套用 Object 顏色標記"** 來手動觸發。

### 測試 5: `Resource` 儲存格顏色標記
1.  在 `wbs` 的 `F2` 和 `F4` 輸入 `工程師A`，在 `F3` 輸入 `工程師B`。
2.  **驗證 (自動觸發)**:
    *   `F2` 和 `F4` 儲存格應自動變為相同的背景顏色。
    *   `F3` 儲存格應有不同的背景顏色。
3.  你也可以透過選單 **"4. 套用 Resource 顏色標記"** 來手動觸發。