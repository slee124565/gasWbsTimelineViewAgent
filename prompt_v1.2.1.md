**專案目標：**

基於現有的 `code.js` 軟體設計，以及 `gas-spec-v1.2.0.md` 中定義的 `Column Structure & Automation` 資料集結構，為 WBS (Work Breakdown Structure) 系統擴充 Web 存取功能。

**主要交付項目：**

1.  一個能提供 WBS 資料集的 **Web API**，支援 `JSON` 和 `CSV` 兩種輸出格式。
2.  一個**多頁面 Web 應用程式**框架，並優先實作第一個頁面：一個視覺化的**任務儀表板 (Timeline Dashboard)**，用於展示任務的執行狀態。

---

### **第一部分：後端與 API 設計規格**

#### **1. 安全性與存取權限**
*   **存取層級**: Web API 與網頁儀表板的存取權限，應設定為「**限定組織成員**」。此設定需在部署 Web App 時，於「誰可以存取」選項中配置為「[您的組織名稱] 中的任何人」。

#### **2. 資料來源**
*   **目標工作表**: 所有資料讀取操作應固定指向名稱為 `wbs` 的單一工作表。

#### **3. Web App 路由邏輯**
*   **單一進入點**: 使用 `doGet(e)` 函數作為所有請求的統一入口。
*   **API 路由**:
    *   當 URL 包含 `?output=json` 參數時，呼叫 `getWbsDataAsJson('wbs')` 函數。
    *   當 URL 包含 `?output=csv` 參數時，呼叫 `getWbsDataAsCsv('wbs')` 函數。
*   **網頁路由**:
    *   當 URL 包含 `?page=timeline_dashboard` 參數，或不含任何可識別的路由參數時，預設呼叫 `showTimelineDashboard('wbs')` 函數。

#### **4. API 功能實作**
*   **`getWbsDataAsJson(sheetName)`**: 讀取 `wbs` 工作表的所有資料，將其轉換為 JSON 物件陣列，並使用 `ContentService` 以 `application/json` 格式回傳。
*   **`getWbsDataAsCsv(sheetName)`**: 讀取 `wbs` 工作表的所有資料，將其轉換為標準 CSV 格式的字串，並使用 `ContentService` 以 `text/csv` 格式回傳。

#### **5. 錯誤處理**
*   **工作表不存在**: 若目標 `wbs` 工作表不存在，API 應回傳 HTTP 狀態碼 404 (Not Found) 並提供明確的錯誤訊息；網頁儀表板則應顯示一個使用者友善的提示頁面，引導使用者先執行 WBS 系統的初始化。

---

### **第二部分：前端儀表板 (Timeline Dashboard) 設計規格**

這部分將作為一個独立的提示詞，交由負責前端開發的 LLM Agent 來完成 `TimelineDashboard.html` 檔案的建立。

#### **LLM Agent 提示詞 (Prompt for Frontend Development)**

**角色 (Role):** 你是一位專業的前端開發工程師，擅長使用純 HTML 和 CSS 打造乾淨、高效能、響應式的單一頁面應用程式。

**目標 (Objective):** 你的任務是為一個 Google Apps Script 專案，建立一個名為「Timeline Dashboard」的完整 HTML 頁面。這個頁面將用於視覺化展示 WBS (Work Breakdown Structure) 系統中的任務狀態。

**核心要求 (Core Requirements):**

1.  **單一檔案**: 所有的 HTML 和 CSS 都必須包含在同一個 `.html` 檔案中。
2.  **純粹技術**: 禁止使用任何外部 CSS 框架 (如 Bootstrap, Tailwind CSS) 或 JavaScript 函式庫 (如 jQuery, React)。
3.  **內嵌樣式**: 所有的 CSS 樣式都必須寫在 `<head>` 標籤內的 `<style>` 區塊中。
4.  **Google Apps Script 整合**: 頁面必須使用 Google Apps Script 的樣板語法 (Scriptlets) 來動態渲染資料。**禁止**使用客戶端 `<script>` 標籤。
5.  **響應式設計**: 頁面佈局在桌面瀏覽器上應為多欄式，在行動裝置等窄螢幕上應自動轉換為單欄堆疊式。

---

**詳細設計規格 (Detailed Design Specifications):**

**1. 頁面整體結構 (Overall Page Structure)**

*   使用標準的 HTML5 結構 (`<!DOCTYPE html>`, `<html>`, `<head>`, `<body>`)。
*   在 `<head>` 中必須包含 `<base target="_top">` 以確保在 Google 環境中正常運作。
*   頁面主體應包含一個**頁首區 (Header Section)** 和一個**內容區 (Content Section)**。

**2. 頁首區 (Header Section)**

*   包含一個 `<h1>` 標題，內容為：「SWD OKR Dashboard」。
*   在標題下方，包含一個 `<p>` 標籤，用於顯示最後更新時間。請使用樣板語法輸出一個名為 `lastUpdated` 的變數：`<p>最後更新時間: <?= lastUpdated ?></p>`。

**3. 後端資料處理邏輯 (Backend Data Processing - `showTimelineDashboard` function)**

*   **日期基準**:
    *   `today` (今日)
    *   `fourWeeksAgo` (今日 - 28天)
    *   `fourWeeksHence` (今日 - 28天)
*   **資料篩選與處理**: 後端需準備好以下三個資料陣列，並將它們傳遞給 HTML 樣板。在傳遞給前端前，所有日期欄位（如 `DueDate`）應格式化為 `YYYY-MM-DD` 的字串。
    1.  **`overdueTasks`**: `DueDate` < `today` **且** `TaskStatus` **不是** "Done" **也** **不是** "Blocked"。**排序方式**：首先依 `Object` **升序**，然後在此基礎上再依 `DueDate` **升序**。
    2.  **`futureTasks`**: `DueDate` >= `today` **且** `DueDate` <= `fourWeeksHence` **且** `TaskStatus` **不是** "Done"。**排序方式**：首先依 `Object` **升序**，然後在此基礎上再依 `DueDate` **升序**。
    3.  **`pastTasks`**: `DueDate` >= `fourWeeksAgo` **且** `DueDate` < `today` **且** `TaskStatus` **是** "Done"。**排序方式**：首先依 `Object` **升序**，然後在此基礎上再依 `DueDate` **降序**。

**4. 內容區：三欄式卡片佈局 (Content Section: 3-Column Card Layout)**

*   建立一個容器 `div` (例如，`class="card-container"`) 來包裹三個任務卡片。
*   使用 CSS Flexbox 來實現佈局：
    *   在寬螢幕上，卡片水平排列 (`flex-direction: row`)。
    *   使用 `@media (max-width: 768px)` 媒體查詢，在窄螢幕上將卡片變為垂直堆疊 (`flex-direction: column`)。

**5. 任務狀態標籤樣式 (Task Status Label Styles)**
*   **通用樣式**: `Status` 欄位中的文字應包裹在 `<span>` 元素中，並應用 `.status` CSS 類別，使其呈現膠囊狀、白色文字、粗體、置中對齊且有最小寬度。
*   **顏色定義**:
    *   `Done`: 應用 `.status-done` 類別，背景色為綠色 (`#4CAF50`)。
    *   `InProgress`: 應用 `.status-inprogress` 類別，背景色為琥珀色 (`#FFC107`)，文字顏色為深色 (`#333`)。
    *   `NotStarted`: 應用 `.status-notstarted` 類別，背景色為灰色 (`#9E9E9E`)。
    *   `Blocked`: 應用 `.status-blocked` 類別，背景色為紅色 (`#F44336`)。

**6. 卡片設計 (Card Design)**

你需要建立**三個**外觀相似但顏色和內容不同的卡片。

*   **卡片 1: 已過期任務 (Overdue Tasks)**
    *   **數據來源**: `overdueTasks` 物件陣列。
    *   **卡片標題**: 背景色為淡紅色 (`#FFCDD2`)，內容為：`⚠️ 已過期任務 <span class="badge"><?= overdueTasks.length ?></span>`。
    *   **卡片內容**: 如果陣列為空，顯示「太棒了！目前沒有任何過期的任務。」；否則，顯示一個包含 `Object`, `TaskDescription-1`, `Resource`, `DueDate` 欄位的表格。`DueDate` 欄位文字需為**粗體紅色**，且**表格內容應按 `Object` 屬性進行分組顯示。**

*   **卡片 2: 未來四周計畫 (Upcoming Tasks)**
    *   **數據來源**: `futureTasks` 物件陣列。
    *   **卡片標題**: 背景色為淡藍色 (`#BBDEFB`)，內容為：`🚀 未來四周計畫 <span class="badge"><?= futureTasks.length ?></span>`。
    *   **卡片內容**: 如果陣列為空，顯示「未來四周沒有已排定的任務。」；否則，顯示一個包含 `Object`, `TaskDescription-2`, `Status`, `StartDate`, `DueDate` 欄位的表格。**表格內容應按 `Object` 屬性進行分組顯示，且 `Status` 欄位應根據其值動態應用上述狀態顏色樣式。**

*   **卡片 3: 最近四周成果 (Recent Accomplishments)**
    *   **數據來源**: `pastTasks` 物件陣列。
    *   **卡片標題**: 背景色為淡綠色 (`#C8E6C9`)，內容為：`✅ 最近四周成果 <span class="badge"><?= pastTasks.length ?></span>`。
    *   **卡片內容**: 如果陣列為空，顯示「最近四周沒有已完成的任務。」；否則，顯示一個包含 `Object`, `TaskDescription-2`, `Status`, `StartDate`, `DueDate`, `DoneDate` 欄位的表格。**表格內容應按 `Object` 屬性進行分組顯示，且 `Status` 欄位應根據其值動態應用上述狀態顏色樣式。**

**7. Google Apps Script 樣板語法實作 (Scriptlet Implementation)**

*   使用 `<? for (...) { ... } ?>` 遍歷任務陣列並生成表格列。
*   使用 `<?= ... ?>` 輸出任務的具體屬性值。
*   使用 `<? if (...) { ... } else { ... } ?>` 處理「空狀態」的顯示邏輯。
 