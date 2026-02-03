# Google Apps Script Specification: WBS Automation v1.1.0

This document outlines the revised user workflow and script functionality for `gasWbsTimelineViewAgent`, adapting to the limitation that Google Sheets' "Timeline" view cannot be programmatically created. The new process prioritizes user experience by starting from a pre-configured template.

---

## 1. Revised System Overview & User Workflow

The core automation logic remains, but the user setup process is fundamentally changed. Instead of the script generating the required sheets, the user will now **copy a master template** that already contains the necessary structure, formatting, and pre-configured Timeline views.

### User Workflow:
1.  **Copy Template**: The user begins by making a copy of the official [WBS to Gantt & Resource Chart View Template](https://docs.google.com/spreadsheets/d/1o6I1fYqk0xD9SdPk_qnZDObbcgyKA6gajaSGauc-VlM/edit?usp=sharing).
2.  **Install Script**: The user opens the Apps Script editor within their copied spreadsheet (`Extensions > Apps Script`) and pastes the content of `code.js`.
3.  **Use Automation**: After reloading the sheet, a custom menu appears, providing tools to manage the WBS data within the existing structure.

---

## 2. Core Script Components & Triggers

### 2.1. `onOpen()` - Custom Menu
*   **Purpose**: Creates a custom menu in the spreadsheet UI when the document is opened.
*   **Execution**: The menu is now labeled **"🚀 WBS 自動化工具"** and contains simplified options reflecting the new workflow.
    1.  **重設任務內容與公式 (Reset Task Content & Formulas)**: Clears existing task data and resets formulas.
    2.  **套用 Object 顏色標記 (Apply Object Color-Coding)**: Applies color-coding to the `Object` column.
    3.  **套用 Resource 顏色標記 (Apply Resource Color-Coding)**: Applies color-coding to the `Resource` column.
*   **Removed Functionality**: The `initializeWBSSystem()` function and its corresponding menu item ("建立新 WBS 工作表") are now obsolete and have been removed, as the template provides the necessary sheets.

### 2.2. `onEdit(e)`
*   **Purpose**: An automatic trigger that runs whenever a user edits a cell.
*   **Execution**:
    *   Continues to monitor the `TaskStatus` column to automatically manage the `DoneDate`.
    *   Continues to silently trigger the color-coding logic when a value in the `Object` or `Resource` column is changed.

### 2.3. `resetWBSContentFormulas()`
*   **Purpose**: A consolidated function to clear task content and reset automated formulas, preparing the sheet for reuse.
*   **Execution Logic**:
    *   This function no longer needs to create headers or apply initial formatting.
    *   It now targets the `wbs` sheet.
    *   It clears the content of all data rows (from row 2 downwards) **except for the first column (`Object`)**, which is preserved.
    *   It reapplies the necessary array formulas for `DueDate` and `TaskDescription-2` in their respective header cells (`I1` and `J1`).

### 2.4. Color-Coding Functions
*   `applyObjectColorCoding()` and `applyResourceColorCoding()` remain functionally unchanged. They continue to apply unique background colors to cells in their respective columns based on the cell's value.

---

## 3. Pre-configured Sheets in the Template

The user's copied spreadsheet will contain the following pre-configured sheets:

### 3.1. `wbs`
*   **Purpose**: The primary sheet for WBS data entry.
*   **Structure**: Contains all necessary columns with headers, data validation (for `TaskStatus`), and conditional formatting.
*   **Automation**: The script will manage formulas and color-coding within this sheet.

### 3.1.1 Column Structure & Automation

The column structure remains the same as in v0.0.4.

| Column | Name | Type | Description & Implemented Logic |
|---|---|---|---|
| A | **Object** | String | **Manual Input:** The name of the project or target. **This column now drives the row color-coding.** |
| ... | ... | ... | ... |

### 3.1.2. Formatting

*   **Header**: The first row is frozen, with a background color of `#cfe2f3` and bold text.
*   **Data Types**: Date and number columns are formatted as specified.
*   **Row Color-Coding**:
    *   All cells within a single data row will have a consistent background color.
    *   The color is determined by the value in that row's `Object` cell (Column A).
    *   All rows with the same `Object` value will share the same background color.
    *   The script uses a predefined palette of distinct colors to ensure that different `Object` values have different colors.


### 3.2. `TW_HOLIDAYS`
*   **Purpose**: Defines non-working dates (holidays) to be excluded from `DueDate` calculations.
*   **Content**: Pre-populated with a list of Taiwanese holidays, which users can customize.

### 3.3. `Gantt-Chart`
*   **Purpose**: A pre-configured **Timeline View** sheet.
*   **Configuration**: It is set up to use the data from the `wbs` sheet to generate a project-level Gantt chart, grouped by the `Object` column.

### 3.4. `Resource-Chart`
*   **Purpose**: A second pre-configured **Timeline View** sheet.
*   **Configuration**: It is set up to use the data from the `wbs` sheet to generate a resource allocation chart, grouped by the `Resource` column.

---
【END】
