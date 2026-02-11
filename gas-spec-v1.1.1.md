# Google Apps Script Specification: WBS Automation v1.2.0

This document details the functionality and automation rules for the `gasWbsTimelineViewAgent` script. This version shifts the workflow from using a pre-defined template to a script-driven approach, where the WBS environment is dynamically initialized within any Google Sheet.

---

## 1. System Overview & User Workflow

The script provides a set of tools to automate the creation and management of a Work Breakdown Structure (WBS) within a Google Sheet. The workflow is initiated directly from a custom menu within the user's spreadsheet, eliminating the need for a separate template file.

### User Workflow:
1.  **Install Script**: The user opens the Apps Script editor within their spreadsheet (`Extensions > Apps Script`) and pastes the content of `code.js`.
2.  **Initialize Environment**: After reloading the sheet, a custom menu appears. The user selects **"1. 建立新 WBS 工作表"** to run the `initializeWBSSystem` function.
3.  **Automated Setup**: The script creates a new `wbs` sheet (or `wbs-1`, `wbs-2`, etc., to avoid conflicts) and a `TW_HOLIDAYS` sheet if it doesn't exist. It automatically applies all necessary headers, formatting, formulas, and data validation.
4.  **Manage Data**: The user can then manage WBS data using the other menu options for resetting content or applying visual color-coding.

---

## 2. Core Script Components & Triggers

### 2.1. `onOpen()` - Custom Menu
*   **Purpose**: Creates a custom menu in the spreadsheet UI, including the script version, when the document is opened.
*   **Execution**: The menu is labeled **"🚀 WBS 自動化工具 (v1.2.0)"** and provides access to all major functions:
    1.  **建立新 WBS 工作表 (Create New WBS Sheet)**: Triggers `initializeWBSSystem()`.
    2.  **重設任務內容與公式 (Reset Task Content & Formulas)**: Triggers `resetWBSContentFormulas()`.
    3.  **套用 Object 顏色標記 (Apply Object Color-Coding)**: Triggers `applyObjectColorCoding()`.
    4.  **套用 Resource 顏色標記 (Apply Resource Color-Coding)**: Triggers `applyResourceColorCoding()`.

### 2.2. `initializeWBSSystem()`
*   **Purpose**: Sets up a new, fully configured WBS sheet and its dependencies. This is the primary entry point for users.
*   **Execution**:
    *   Creates a new sheet named `wbs`. If a sheet with that name already exists, it creates a sequentially numbered version (`wbs-1`, `wbs-2`, etc.).
    *   Creates a `TW_HOLIDAYS` sheet if one does not already exist.
    *   Applies headers, cell formatting (for dates and numbers), data validation rules (for the `TaskStatus` column), and freezes the header row.
    *   Sets per-cell formulas for `DueDate` and `TaskDescription-2` in the second row and uses `autoFill()` to propagate them down the entire sheet.

### 2.3. `onEdit(e)`
*   **Purpose**: An automatic trigger that runs on any cell edit to enforce data rules and automation.
*   **Execution**:
    *   **DoneDate Automation**: If a cell in the `TaskStatus` column (G) is changed to "Done", the corresponding `DoneDate` cell (H) is populated with the current date. If the status is changed to anything else, the `DoneDate` cell is cleared.
    *   **Color-Coding Automation**: If a cell in the `Object` column (A) or `Resource` column (F) is modified, the corresponding color-coding function is automatically triggered to update the cell's background color.

### 2.4. `resetWBSContentFormulas()`
*   **Purpose**: Clears all task data from the active WBS sheet and resets the automated formulas and formatting.
*   **Execution Logic**:
    *   Confirms the active sheet name starts with `wbs`.
    *   Clears the content of **all** data rows (from A2 downwards), including the `Object` column.
    *   Clears the background colors from all data rows.
    *   Re-applies the per-cell formulas for `DueDate` and `TaskDescription-2` and propagates them down the sheet.

### 2.5. Color-Coding Functions (`applyObjectColorCoding()` & `applyResourceColorCoding()`)
*   **Purpose**: To visually group related items by applying background colors to specific cells.
*   **`applyObjectColorCoding()`**: Assigns a unique background color to each unique value in the `Object` column (Column A). All cells in Column A with the same text will have the same background color.
*   **`applyResourceColorCoding()`**: Assigns a unique background color to each unique value in the `Resource` column (Column F). All cells in Column F with the same text will have the same background color.

---

## 3. Script-Generated Sheets

The script dynamically creates and configures the following sheets as needed.

### 3.1. `wbs` (or `wbs-x`)
*   **Purpose**: The primary sheet for WBS data entry, created by `initializeWBSSystem`.
*   **Structure**: The script sets up all necessary columns, headers, and data validation rules for `TaskStatus`.
*   **Automation**: The script manages all formulas and color-coding within this sheet.

### 3.1.1 Column Structure & Automation

The column structure is defined and created by the `initializeWBSSystem` function.

| Column | Name | Type | Description & Implemented Logic |
|---|---|---|---|
| A | **Object** | String | **Manual Input:** The name of the project or target. **This column's cells are color-coded by `applyObjectColorCoding`.** Note: Content is cleared by `resetWBSContentFormulas`. |
| B | **TaskTitle** | String | **Manual Input:** The title of the specific task. |
| C | **TaskDescription-1** | String | **Manual Input:** Detailed description of the task. |
| D | **StartDate** | Date | **Manual Input:** The planned start date of the task. |
| E | **WorkDays** | Integer | **Manual Input:** The number of working days required. |
| F | **Resource** | String | **Manual Input:** The person assigned to the task. **This column's cells are color-coded by `applyResourceColorCoding`.** |
| G | **TaskStatus** | Enum | **Data Validation:** A dropdown list restricts input to `NotStarted`, `InProgress`, `Done`, or `Blocked`. |
| H | **DoneDate** | Date | **Automated by `onEdit` trigger:** Automatically populated with the current date when `TaskStatus` is set to `Done`; cleared otherwise. |
| I | **DueDate** | Date | **Automated by Formula:** The formula `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, TW_HOLIDAYS!A$2:A), "")` is applied to each cell in this column. |
| J | **TaskDescription-2**| String | **Automated by Formula:** The formula `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")` is applied to each cell in this column. |

### 3.1.2. Formatting

*   **Header**: The first row is frozen, with a background color of `#cfe2f3` and bold text.
*   **Data Types**: Columns `D`, `H`, and `I` are formatted as `yyyy-mm-dd`. Column `E` is formatted as a number.
*   **Cell Color-Coding**:
    *   The script does **not** color entire rows.
    *   It applies background colors to individual cells in the `Object` (Column A) and `Resource` (Column F) columns.
    *   Cells with the same text value within the same column share the same color, allowing for easy visual grouping.

### 3.2. `TW_HOLIDAYS`
*   **Purpose**: Defines non-working dates (holidays) used in `DueDate` calculations.
*   **Creation**: This sheet is created automatically by `initializeWBSSystem` if it does not already exist in the spreadsheet. The user is expected to populate it with relevant holiday dates.

---
【END】
