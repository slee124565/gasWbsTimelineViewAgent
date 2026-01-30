# Google Apps Script Specification: WBS Automation v0.0.3

This document details the implemented functionality and automation rules for the `gasWbsTimelineViewAgent` script, reflecting the state of `code.js` as of 2026-01-30. This version introduces a content reset feature.

---

## 1. System Overview

The script provides a set of tools to automate the creation and management of a Work Breakdown Structure (WBS) within a Google Sheet. It is designed to reduce manual data entry and ensure data consistency.

The primary user entry point is a custom menu labeled **"🚀 WBS 自動化工具"**, which is automatically added to the spreadsheet's UI when the document is opened (`onOpen` trigger). This menu provides the following options:
1.  **建立新 WBS 工作表 (Create New WBS Sheet)**: Sets up a new, fully configured WBS sheet.
2.  **清空任務內容 (保留首欄) (Clear Task Content (Keep First Column))**: Resets the current WBS sheet for reuse as a template.

---

## 2. Core Script Components & Triggers

### 2.1. `initializeWBSSystem()`

*   **Purpose**: This function, triggered from the custom menu, sets up the entire WBS environment.
*   **Execution**:
    1.  Creates a new `wbs` worksheet. If `wbs` already exists, it creates a sequentially numbered sheet (e.g., `wbs-1`, `wbs-2`).
    2.  Creates a `holidays-tw` worksheet if one does not already exist.
    3.  Applies headers, formatting, data validation rules, and automated formulas to the newly created `wbs` sheet.

### 2.2. `onEdit(e)`

*   **Purpose**: An automatic trigger that runs whenever a user edits any cell in the spreadsheet.
*   **Execution**:
    *   It specifically monitors edits in the `TaskStatus` column (Column G) of any sheet whose name begins with "wbs".
    *   Based on the new status, it automatically updates the `DoneDate` for that task to enforce data consistency rules.

### 2.3. `resetWBSContent()`

*   **Purpose**: This function, triggered from the custom menu, clears the active WBS sheet of all task-related data, allowing it to be reused as a template.
*   **Execution**:
    1.  Confirms the active sheet is a `wbs` sheet.
    2.  Clears the cell **values** in the range from the second column of the second row (cell B2) to the last active cell.
    3.  **Crucially, this action preserves**:
        *   The header row (Row 1).
        *   All content in the first column (`Object`, Column A).
        *   All cell formulas (e.g., in `DueDate` and `TaskDescription-2`).
        *   All data validation rules and cell formatting.

---

## 3. Sheet: `wbs`

### 3.1. Column Structure & Automation

| Column | Name | Type | Description & Implemented Logic |
|---|---|---|---|
| A | **Object** | String | **Manual Input:** The name of the project or target. **(Preserved by `resetWBSContent`)** |
| B | **TaskTitle** | String | **Manual Input:** The title of the specific task. |
| C | **TaskDescription-1** | String | **Manual Input:** Detailed description of the task. |
| D | **StartDate** | Date | **Manual Input:** The planned start date of the task. |
| E | **WorkDays** | Integer | **Manual Input:** The number of working days required. |
| F | **Resource** | String | **Manual Input:** The person assigned to the task. |
| G | **TaskStatus** | Enum | **Data Validation:** Users must select from a dropdown list containing: `NotStarted`, `InProgress`, `Done`, `Blocked`. Invalid entries are rejected. |
| H | **DoneDate** | Date | **Automated by `onEdit` trigger:** <br> - If `TaskStatus` is set to `Done`, this cell is automatically populated with the current date. <br> - If `TaskStatus` is changed to any other value, this cell's content is cleared. |
| I | **DueDate** | Date | **Automated by Sheet Formula:** The cell contains the formula: `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, 'holidays-tw'!A$2:A), "")`. It calculates the due date based on `StartDate`, `WorkDays`, and the holiday list. |
| J | **TaskDescription-2**| String | **Automated by Sheet Formula:** The cell contains the formula: `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")`. It creates a summary string from the `Resource` and `TaskDescription-1` fields. |

### 3.2. Formatting

*   **Header**: The first row is frozen, with a background color of `#cfe2f3` and bold text.
*   **Data Types**: Columns `D`, `H`, and `I` are formatted as `yyyy-mm-dd`. Column `E` is formatted as a number.

---

## 4. Sheet: `holidays-tw`

**Purpose**: Defines a list of non-working dates (holidays) to be used in `DueDate` calculations.

| Column | Name | Type | Description |
|---|---|---|---|
| A | **date** | Date | The date of the holiday. |
| B | **holiday-name** | String | The name of the holiday. |

---
【END】
