# Google Apps Script Specification: WBS Automation v0.0.5

This document details the implemented functionality and automation rules for the `gasWbsTimelineViewAgent` script. This version introduces automated row color-coding for visual grouping of tasks.

---

## 1. System Overview

The script provides a set of tools to automate the creation and management of a Work Breakdown Structure (WBS) within a Google Sheet. It is designed to reduce manual data entry and ensure data consistency.

The primary user entry point is a custom menu labeled **"🚀 WBS 自動化工具"**, which is automatically added to the spreadsheet's UI when the document is opened (`onOpen` trigger). This menu provides the following options:
1.  **建立新 WBS 工作表 (Create New WBS Sheet)**: Sets up a new, fully configured WBS sheet.
2.  **重設任務內容與公式 (保留首欄) (Reset Task Content & Formulas (Keep First Column))**: Resets the current WBS sheet's content and repairs its automated formulas.
3.  **套用 Object 顏色標記 (Apply Object Color-Coding)**: Applies a unique background color to all rows associated with the same `Object` value for easy visual grouping.

---

## 2. Core Script Components & Triggers

### 2.1. `initializeWBSSystem()`

*   **Purpose**: Sets up the entire WBS environment.
*   **Execution**: Creates and configures the `wbs` and `holidays-tw` sheets with all necessary headers, formats, and formulas.

### 2.2. `onEdit(e)`

*   **Purpose**: An automatic trigger that runs whenever a user edits a cell.
*   **Execution**:
    *   Monitors the `TaskStatus` column (Column G) to automatically manage the `DoneDate` field.
    *   **New in v0.0.5**: It can be extended to automatically trigger the color-coding logic when a value in the `Object` column (Column A) is changed.

### 2.3. `resetWBSContentFormulas()`

*   **Purpose**: A consolidated function to clear task content and reset automated formulas, preparing the sheet for reuse.

### 2.4. `applyObjectColorCoding()`

*   **Purpose**: This function, triggered from the custom menu, applies consistent background colors to rows based on the value in the `Object` column.
*   **Execution**:
    1.  Scans the entire `Object` column (Column A) to identify all unique text values.
    2.  Assigns a unique color to each unique `Object` value from a predefined, non-repeating palette.
    3.  Iterates through every data row (from row 2 downwards).
    4.  For each row, it sets the background color of the entire row to match the color assigned to that row's `Object` value.
    5.  The header row's color remains unchanged.

---

## 3. Sheet: `wbs`

### 3.1. Column Structure & Automation

The column structure remains the same as in v0.0.4.

| Column | Name | Type | Description & Implemented Logic |
|---|---|---|---|
| A | **Object** | String | **Manual Input:** The name of the project or target. **This column now drives the row color-coding.** |
| ... | ... | ... | ... |

### 3.2. Formatting

*   **Header**: The first row is frozen, with a background color of `#cfe2f3` and bold text.
*   **Data Types**: Date and number columns are formatted as specified.
*   **Row Color-Coding**:
    *   All cells within a single data row will have a consistent background color.
    *   The color is determined by the value in that row's `Object` cell (Column A).
    *   All rows with the same `Object` value will share the same background color.
    *   The script uses a predefined palette of distinct colors to ensure that different `Object` values have different colors.

---

## 4. Sheet: `holidays-tw`

No changes from the previous version. The purpose remains to define non-working dates for `DueDate` calculations.

---
【END】
