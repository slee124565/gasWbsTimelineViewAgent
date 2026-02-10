### Project Overview

This is a Google Apps Script project designed to automate Work Breakdown Structure (WBS) management within Google Sheets. Its core functionality shifts the workflow from using a pre-defined template to a script-driven approach, where the WBS environment is dynamically initialized within any Google Sheet.

The script's primary feature is its ability to create new `wbs` sheets on-demand via a custom menu. If a sheet named `wbs` already exists, it intelligently creates numbered versions (`wbs-1`, `wbs-2`, etc.) to prevent conflicts, allowing for multiple WBS datasets within a single spreadsheet.

Core features include:
-   **Dynamic Sheet Initialization**: The `initializeWBSSystem` function creates a fully structured WBS sheet with predefined headers, cell formatting, data validation rules, and frozen rows. It also creates a `TW_HOLIDAYS` sheet if one doesn't exist.
-   **Automated Calculations**: It applies per-cell formulas to each row for `DueDate` and `TaskDescription-2`, ensuring that calculations are automatically handled as new tasks are added.
-   **Status-driven Actions**: Automatically populates the `DoneDate` with the current date when a task's `TaskStatus` is set to "Done" via an `onEdit` trigger.
-   **Data Validation**: Implements a dropdown list for the `TaskStatus` column to ensure data consistency, restricting inputs to 'NotStarted', 'InProgress', 'Done', or 'Blocked'.
-   **Visual Grouping**: Automatically applies unique background colors to cells in the "Object" (Column A) and "Resource" (Column F) columns based on their content, allowing for easy visual identification of related tasks or resources.
-   **Custom UI Menu**: Adds a menu named "🚀 WBS 自動化工具" to the Google Sheet UI, providing easy access to all major functions like sheet creation, content reset, and manual color-coding.

The main logic is contained in `code.js` and configured via `appsscript.json`. The project is intended to be managed and deployed using `clasp`, the command-line interface for Google Apps Script.

### Building and Running

This is a Google Apps Script project, which means there is no traditional "build" process. The code is deployed directly to a Google Sheet.

**Deployment Steps:**

1.  **Use `clasp` (Recommended)**:
    *   Ensure you have `clasp` installed (`npm install -g @google/clasp`).
    *   Log in to your Google account using `clasp login`.
    *   Create a new, blank Google Sheet.
    *   In your local project directory, run `clasp create --type sheets --title "Your Sheet Name"` to link the project to the new sheet.
    *   Run `clasp push` to upload the `code.js` and `appsscript.json` files to the script bound to your Google Sheet.
2.  **Manual Setup (Alternative)**:
    *   Create a new, blank Google Sheet.
    *   Open the Apps Script editor via **Extensions > Apps Script**.
    *   Copy the entire content of the local `code.js` file and paste it into the editor, replacing any existing code.
    *   Save the script project.

**Execution:**

1.  After deployment, reload the Google Sheet. The custom menu "🚀 WBS 自動化工具" should appear.
2.  **The crucial first step is to select "1. 建立新 WBS 工作表" from the menu.** This will execute the `initializeWBSSystem` function, which creates the `wbs` and `TW_HOLIDAYS` sheets and sets up all required automation.
3.  The `onEdit` trigger will automatically handle date and color updates as you modify data in any sheet whose name starts with `wbs`.

### Development Conventions

-   **Triggers**: The script relies on simple triggers (`onOpen`, `onEdit`) for its core automation. The `onEdit` logic is specifically designed to activate on any sheet with a name prefixed by `wbs`.
-   **Sheet Naming**: The script follows a dynamic naming convention, creating `wbs` for the first instance and `wbs-x` (where x is a number) for subsequent sheets. It also depends on a global `TW_HOLIDAYS` sheet for date calculations.
-   **Column Structure**: The automation logic is tightly coupled to specific column numbers (e.g., Column G for `TaskStatus`, Column A for `Object`). This structure is defined in the `initializeWBSSystem` function.
-   **User Interaction**: All user-facing alerts and menu items are in Traditional Chinese (繁體中文).
-   **Formulas**: The script has evolved from using a single `ARRAYFORMULA` to setting per-cell formulas via `setFormula()` and propagating them down the sheet using `autoFill()`. This approach is more robust against accidental formula deletion in a single cell.