# Manual Test Guide for Script-Driven WBS

This document provides the steps to manually test the functionality of the `gasWbsTimelineViewAgent` script, based on its script-driven initialization design.

## 1. Setup and Initialization

**Objective:** Verify that the script correctly creates and initializes the `wbs` and `TW_HOLIDAYS` sheets with all automated rules from a blank Google Sheet.

**Steps:**

1.  Create a new, blank Google Sheet.
2.  Open the script editor by going to **Extensions > Apps Script**.
3.  Paste the code from `code.js` into the editor, replacing any existing content.
4.  Save the script project.
5.  Reload the Google Sheet. You should see a new menu item: **🚀 WBS 自動化工具**.
6.  Click on the **🚀 WBS 自動化工具** menu.
7.  Select **1. 建立新 WBS 工作表**.
8.  A popup saying "已成功建立工作表：wbs，並已設定自動化規則。" should appear. Click **OK**.

**Verification:**

*   Check that two new sheets named `wbs` and `TW_HOLIDAYS` have been created.
*   In the `wbs` sheet, confirm the header row (Row 1) is formatted with a blue background and bold text.
*   In the `wbs` sheet, confirm that Row 1 is frozen.
*   Click on any cell in the `TaskStatus` column (Column G, from G2 downwards). A dropdown arrow should appear. Clicking it should show the options: `NotStarted`, `InProgress`, `Done`, `Blocked`.
*   Click on cell `I2` (DueDate) and verify that its formula is `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, TW_HOLIDAYS!A$2:A), "")`.
*   Click on cell `J2` (TaskDescription-2) and verify that its formula is `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")`.

---

## 2. Test Cases (after Initialization)

### Test Case 2.1: Automated `DueDate` Calculation

**Objective:** Verify that the `DueDate` is calculated automatically based on `StartDate`, `WorkDays`, and holidays.

**Steps:**

1.  Go to the `TW_HOLIDAYS` sheet. In cell `A2`, enter a date, for example, `2026-02-03`.
2.  Go to the `wbs` sheet.
3.  In cell `D2` (`StartDate`), enter the date `2026-02-01`.
4.  In cell `E2` (`WorkDays`), enter the number `5`.
5.  **Verification:**
    *   Check cell `I2` (`DueDate`). With a start date of `2026-02-01`, a holiday on `2026-02-03`, and 5 workdays, the expected due date should be `2026-02-09`.
    *   Clear the value in cell `D2`. The `DueDate` in `I2` should become blank.

### Test Case 2.2: Automated `TaskDescription-2` Generation

**Objective:** Verify that `TaskDescription-2` is generated correctly based on `Resource` and `TaskDescription-1`.

**Steps:**

1.  Go to the `wbs` sheet.
2.  In cell `C2` (`TaskDescription-1`), type `Develop feature X`.
3.  **Verification:**
    *   Check cell `J2`. It should automatically display `[未指派]-Develop feature X`.
4.  Now, in cell `F2` (`Resource`), type `Alice`.
5.  **Verification:**
    *   Check cell `J2` again. It should automatically update to `[Alice]-Develop feature X`.

### Test Case 2.3: `onEdit` Trigger for `DoneDate`

**Objective:** Verify that the `DoneDate` is automatically populated when `TaskStatus` is set to "Done" and cleared otherwise.

**Steps:**

1.  Go to the `wbs` sheet.
2.  In cell `G2` (`TaskStatus`), select **Done** from the dropdown list.
3.  **Verification:**
    *   Check cell `H2` (`DoneDate`). It should be automatically populated with today's date.
4.  Now, click on cell `G2` again.
5.  Change the status to **InProgress**.
6.  **Verification:**
    *   Check cell `H2` again. The date should now be cleared.

### Test Case 2.4: Object-Based Cell Color-Coding

**Objective:** Verify that only the Object cell (Column A) is colored based on its value, both via the menu and on edit.

**Steps (Part A: Manual Trigger):**

1.  In the `wbs` sheet, enter the following data:
    *   `A2`: `Project Alpha`
    *   `A3`: `Project Alpha`
    *   `A4`: `Project Beta`
2.  Click **🚀 WBS 自動化工具 > 3. 套用 Object 顏色標記**.
3.  **Verification:**
    *   Cells `A2` and `A3` should have the same background color.
    *   Cell `A4` should have a different background color.
    *   Other cells (e.g., in columns B, C, D) should NOT have any background color.

**Steps (Part B: Automatic `onEdit` Trigger):**

1.  In cell `A4`, change the value from `Project Beta` to `Project Alpha`.
2.  **Verification:**
    *   The background color of cell `A4` should automatically change to match the color of cells `A2` and `A3`.

### Test Case 2.5: Resource-Based Cell Color-Coding

**Objective:** Verify that only the Resource cell (Column F) is colored based on its value, both via the menu and on edit.

**Steps (Part A: Manual Trigger):**

1.  In the `wbs` sheet, enter the following data:
    *   `F2`: `Alice`
    *   `F3`: `Bob`
    *   `F4`: `Alice`
2.  Click **🚀 WBS 自動化工具 > 4. 套用 Resource 顏色標記**.
3.  **Verification:**
    *   Cells `F2` and `F4` should have the same background color.
    *   Cell `F3` should have a different background color.

**Steps (Part B: Automatic `onEdit` Trigger):**

1.  In cell `F3`, change the value from `Bob` to `Alice`.
2.  **Verification:**
    *   The background color of cell `F3` should automatically change to match the color of cells `F2` and `F4`.

### Test Case 2.6: Reset WBS Content and Formulas

**Objective:** Verify that the "Reset" function clears all task data (including `Object`) from row 2 downwards, resets formulas, and removes background colors.

**Steps:**

1.  Populate data across multiple rows (e.g., `A2:F4`).
2.  Use the color-coding functions to apply background colors to columns A and F.
3.  Click **🚀 WBS 自動化工具 > 2. 重設任務內容與公式**.
4.  A confirmation alert should appear. Click **OK**.
5.  **Verification:**
    *   Verify that the range `A2:J4` (or as far as you entered data) is completely empty. **Crucially, verify that cell `A2` is now blank.**
    *   Verify that the background colors in the data rows have been removed.
    .
    *   Verify that the formulas in `I2` and `J2` have been correctly reset.
    *   Verify that the header row (row 1) remains unchanged.

---
**End of Test Guide**
