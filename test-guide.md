# Manual Test Guide for WBS Automation Script

This document provides the steps to manually test the functionality of the `gasWbsTimelineViewAgent` script after the v0.0.2 updates.

## 1. Prerequisites

1.  Open a Google Sheet.
2.  Open the script editor by going to **Extensions > Apps Script**.
3.  Paste the updated code from `code.js` into the editor.
4.  Save the script project.
5.  Reload the Google Sheet. You should see a new menu item: **🚀 WBS 自動化工具**.

---

## 2. Test Cases

### Test Case 1: Initial Sheet Creation and Setup

**Objective:** Verify that the script correctly creates and initializes the `wbs` and `holidays-tw` sheets with all automated rules.

**Steps:**

1.  Click on the **🚀 WBS 自動化工具** menu.
2.  Select **1. 建立新 WBS 工作表**.
3.  A popup saying "已成功建立工作表：wbs，並已設定自動化規則。" should appear. Click **OK**.
4.  **Verification:**
    *   Check that two new sheets named `wbs` and `holidays-tw` have been created.
    *   In the `wbs` sheet, confirm the header row is blue and bold.
    *   Click on any cell in the `TaskStatus` column (Column G, from G2 downwards). A dropdown arrow should appear. Clicking it should show the options: `NotStarted`, `InProgress`, `Done`, `Blocked`.
    *   Click on cell `I2` (DueDate) and verify that its formula is `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, 'holidays-tw'!A$2:A), "")`.
    *   Click on cell `J2` (TaskDescription-2) and verify that its formula is `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")`.

### Test Case 2: Automated `DueDate` Calculation

**Objective:** Verify that the `DueDate` is calculated automatically based on `StartDate`, `WorkDays`, and holidays.

**Steps:**

1.  Go to the `holidays-tw` sheet. In cell `A2`, enter a date that falls within your test's timeframe, for example, `2026-02-03`.
2.  Go to the `wbs` sheet.
3.  In cell `D2` (`StartDate`), enter the date `2026-02-01`.
4.  In cell `E2` (`WorkDays`), enter the number `5`.
5.  **Verification:**
    *   Check cell `I2` (`DueDate`). The date should be automatically calculated. With a start date of `2026-02-01` (Sunday), `2026-02-03` as a holiday, and 5 workdays, the expected due date should be `2026-02-09`. (Workdays: Feb 2, 4, 5, 6, 9).
    *   Clear the value in cell `D2`. The `DueDate` in `I2` should become blank.

### Test Case 3: Automated `TaskDescription-2` Generation

**Objective:** Verify that `TaskDescription-2` is generated correctly based on `Resource` and `TaskDescription-1`.

**Steps:**

1.  Go to the `wbs` sheet.
2.  In cell `C2` (`TaskDescription-1`), type `Develop feature X`.
3.  **Verification:**
    *   Check cell `J2` (`TaskDescription-2`). It should automatically display `[未指派]-Develop feature X`.
4.  Now, in cell `F2` (`Resource`), type `Alice`.
5.  **Verification:**
    *   Check cell `J2` again. It should automatically update to `[Alice]-Develop feature X`.

### Test Case 4: `onEdit` Trigger for `DoneDate`

**Objective:** Verify that the `DoneDate` is automatically populated when `TaskStatus` is set to "Done" and cleared otherwise.

**Steps:**

1.  Go to the `wbs` sheet.
2.  In row 3, enter some sample data for a new task.
3.  Click on cell `G3` (`TaskStatus`).
4.  Select **Done** from the dropdown list.
5.  **Verification:**
    *   Check cell `H3` (`DoneDate`). It should be automatically populated with today's date.
6.  Now, click on cell `G3` again.
7.  Change the status to **InProgress**.
8.  **Verification:**
    *   Check cell `H3` again. The date should now be cleared and the cell should be blank.

### Test Case 5: Object-Based Row Color-Coding

**Objective:** Verify that rows are automatically colored based on the value in the `Object` column, both via the menu and on edit.

**Steps (Part A: Manual Trigger):**

1.  Go to the `wbs` sheet.
2.  Populate some data in rows 2, 3, 4, and 5.
    *   In cell `A2`, enter `Project Alpha`.
    *   In cell `A3`, enter `Project Alpha`.
    *   In cell `A4`, enter `Project Beta`.
    *   In cell `A5`, enter `Project Gamma`.
3.  Click on the **🚀 WBS 自動化工具** menu.
4.  Select **3. 套用 Object 顏色標記**.
5.  A confirmation alert should appear. Click **OK**.
6.  **Verification:**
    *   Rows 2 and 3 should have the same background color.
    *   Row 4 should have a different background color from rows 2/3.
    *   Row 5 should have a different background color from rows 2/3 and 4.

**Steps (Part B: Automatic `onEdit` Trigger):**

1.  Continuing from the previous state, go to cell `A5`.
2.  Change the value from `Project Gamma` to `Project Alpha`.
3.  **Verification:**
    *   The background color of row 5 should automatically change to match the color of rows 2 and 3. No need to run the menu item again.
4.  Now, go to cell `A4`.
5.  Change the value from `Project Beta` to `Project Zeta`.
6.  **Verification:**
    *   The background color of row 4 should change to a new, different color.

### Test Case 6: Resource-Based Row Color-Coding

**Objective:** Verify that rows are automatically colored based on the value in the `Resource` column, both via the menu and on edit.

**Steps (Part A: Manual Trigger):**

1.  Go to the `wbs` sheet.
2.  Populate `Resource` data in rows 2, 3, and 4.
    *   In cell `F2`, enter `Alice`.
    *   In cell `F3`, enter `Bob`.
    *   In cell `F4`, enter `Alice`.
3.  Click on the **🚀 WBS 自動化工具** menu.
4.  Select **4. 套用 Resource 顏色標記**.
5.  A confirmation alert should appear. Click **OK**.
6.  **Verification:**
    *   Rows 2 and 4 should have the same background color.
    *   Row 3 should have a different background color.
    *   Note: These colors will overwrite the `Object` colors from the previous test.

**Steps (Part B: Automatic `onEdit` Trigger):**

1.  Continuing from the previous state, go to cell `F3`.
2.  Change the value from `Bob` to `Alice`.
3.  **Verification:**
    *   The background color of row 3 should automatically change to match the color of rows 2 and 4.

---
**End of Test Guide**
