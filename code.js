/**
 * onEdit Trigger: When a user edits a cell, this function is automatically triggered.
 * It handles the automatic update of the 'DoneDate' based on 'TaskStatus' and
 * triggers the color-coding functions when 'Object' or 'Resource' columns are changed.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Only trigger in the 'wbs' sheet
  if (sheetName !== 'wbs') {
    return;
  }

  const editedCol = range.getColumn();
  const editedRow = range.getRow();
  
  // If 'TaskStatus' (Column G) is edited (and it's not the header row)
  if (editedCol === 7 && editedRow > 1) {
    const status = range.getValue();
    const doneDateCell = sheet.getRange(editedRow, 8); // Column H

    if (status === 'Done') {
      doneDateCell.setValue(new Date());
    } else {
      doneDateCell.clearContent();
    }
  }

  // If 'Object' (Column A) is edited, trigger color update
  if (editedCol === 1 && editedRow > 1) {
    applyObjectColorCoding();
  }

  // If 'Resource' (Column F) is edited, trigger color update
  if (editedCol === 6 && editedRow > 1) {
    applyResourceColorCoding();
  }
}

/**
 * Clears the content of the WBS sheet and resets all automation formulas.
 * This version preserves the content of the first column ('Object').
 * It's used for reusing the WBS template or repairing broken formulas.
 */
function resetWBSContentFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('wbs');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('The "wbs" sheet was not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // If there's no data to clear (only a header row), do nothing.
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('There is no data to clear.');
    return;
  }

  // Clear content from the second column to the last column, preserving the 'Object' column.
  const contentRangeToClear = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
  contentRangeToClear.clearContent();

  // Clear all background colors (except for the header)
  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);
  
  SpreadApp.getUi().alert('Task content has been cleared and formulas have been reset.');
}


/**
 * Applies a unique background color to cells in the 'Object' column (Column A)
 * based on their value for easy visual grouping.
 */
function applyObjectColorCoding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('wbs');
  if (!sheet) return;

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const colors = [
    '#fff2cc', '#d9ead3', '#f4cccc', '#d0e0e3', '#ead1dc',
    '#c9daf8', '#d9d2e9', '#ace6e6', '#ffe5b4', '#cccccc'
  ];
  let colorIndex = 0;
  const objectColorMap = {};

  // Start from row 2 (index 1) to skip the header
  for (let i = 1; i < values.length; i++) {
    const objectName = values[i][0]; // Column A

    if (objectName) {
      if (!objectColorMap[objectName]) {
        objectColorMap[objectName] = colors[colorIndex % colors.length];
        colorIndex++;
      }
      sheet.getRange(i + 1, 1).setBackground(objectColorMap[objectName]);
    } else {
      sheet.getRange(i + 1, 1).setBackground(null);
    }
  }
}

/**
 * Applies a unique background color to cells in the 'Resource' column (Column F)
 * based on their value for easy visual grouping.
 */
function applyResourceColorCoding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('wbs');
  if (!sheet) return;

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const colors = [
    '#b6d7a8', '#a2c4c9', '#ea9999', '#f9cb9c', '#b4a7d6',
    '#a4c2f4', '#cfc2b6', '#8ee4af', '#fada5e', '#f5b7b1'
  ];
  let colorIndex = 0;
  const resourceColorMap = {};

  // Start from row 2 (index 1) to skip the header
  for (let i = 1; i < values.length; i++) {
    const resourceName = values[i][5]; // Column F

    if (resourceName) {
      if (!resourceColorMap[resourceName]) {
        resourceColorMap[resourceName] = colors[colorIndex % colors.length];
        colorIndex++;
      }
      sheet.getRange(i + 1, 6).setBackground(resourceColorMap[resourceName]);
    } else {
      sheet.getRange(i + 1, 6).setBackground(null);
    }
  }
}

/**
 * Creates the custom menu in the spreadsheet UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 WBS 自動化工具')
    .addItem('重設任務內容與公式 (保留首欄)', 'resetWBSContentFormulas')
    .addSeparator()
    .addItem('套用 Object 顏色標記', 'applyObjectColorCoding')
    .addItem('套用 Resource 顏色標記', 'applyResourceColorCoding')
    .addToUi();
}