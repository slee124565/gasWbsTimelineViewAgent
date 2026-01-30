/**
 * 調整後的初始化腳本：檢查 wbs 是否存在，並建立對應名稱的工作表
 */
function initializeWBSSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseName = 'wbs';
  let targetSheetName = '';

  // --- 1. 決定目標工作表名稱 ---
  const existingSheet = ss.getSheetByName(baseName);
  
  if (!existingSheet) {
    // 案例 A：若 wbs 不存在，則使用 wbs
    targetSheetName = baseName;
  } else {
    // 案例 B：若 wbs 已存在，尋找下一個序號 (wbs-1, wbs-2, ...)
    let counter = 1;
    while (ss.getSheetByName(`${baseName}-${counter}`)) {
      counter++;
    }
    targetSheetName = `${baseName}-${counter}`;
  }

  // 建立新工作表
  const newWbsSheet = ss.insertSheet(targetSheetName);

  // --- 2. 設定 wbs 欄位結構 (依據規格書 v0.0.2) ---
  const wbsHeaders = [
    'Object', 'TaskTitle', 'TaskDescription-1', 'StartDate', 
    'WorkDays', 'Resource', 'TaskStatus', 'DoneDate', 
    'DueDate', 'TaskDescription-2'
  ];
  
  newWbsSheet.getRange(1, 1, 1, wbsHeaders.length).setValues([wbsHeaders])
    .setBackground('#cfe2f3')
    .setFontWeight('bold');

  // 設定格式：日期與數字
  newWbsSheet.getRange('D2:D').setNumberFormat('yyyy-mm-dd'); // StartDate
  newWbsSheet.getRange('H2:H').setNumberFormat('yyyy-mm-dd'); // DoneDate
  newWbsSheet.getRange('I2:I').setNumberFormat('yyyy-mm-dd'); // DueDate
  newWbsSheet.getRange('E2:E').setNumberFormat('0');          // WorkDays
  newWbsSheet.setFrozenRows(1);

  // --- 3. 新增：設定自動化公式 ---
  // A. DueDate (I欄) 公式
  const dueDateFormula = `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, 'holidays-tw'!A$2:A), "")`;
  newWbsSheet.getRange('I2').setFormula(dueDateFormula);

  // B. TaskDescription-2 (J欄) 公式
  const taskDesc2Formula = `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")`;
  newWbsSheet.getRange('J2').setFormula(taskDesc2Formula);

  // C. 將公式應用到後續儲存格 (透過複製貼上)
  newWbsSheet.getRange('I2:J2').autoFill(newWbsSheet.getRange('I2:J' + newWbsSheet.getMaxRows()), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);


  // --- 4. 新增：設定資料驗證 ---
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['NotStarted', 'InProgress', 'Done', 'Blocked'])
    .setAllowInvalid(false) // 不允許無效輸入
    .setHelpText('請選擇一個有效的任務狀態。')
    .build();
  newWbsSheet.getRange('G2:G').setDataValidation(statusRule);


  // --- 5. 檢查並初始化 holidays-tw (此表通常為全域參考，若不存在才建立) ---
  const holidayName = 'holidays-tw';
  let holidaySheet = ss.getSheetByName(holidayName);
  if (!holidaySheet) {
    holidaySheet = ss.insertSheet(holidayName);
    const holidayHeaders = ['date', 'holiday-name'];
    holidaySheet.getRange(1, 1, 1, holidayHeaders.length).setValues([holidayHeaders])
      .setBackground('#d9ead3')
      .setFontWeight('bold');
    holidaySheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
    holidaySheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(`已成功建立工作表：${targetSheetName}，並已設定自動化規則。`);
}

/**
 * onEdit Trigger: 當使用者編輯儲存格時自動觸發
 * 處理 TaskStatus 與 DoneDate 的連動
 * 處理 Object 變更時的顏色自動更新
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 僅在 'wbs' 或 'wbs-x' 工作表上觸發
  if (!sheetName.startsWith('wbs')) {
    return;
  }

  const editedCol = range.getColumn();
  const editedRow = range.getRow();
  
  // 如果編輯的是 TaskStatus (G欄, 第7欄) 且不是標頭列
  if (editedCol === 7 && editedRow > 1) {
    const status = range.getValue();
    const doneDateCell = sheet.getRange(editedRow, 8); // H欄, 第8欄

    if (status === 'Done') {
      // 當狀態改為 "Done"，自動填入今天日期
      doneDateCell.setValue(new Date());
    } else {
      // 當狀態不是 "Done" 時，清空完成日期
      doneDateCell.clearContent();
    }
  }

  // 如果編輯的是 Object (A欄, 第1欄) 且不是標頭列，就觸發顏色更新
  if (editedCol === 1 && editedRow > 1) {
    applyObjectColorCoding();
  }

  // 如果編輯的是 Resource (F欄, 第6欄) 且不是標頭列，就觸發顏色更新
  if (editedCol === 6 && editedRow > 1) {
    applyResourceColorCoding();
  }
}

/**
 * 清空 WBS 工作表的內容（任務資料），並重設所有自動化公式。
 * 保留第一列（標頭）與第一欄（Object 欄位）。
 * 主要用於模板化重用 WBS 結構或修復被意外刪除/修改的公式。
 */
function resetWBSContentFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // 1. 確認是在 wbs 工作表上操作
  if (!sheetName.startsWith('wbs')) {
    SpreadsheetApp.getUi().alert('此功能只能在 "wbs" 或 "wbs-x" 工作表上執行。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // 2. 如果只有標頭列，則無需操作
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('工作表沒有需要清除的資料。');
    return;
  }

  // 3. 清除所有任務內容（從 B2 到最後一列的最後一欄），保留第一欄 (Object) 和所有格式/驗證規則。
  const contentRangeToClear = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
  contentRangeToClear.clearContent();

  // 4. 重設自動化公式 (DueDate 在 I欄, TaskDescription-2 在 J欄)
  const dueDateFormula = `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, 'holidays-tw'!A$2:A), "")`;
  const taskDesc2Formula = `=IF(C2<>"", IF(F2<>"", "["&F2&"]-"&C2, "[未指派]-"&C2), "")`;

  sheet.getRange('I2:J' + sheet.getMaxRows()).clearContent(); // Clear to max rows

  sheet.getRange('I2').setFormula(dueDateFormula);
  sheet.getRange('J2').setFormula(taskDesc2Formula);

  const sourceRange = sheet.getRange('I2:J2');
  const destinationRange = sheet.getRange('I2:J' + sheet.getMaxRows()); // 擴展到最大行數
  sourceRange.autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  // 5. 清除所有背景顏色（除了標頭）
  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);


  SpreadsheetApp.getUi().alert(`已成功清除 "${sheetName}" 的任務內容並重設欄位公式。`);
}

/**
 * 根據 Object 欄位的值，為 WBS 表格的每一列應用不同的背景顏色。
 */
function applyObjectColorCoding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith('wbs')) {
    SpreadsheetApp.getUi().alert('此功能只能在 "wbs" 或 "wbs-x" 工作表上執行。');
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const lastCol = dataRange.getLastColumn();

  // 定義一個顏色列表，用於循環
  const colors = [
    '#fff2cc', // Light Yellow
    '#d9ead3', // Light Green
    '#f4cccc', // Light Red
    '#d0e0e3', // Light Blue
    '#ead1dc', // Light Purple
    '#c9daf8', // Light Royal Blue
    '#d9d2e9', // Light Violet
    '#ace6e6', // Light Cyan
    '#ffe5b4', // Light Orange
    '#cccccc'  // Light Gray
  ];
  let colorIndex = 0;
  const objectColorMap = {};

  // 從第二列開始遍歷（跳過標頭）
  for (let i = 1; i < values.length; i++) {
    const objectName = values[i][0]; // 第 A 欄是 Object

    if (objectName) {
      // 如果這個 Object 還沒有分配顏色，就給它一個新的
      if (!objectColorMap[objectName]) {
        objectColorMap[objectName] = colors[colorIndex % colors.length];
        colorIndex++;
      }
      
      // 應用顏色到整列
      const rowRange = sheet.getRange(i + 1, 1, 1, lastCol);
      rowRange.setBackground(objectColorMap[objectName]);
    } else {
      // 如果 Object 為空，則清除背景顏色
      const rowRange = sheet.getRange(i + 1, 1, 1, lastCol);
      rowRange.setBackground(null);
    }
  }
  SpreadsheetApp.getUi().alert('已成功套用 Object 顏色標記。');
}

/**
 * 根據 Resource 欄位的值，為 WBS 表格的每一列應用不同的背景顏色。
 */
function applyResourceColorCoding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith('wbs')) {
    SpreadsheetApp.getUi().alert('此功能只能在 "wbs" 或 "wbs-x" 工作表上執行。');
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const lastCol = dataRange.getLastColumn();

  // 定義一個顏色列表，用於循環 (可以使用與 Object 不同的顏色，或從同一調色盤中選取)
  const colors = [
    '#b6d7a8', // Light Green (different shade)
    '#a2c4c9', // Light Blue (different shade)
    '#ea9999', // Light Red (different shade)
    '#f9cb9c', // Light Orange (different shade)
    '#b4a7d6', // Light Purple (different shade)
    '#a4c2f4', // Light Blue
    '#cfc2b6', // Light Brown
    '#8ee4af', // Mint Green
    '#fada5e', // Saffron
    '#f5b7b1'  // Light Coral
  ];
  let colorIndex = 0;
  const resourceColorMap = {};

  // 從第二列開始遍歷（跳過標頭）
  for (let i = 1; i < values.length; i++) {
    const resourceName = values[i][5]; // 第 F 欄是 Resource (索引為 5)

    if (resourceName) {
      // 如果這個 Resource 還沒有分配顏色，就給它一個新的
      if (!resourceColorMap[resourceName]) {
        resourceColorMap[resourceName] = colors[colorIndex % colors.length];
        colorIndex++;
      }
      
      // 應用顏色到整列
      const rowRange = sheet.getRange(i + 1, 1, 1, lastCol);
      rowRange.setBackground(resourceColorMap[resourceName]);
    } else {
      // 如果 Resource 為空，則清除背景顏色
      const rowRange = sheet.getRange(i + 1, 1, 1, lastCol);
      rowRange.setBackground(null);
    }
  }
  SpreadsheetApp.getUi().alert('已成功套用 Resource 顏色標記。');
}


/**
 * 新增自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 WBS 自動化工具')
    .addItem('1. 建立新 WBS 工作表', 'initializeWBSSystem')
    .addSeparator()
    .addItem('2. 重設任務內容與公式 (保留首欄)', 'resetWBSContentFormulas')
    .addItem('3. 套用 Object 顏色標記', 'applyObjectColorCoding')
    .addItem('4. 套用 Resource 顏色標記', 'applyResourceColorCoding') // 新增 Resource 顏色標記選項
    .addToUi();
}