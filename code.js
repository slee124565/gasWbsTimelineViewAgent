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
}


/**
 * 新增自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 WBS 自動化工具')
    .addItem('1. 建立新 WBS 工作表', 'initializeWBSSystem')
    .addToUi();
}