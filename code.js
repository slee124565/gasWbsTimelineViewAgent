const SCRIPT_VERSION = "v1.2.0";

/**
 * 調整後的初始化腳本：檢查 wbs 是否存在，並建立對應名稱的工作表
 */
function initializeWBSSystem() {
  console.log('initializeWBSSystem function triggered');
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
  const dueDateFormula = `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, TW_HOLIDAYS!A$2:A), "")`;
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
  const holidayName = 'TW_HOLIDAYS';
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
  console.log('onEdit function triggered');
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
 * 此版本會清除所有欄位（包含第一欄 Object）的內容，僅保留標頭。
 * 主要用於模板化重用 WBS 結構或修復被意外刪除/修改的公式。
 */
function resetWBSContentFormulas() {
  console.log('resetWBSContentFormulas function triggered');
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

  // 3. 清除所有任務內容（從 A2 到最後一列的最後一欄），並保留所有格式/驗證規則。
  const contentRangeToClear = sheet.getRange(2, 1, lastRow - 1, lastCol);
  contentRangeToClear.clearContent();

  // 4. 重設自動化公式 (DueDate 在 I欄, TaskDescription-2 在 J欄)
  const dueDateFormula = `=IF(AND(D2<>"", E2<>""), WORKDAY(D2, E2, TW_HOLIDAYS!A$2:A), "")`;
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
  console.log('applyObjectColorCoding function triggered');
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
      
      // 應用顏色到 Object 儲存格
      const cellRange = sheet.getRange(i + 1, 1); // A 欄
      cellRange.setBackground(objectColorMap[objectName]);
    } else {
      // 如果 Object 為空，則清除背景顏色
      const cellRange = sheet.getRange(i + 1, 1); // A 欄
      cellRange.setBackground(null);
    }
  }

}

/**
 * 根據 Resource 欄位的值，為 WBS 表格的每一列應用不同的背景顏色。
 */
function applyResourceColorCoding() {
  console.log('applyResourceColorCoding function triggered');
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
      
      // 應用顏色到 Resource 儲存格
      const cellRange = sheet.getRange(i + 1, 6); // F 欄
      cellRange.setBackground(resourceColorMap[resourceName]);
    } else {
      // 如果 Resource 為空，則清除背景顏色
      const cellRange = sheet.getRange(i + 1, 6); // F 欄
      cellRange.setBackground(null);
    }
  }

}

/**
 * 顯示目前腳本的版本號
 */
function showVersion() {
  SpreadsheetApp.getUi().alert(`目前腳本版本：${SCRIPT_VERSION}`);
}


/**
 * 處理 HTTP GET 請求的單一進入點 (路由)。
 * 根據 URL 參數 `output` 決定回傳 JSON 或 CSV 格式的 WBS 資料，
 * 或根據 `page` 參數（或預設）渲染 Timeline Dashboard 網頁。
 * 同時處理 `wbs` 工作表不存在的錯誤情況。
 *
 * @param {Object} e - 事件物件，包含請求參數。
 * @returns {ContentService.TextOutput|HtmlService.HtmlOutput} - 根據請求類型回傳 JSON/CSV 或 HTML 內容。
 */
function doGet(e) {
  console.log('doGet function triggered');
  const outputFormat = e.parameter.output; // 例如 ?output=json 或 ?output=csv
  const page = e.parameter.page; // 例如 ?page=timeline_dashboard
  const sheetName = 'wbs'; // 固定為 'wbs' 工作表，依據需求規格

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // 錯誤處理：如果找不到指定的工作表 'wbs'
  if (!sheet) {
    if (outputFormat === 'json' || outputFormat === 'csv') {
      // API 錯誤回應
      const errorResponse = {
        status: 'error',
        message: `Sheet "${sheetName}" not found. Please initialize the WBS system.`
      };
      return ContentService.createTextOutput(JSON.stringify(errorResponse))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      // 網頁錯誤回應
      return HtmlService.createHtmlOutput('<h1>錯誤：WBS 工作表未找到</h1><p>請先執行 WBS 系統的初始化 (透過 Google Sheet 介面)。</p>');
    }
  }

  // API 路由
  if (outputFormat === 'json') {
    return getWbsDataAsJson(sheetName);
  } else if (outputFormat === 'csv') {
    return getWbsDataAsCsv(sheetName);
  }

  // 網頁路由
  if (page === 'timeline_dashboard') {
    return showTimelineDashboard(sheetName);
  } else {
    // 若無 output 參數或 page 參數不符，則預設顯示 Timeline Dashboard
    return showTimelineDashboard(sheetName);
  }
}

/**
 * 取得指定 WBS 工作表的資料，並轉換為 JSON 格式回傳。
 * (原 doGet 函數的核心邏輯)
 *
 * @param {string} sheetName - 要讀取的工作表名稱。
 * @returns {ContentService.TextOutput} - JSON 格式的回應。
 */
function getWbsDataAsJson(sheetName) {
  console.log(`getWbsDataAsJson triggered for sheet: ${sheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // 根據 doGet 的錯誤處理，此處 sheet 應已存在，但為確保函數獨立性，再次檢查。
  if (!sheet) {
    const errorResponse = {
      status: 'error',
      message: `Sheet "${sheetName}" not found during JSON data retrieval.`
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) { // 錯誤處理：如果工作表是空的或只有標頭
    const emptyResponse = {
      status: 'success',
      sheet: sheetName,
      data: []
    };
    return ContentService.createTextOutput(JSON.stringify(emptyResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values.shift(); // 將第一列作為標頭
  
  // 篩選出 Column A (索引為 0) 有文字資料的行
  const filteredValues = values.filter(row => {
    return row[0] !== undefined && String(row[0]).trim() !== '';
  });

  const jsonData = filteredValues.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      // 處理日期物件：Apps Script 會將試算表中的日期自動轉換為 Date 物件
      if (row[index] instanceof Date) {
        // 格式化為 YYYY-MM-DD 字串
        obj[header] = Utilities.formatDate(row[index], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[header] = row[index];
      }
    });
    return obj;
  });

  const successResponse = {
    status: 'success',
    sheet: sheetName,
    data: jsonData
  };

  return ContentService.createTextOutput(JSON.stringify(successResponse, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 佔位符函數：取得指定 WBS 工作表的資料，並轉換為 CSV 格式回傳。
 *
 * @param {string} sheetName - 要讀取的工作表名稱。
 * @returns {ContentService.TextOutput} - CSV 格式的回應。
 */
function getWbsDataAsCsv(sheetName) {
  console.log(`getWbsDataAsCsv triggered for sheet: ${sheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService.createTextOutput(`Error: Sheet "${sheetName}" not found during CSV data retrieval.`)
      .setMimeType(ContentService.MimeType.PLAIN_TEXT);
  }

  const data = sheet.getDataRange().getDisplayValues(); // 取得顯示值，包含公式結果
  const csv = data.map(row => 
    row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
  ).join('\n');

  // 修正：使用 downloadAsFile 參數來設定檔名，並設定正確的 MimeType
  const output = ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.PLAIN_TEXT)
    .downloadAsFile(`${sheetName}.csv`);
  
  return output;
}

/**
 * 佔位符函數：渲染 Timeline Dashboard 網頁。
 *
 * @param {string} sheetName - 要讀取的工作表名稱。
 * @returns {HtmlService.HtmlOutput} - HTML 網頁內容。
 */
function showTimelineDashboard(sheetName) {
  console.log(`showTimelineDashboard triggered for sheet: ${sheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // 此處假設 sheet 已在 doGet 中檢查過存在性，但仍進行防禦性檢查
  if (!sheet) {
    return HtmlService.createHtmlOutput('<h1>錯誤：WBS 工作表未找到</h1><p>請先執行 WBS 系統的初始化。</p>');
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) { // 如果工作表是空的或只有標頭
    const template = HtmlService.createTemplateFromFile('TimelineDashboard');
    template.lastUpdated = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
    template.overdueTasks = [];
    template.futureTasks = [];
    template.pastTasks = [];
    return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const headers = values.shift(); // 取得標頭，並移除第一列
  const allTasks = values.filter(row => row[0] !== undefined && String(row[0]).trim() !== ''); // 篩選 Object 欄位有內容的行

  // 將所有任務轉換為物件陣列，並格式化日期
  const tasksAsObjects = allTasks.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      const value = row[index];
      // 對日期欄位進行格式化
      if (value instanceof Date) {
        obj[header] = Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[header] = value;
      }
    });
    return obj;
  });

  // --- 日期基準計算 ---
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 將時間歸零以便比較日期
  const fourWeeksAgo = new Date(today.getTime() - 28 * 24 * 60 * 60 * 1000);
  const fourWeeksHence = new Date(today.getTime() + 28 * 24 * 60 * 60 * 1000);

  // --- 篩選資料 ---
  let overdueTasks = [];
  let futureTasks = [];
  let pastTasks = [];

  tasksAsObjects.forEach(task => {
    const dueDate = task.DueDate ? new Date(task.DueDate) : null;
    const taskStatus = task.TaskStatus;

    if (dueDate) {
      if (dueDate < today && taskStatus !== 'Done' && taskStatus !== 'Blocked') {
        overdueTasks.push(task);
      } else if (dueDate >= today && dueDate <= fourWeeksHence && taskStatus !== 'Done') {
        futureTasks.push(task);
      } else if (dueDate >= fourWeeksAgo && dueDate < today && taskStatus === 'Done') {
        pastTasks.push(task);
      }
    }
  });

  // --- 排序資料 ---
  // 調整 `overdueTasks` 排序以支援 Object 分組：先依 Object 排序，再依 DueDate 排序
  overdueTasks.sort((a, b) => {
    if (a.Object < b.Object) return -1;
    if (a.Object > b.Object) return 1;
    return new Date(a.DueDate).getTime() - new Date(b.DueDate).getTime(); // 次要排序條件：DueDate (升序)
  });
  
  // 調整 `futureTasks` 排序以支援 Object 分組：先依 Object 排序，再依 DueDate 排序
  futureTasks.sort((a, b) => {
    if (a.Object < b.Object) return -1;
    if (a.Object > b.Object) return 1;
    return new Date(a.DueDate).getTime() - new Date(b.DueDate).getTime(); // 次要排序條件：DueDate (升序)
  });
  
  // 調整 `pastTasks` 排序以支援 Object 分組：先依 Object 排序，再依 DueDate 排序
  pastTasks.sort((a, b) => {
    // 主要排序條件：Object (升序)
    if (a.Object < b.Object) return -1;
    if (a.Object > b.Object) return 1;
    
    // 次要排序條件：DueDate (降序)
    return new Date(b.DueDate).getTime() - new Date(a.DueDate).getTime();
  });

  const template = HtmlService.createTemplateFromFile('TimelineDashboard');
  template.lastUpdated = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  template.overdueTasks = overdueTasks;
  template.futureTasks = futureTasks;
  template.pastTasks = pastTasks;

  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}




/**
 * 新增自訂選單
 */
function onOpen() {
  console.log('onOpen function triggered');
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(`🚀 WBS 自動化工具 (${SCRIPT_VERSION})`)
    .addItem('1. 建立新 WBS 工作表', 'initializeWBSSystem')
    .addSeparator()
    .addItem('2. 重設任務內容與公式', 'resetWBSContentFormulas')
    .addItem('3. 套用 Object 顏色標記', 'applyObjectColorCoding')
    .addItem('4. 套用 Resource 顏色標記', 'applyResourceColorCoding')
    .addToUi();
}