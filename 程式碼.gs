// ===== 全域快取和常數 =====
const CACHE = {
  orders: null,
  targetSheet: null,
  sourceData: null,
  sourceDataMap: null,
  // 人員配置快取
  staffConfigData: null
};

// 更新 CONFIG.vlookupCells 設定，加入電話號碼特殊處理標記
const CONFIG = {
  mainSheetName: "主工作表",
  targetSheetName: "工作安排",
  emailSheetName: "電子郵件寄送清單",
  sourceSheetName: "工作安排",
  staffConfigSheetName: "人員配置表",
  calendarId: 'c_3fc4a74cc751b2cbe15fcd2b384d078d86fee64693db9f7c7132b20ef6a475fb@group.calendar.google.com',

  // VLOOKUP 對應的儲存格配置 - 加入電話號碼特殊處理
  vlookupCells: [
    { range: "D4", sourceColumn: 27, defaultValue: "", description: "收單時間" },
    { range: "B6", sourceColumn: 18, defaultValue: "", description: "申請人／單位" },
    { range: "D6", sourceColumn: 22, defaultValue: "", suffix: "人", description: "參訪人數" },
    { range: "B7", sourceColumn: 16, defaultValue: "", description: "預計參訪日期" },
    { range: "D7", sourceColumn: 17, defaultValue: "", description: "時段" },
    { range: "B8", sourceColumns: [20, 21], defaultValue: "", description: "領隊姓名／職稱", concat: true },
    { range: "B9", sourceColumn: 24, defaultValue: "", description: "聯絡電話", isPhone: true },
    { range: "B10", sourceColumn: 25, defaultValue: "", description: "電子郵件" },
    { range: "B11", sourceColumn: 19, defaultValue: "", description: "課程名稱" },
    { range: "B12", sourceColumn: 26, defaultValue: "", prefix: "　", description: "申請備註" },
    { range: "D8", sourceColumn: 15, defaultValue: "", description: "選擇方案" }
  ],

  // 其他設定保持不變...
  cells: {
    dropdown: [
      { range: "B14", column: "C", columnIndex: 0, staffType: "lecturer" },
      { range: "C14", column: "D", columnIndex: 1, staffType: "lecturer" },
      { range: "D14", column: "E", columnIndex: 2, staffType: "lecturer" },
      { range: "B15", column: "F", columnIndex: 3, staffType: "guide" },
      { range: "C15", column: "G", columnIndex: 4, staffType: "guide" },
      { range: "D15", column: "H", columnIndex: 5, staffType: "guide" },
      { range: "B16", column: "I", columnIndex: 6, staffType: "facilitator" },
      { range: "C16", column: "J", columnIndex: 7, staffType: "facilitator" },
      { range: "D16", column: "K", columnIndex: 8, staffType: "facilitator" }
    ],
    checkbox: [
      { range: "D18", column: "M", columnIndex: 10 }
    ],
    text: [
      { range: "B18", column: "L", columnIndex: 9 },
      { range: "C18", column: "N", columnIndex: 11 }
    ]
  }
};

// ===== 人員配置資料載入函數 =====
function getStaffConfigData() {
  if (CACHE.staffConfigData === null) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const staffConfigSheet = spreadsheet.getSheetByName(CONFIG.staffConfigSheetName);
    
    if (!staffConfigSheet) {
      console.error(`找不到工作表：${CONFIG.staffConfigSheetName}`);
      return null;
    }
    
    try {
      // 取得各欄位的資料範圍（A2:A, B2:B, C2:C）
      const lastRowA = getLastNonEmptyRow(staffConfigSheet, 1); // A欄
      const lastRowB = getLastNonEmptyRow(staffConfigSheet, 2); // B欄
      const lastRowC = getLastNonEmptyRow(staffConfigSheet, 3); // C欄
      
      const lecturerData = lastRowA >= 2 ? 
        staffConfigSheet.getRange(2, 1, lastRowA - 1, 1).getValues().flat().filter(v => v !== "") : [];
      const guideData = lastRowB >= 2 ? 
        staffConfigSheet.getRange(2, 2, lastRowB - 1, 1).getValues().flat().filter(v => v !== "") : [];
      const facilitatorData = lastRowC >= 2 ? 
        staffConfigSheet.getRange(2, 3, lastRowC - 1, 1).getValues().flat().filter(v => v !== "") : [];
      
      CACHE.staffConfigData = {
        lecturer: ["未定（無）", ...lecturerData],
        guide: ["未定（無）", ...guideData],
        facilitator: ["未定（無）", ...facilitatorData]
      };
      
      console.log(`載入人員配置資料：講座 ${lecturerData.length} 人，導覽 ${guideData.length} 人，引導 ${facilitatorData.length} 人`);
    } catch (error) {
      console.error("載入人員配置資料時發生錯誤:", error);
      // 使用預設值
      CACHE.staffConfigData = {
        lecturer: ["未定（無）"],
        guide: ["未定（無）"],
        facilitator: ["未定（無）"]
      };
    }
  }

  return CACHE.staffConfigData;
}

// ===== 輔助函數：取得欄位最後非空白行 =====
function getLastNonEmptyRow(sheet, column) {
  const values = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues().flat();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] !== "" && values[i] !== null && values[i] !== undefined) {
      return i + 1; // 轉換為1-based行號
    }
  }
  return 1; // 如果都是空白，回傳第1行
}

// ===== 取得人員選項清單 =====
function getStaffOptions(staffType) {
  const staffData = getStaffConfigData();
  if (!staffData || !staffData[staffType]) {
    console.warn(`無法取得人員類型 ${staffType} 的資料`);
    return ["未定（無）"];
  }
  return staffData[staffType];
}

// ===== 原始資料查詢函數（修改為從工作安排表讀取） =====
function getSourceData() {
  if (CACHE.sourceData === null) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName(CONFIG.sourceSheetName);
    
    if (!sourceSheet) {
      console.error(`找不到工作表：${CONFIG.sourceSheetName}`);
      return null;
    }
    
    // 取得資料範圍（假設資料在 A1 開始，動態取得最後一行）
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) {
      console.warn("工作安排表沒有資料");
      return null;
    }
    
    // 取得所有資料（A1 到最後一欄，共27欄）
    CACHE.sourceData = sourceSheet.getRange(1, 1, lastRow, 27).getValues();
    
    // 建立查詢對應表（以第一欄作為 key）
    CACHE.sourceDataMap = new Map();
    for (let i = 1; i < CACHE.sourceData.length; i++) { // 跳過標題行
      const key = CACHE.sourceData[i][0]; // A 欄作為 key
      if (key) {
        CACHE.sourceDataMap.set(key, CACHE.sourceData[i]);
      }
    }
    
    console.log(`載入工作安排資料：${CACHE.sourceData.length - 1} 筆記錄`);
  }

  return CACHE.sourceData;
}

function getSourceDataByKey(key) {
  getSourceData(); // 確保資料已載入
  return CACHE.sourceDataMap ? CACHE.sourceDataMap.get(key) : null;
}

// ===== 更新：使用 museum_id 查詢行號（每次都重新查詢） =====
function findMatchingRow(selectedOrder) {
  if (!selectedOrder) {
    console.warn("selectedOrder 為空");
    return -1;
  }
  
  const orders = getOrdersData();
  for (let i = 0; i < orders.length; i++) {
    if (orders[i] === selectedOrder) {
      const row = i + 2; // 因為資料從第2行開始
      console.log(`找到單號 ${selectedOrder}，行號：${row}`);
      return row;
    }
  }
  
  console.warn(`找不到單號：${selectedOrder}`);
  return -1;
}

// ===== 更新：每次都重新讀取訂單資料 =====
function getOrdersData() {
  if (CACHE.targetSheet) {
    const orders = CACHE.targetSheet.getRange("A2:A").getValues().flat();
    return orders;
  }
  return [];
}

// ===== 更新：工作安排資料讀取（每次都重新查詢行號） =====
function getWorkArrangementData(selectedOrder) {
  if (!selectedOrder) {
    console.warn("selectedOrder 為空");
    return null;
  }
  
  const row = findMatchingRow(selectedOrder);
  
  if (row !== -1) {
    if (!CACHE.targetSheet) {
      console.error("targetSheet 未初始化");
      return null;
    }
    
    try {
      // 一次讀取整行 C 到 N 欄的資料
      const data = CACHE.targetSheet.getRange(`C${row}:N${row}`).getValues()[0];
      console.log(`載入工作安排資料，單號：${selectedOrder}，行號：${row}`);
      return data;
    } catch (error) {
      console.error("讀取工作安排資料時發生錯誤:", error);
      return null;
    }
  }
  
  return null;
}

function getTargetValueBatch(selectedOrder, columnIndex) {
  const data = getWorkArrangementData(selectedOrder);
  if (data && columnIndex >= 0 && columnIndex < data.length) {
    return data[columnIndex];
  }
  return null;
}
// ===== 寄送申請受理通知 =====
function sendApplicantNotification() {
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

try {
  const applicantEmailRange = spreadsheet.getRangeByName("申請人電子郵件");
  if (!applicantEmailRange) {
    SpreadsheetApp.getUi().alert(
      "錯誤", 
      "找不到命名範圍「申請人電子郵件」，請確認已正確設定。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const applicantEmail = applicantEmailRange.getValue();
  if (!applicantEmail || applicantEmail === "") {
    SpreadsheetApp.getUi().alert(
      "錯誤", 
      "申請人電子郵件為空，請確認已填入正確的電子郵件地址。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // 確認對話框
  const response = SpreadsheetApp.getUi().alert(
    "確認寄送", 
    `確定要寄送申請受理通知給申請人嗎？\n\n收件人：${applicantEmail}\n主旨：檔案素養教育（開箱檔案館–到館學習）系統已受理，請靜候承辦人通知說明`, 
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert("取消", "已取消寄送操作。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const subject = "檔案素養教育（開箱檔案館–到館學習）系統已受理，請靜候承辦人通知說明";
  const body = `您好！您的「開箱檔案館–到館學習」申請案已受理。案件仍在處理中，本局將依申請提出順序及各時段受理情形，安排並告知申請結果，感謝您的耐心等候。
黃宇暘敬上

國家發展委員會檔案管理局
應用服務組約用研究員　黃宇暘
TEL：(02)8995-3979
https://www.archives.gov.tw
244013　新北市林口區檔案館路1號`;
  
  GmailApp.sendEmail(applicantEmail, subject, body);
  
  SpreadsheetApp.getUi().alert(
    "成功", 
    `申請受理通知已成功寄送至：${applicantEmail}`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
} catch (error) {
  console.error("寄送申請人通知時發生錯誤:", error);
  SpreadsheetApp.getUi().alert(
    "錯誤", 
    `寄送申請人通知時發生錯誤：${error.message}`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
}
// ===== 更新：批次寫入工作安排（使用 museum_id 查詢） =====
function batchWriteToWorkArrangement(selectedOrder, updates) {
  if (!selectedOrder || updates.length === 0) {
    console.warn("無法寫入工作安排：單號無效或無更新資料");
    return;
  }

  if (!CACHE.targetSheet) {
    console.error("targetSheet 未初始化");
    return;
  }

  try {
    // 每次都重新查詢行號，確保正確性
    const row = findMatchingRow(selectedOrder);
    
    if (row === -1) {
      console.error(`找不到單號：${selectedOrder}`);
      return;
    }

    // 準備批次更新資料
    const currentData = CACHE.targetSheet.getRange(`C${row}:N${row}`).getValues()[0];
    const newData = [...currentData];

    // 應用所有更新
    updates.forEach(update => {
      if (update.columnIndex >= 0 && update.columnIndex < newData.length) {
        newData[update.columnIndex] = update.value;
      }
    });

    // 一次寫入整行
    CACHE.targetSheet.getRange(`C${row}:N${row}`).setValues([newData]);
    
    console.log(`批次寫入工作安排完成，單號：${selectedOrder}，行號：${row}，更新項目：${updates.length}`);
  } catch (error) {
    console.error("批次寫入工作安排時發生錯誤:", error);
  }
}

// ===== VLOOKUP 替代函數（電話號碼格式修正版） =====
function updateVlookupCells(sheet, selectedOrder) {
  if (!selectedOrder) {
    console.warn("selectedOrder 為空，清空所有 VLOOKUP 儲存格");
    // 批次清空所有相關儲存格
    const ranges = CONFIG.vlookupCells.map(config => config.range);
    const values = CONFIG.vlookupCells.map(() => [""]);
    
    // 使用批次更新
    ranges.forEach((range, index) => {
      sheet.getRange(range).setValue(values[index][0]);
    });
    return;
  }

  const sourceRow = getSourceDataByKey(selectedOrder);
  if (!sourceRow) {
    console.warn(`找不到對應的工作安排資料：${selectedOrder}`);
    // 批次清空所有相關儲存格
    const ranges = CONFIG.vlookupCells.map(config => config.range);
    const values = CONFIG.vlookupCells.map(config => config.defaultValue || "");
    
    ranges.forEach((range, index) => {
      sheet.getRange(range).setValue(values[index]);
    });
    return;
  }

  console.log(`更新 VLOOKUP 儲存格，單號：${selectedOrder}`);

  // 準備批次更新資料
  const batchUpdates = [];

  CONFIG.vlookupCells.forEach(config => {
    let value = config.defaultValue || "";
    
    if (config.concat && config.sourceColumns) {
      // 串接多個欄位（如 B8：姓名 + 職稱）
      const values = config.sourceColumns.map(colIndex => {
        const cellValue = sourceRow[colIndex - 1]; // 轉換為 0-based index
        return cellValue || "";
      }).filter(v => v !== "");
      
      value = values.join("");
    } else if (config.sourceColumn) {
      // 單一欄位
      const cellValue = sourceRow[config.sourceColumn - 1]; // 轉換為 0-based index
      
      // 檢查是否為電話號碼欄位
      if (config.isPhone && cellValue !== null && cellValue !== undefined && cellValue !== "") {
        // 電話號碼特殊處理
        let phoneStr = cellValue.toString();
        
        // 移除所有非數字字符（保留 + 號用於國際號碼）
        phoneStr = phoneStr.replace(/[^\d+]/g, '');
        
        // 台灣手機號碼處理：如果是9位數字，前面加0
        if (/^\d{9}$/.test(phoneStr)) {
          phoneStr = "0" + phoneStr;
        }
        
        value = phoneStr;
      } else {
        // 一般欄位處理
        value = cellValue || config.defaultValue || "";
      }
    }
    
    // 加上前綴或後綴
    if (config.prefix && value) value = config.prefix + value;
    if (config.suffix && value) value = value + config.suffix;
    
    batchUpdates.push({
      range: config.range,
      value: value,
      description: config.description,
      isPhone: config.isPhone || false
    });
  });

  // 批次執行更新
  batchUpdates.forEach(update => {
    try {
      const cell = sheet.getRange(update.range);
      
      // 電話號碼欄位特殊處理
      if (update.isPhone) {
        // 設定為文字格式並強制保持完整格式
        cell.setNumberFormat("@");
        cell.setValue("'" + update.value); // 前面加單引號強制為文字
      } else {
        cell.setValue(update.value);
      }
      
      console.log(`更新 ${update.range} (${update.description}): ${update.value}`);
    } catch (error) {
      console.error(`更新 ${update.range} 時發生錯誤:`, error);
    }
  });
}

// ===== 主要 onEdit 函數（更新版） =====
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.mainSheetName) return;

  const range = e.range;
  const editedCell = range.getA1Notation();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 顯示載入狀態
  const statusCell = sheet.getRange("E5");
  spreadsheet.toast("載入中，請稍候...", "提示", 3);
  statusCell.setValue("載入／儲存中").setBackground("#ffcccc");
  SpreadsheetApp.flush();

  // 初始化快取
  CACHE.targetSheet = e.source.getSheetByName(CONFIG.targetSheetName);
  const selectedOrder = sheet.getRange("A4").getValue();

  // 檢查是否為相關儲存格
  const allCells = [
    ...CONFIG.cells.dropdown.map(c => c.range),
    ...CONFIG.cells.checkbox.map(c => c.range),
    ...CONFIG.cells.text.map(c => c.range),
    "A4"
  ];

  if (!allCells.includes(editedCell)) {
    statusCell.setValue("載入完成").setBackground("#ccffcc");
    spreadsheet.toast("載入完成！", "提示", 3);
    return;
  }

  const isA4Edit = editedCell === "A4";

  // 如果是 A4 編輯，更新所有 VLOOKUP 儲存格
  if (isA4Edit) {
    console.log(`A4 儲存格變更，新值：${selectedOrder}`);
    
    // 更新 VLOOKUP 儲存格
    updateVlookupCells(sheet, selectedOrder);
  }

  // 準備批次處理資料
  const batchOperations = [];
  const mainSheetUpdates = []; // 用於批次更新主工作表

  // 處理下拉式選單（修改為使用動態人員清單）
  CONFIG.cells.dropdown.forEach(config => {
    const cell = sheet.getRange(config.range);
    
    // 取得對應的人員選項清單
    const options = getStaffOptions(config.staffType);
    
    // 設定驗證規則
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options)
      .build();
    cell.setDataValidation(rule);
    
    // 處理值更新
    if (isA4Edit) {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) || "未定（無）";
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    } else if (editedCell === config.range) {
      const value = cell.getValue() || "未定（無）";
      batchOperations.push({ 
        columnIndex: config.columnIndex, 
        value: value,
        column: config.column
      });
    }
  });

  // 處理勾選框
  CONFIG.cells.checkbox.forEach(config => {
    const cell = sheet.getRange(config.range);
    
    if (isA4Edit) {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) === true;
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    } else if (editedCell === config.range) {
      batchOperations.push({ 
        columnIndex: config.columnIndex, 
        value: cell.getValue(),
        column: config.column
      });
    }
  });

  // 處理文字輸入
  CONFIG.cells.text.forEach(config => {
    const cell = sheet.getRange(config.range);
    
    if (isA4Edit) {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) || "無";
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    } else if (editedCell === config.range) {
      const value = cell.getValue() || "無";
      batchOperations.push({ 
        columnIndex: config.columnIndex, 
        value: value,
        column: config.column
      });
    }
  });

  // 批次更新主工作表（A4編輯時）
  if (mainSheetUpdates.length > 0) {
    console.log(`批次更新主工作表，項目數：${mainSheetUpdates.length}`);
    mainSheetUpdates.forEach(update => {
      try {
        sheet.getRange(update.range).setValue(update.value);
      } catch (error) {
        console.error(`更新主工作表 ${update.range} 時發生錯誤:`, error);
      }
    });
  }

  // 批次更新到工作安排表（非A4編輯時）
  if (batchOperations.length > 0 && !isA4Edit) {
    console.log(`批次更新工作安排表，項目數：${batchOperations.length}`);
    batchWriteToWorkArrangement(selectedOrder, batchOperations);
  }

  // 清除部分快取（保留 sourceData 和 staffConfigData 快取以提升效能）
  CACHE.orders = null;

  // 更新狀態
  statusCell.setValue("載入完成").setBackground("#ccffcc");
  spreadsheet.toast("載入完成！", "提示", 3);
  SpreadsheetApp.flush();
}

// ===== 輔助函數（更新版） =====
// 保留舊函數以維持相容性
function getTargetValue(selectedOrder, column) {
  // 將欄位字母轉換為索引
  const workArrangementColumns = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
  const columnIndex = workArrangementColumns.indexOf(column);
  if (columnIndex === -1) {
    console.warn(`無效的欄位：${column}`);
    return null;
  }

  return getTargetValueBatch(selectedOrder, columnIndex);
}

// ===== 手動刷新函數（更新版） =====
function refreshCurrentOrder() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const selectedOrder = mainSheet.getRange("A4").getValue();

  // 清除相關快取
  CACHE.sourceData = null;
  CACHE.sourceDataMap = null;
  CACHE.orders = null;
  CACHE.staffConfigData = null; // 清除人員配置快取

  // 初始化 targetSheet
  CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);

  // 更新 VLOOKUP 儲存格
  updateVlookupCells(mainSheet, selectedOrder);

  // 如果有選擇單號，也更新工作安排相關儲存格
  if (selectedOrder) {
    const mainSheetUpdates = [];

    // 批次讀取並更新所有相關儲存格
    CONFIG.cells.dropdown.forEach(config => {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) || "未定（無）";
      const cell = mainSheet.getRange(config.range);
      
      // 重新設定驗證規則（使用最新的人員清單）
      const options = getStaffOptions(config.staffType);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options)
        .build();
      cell.setDataValidation(rule);
      
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    });

    CONFIG.cells.checkbox.forEach(config => {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) === true;
      const cell = mainSheet.getRange(config.range);
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    });

    CONFIG.cells.text.forEach(config => {
      const targetValue = getTargetValueBatch(selectedOrder, config.columnIndex) || "無";
      const cell = mainSheet.getRange(config.range);
      if (cell.getValue() !== targetValue) {
        mainSheetUpdates.push({
          range: config.range,
          value: targetValue
        });
      }
    });

    // 批次執行更新
    if (mainSheetUpdates.length > 0) {
      console.log(`手動刷新：批次更新 ${mainSheetUpdates.length} 個儲存格`);
      mainSheetUpdates.forEach(update => {
        try {
          mainSheet.getRange(update.range).setValue(update.value);
        } catch (error) {
          console.error(`手動刷新更新 ${update.range} 時發生錯誤:`, error);
        }
      });
    }
  }

  spreadsheet.toast("資料已刷新完成！", "提示", 2);
}

// ===== 清除快取函數（更新版） =====
function clearAllCache() {
  CACHE.orders = null;
  CACHE.targetSheet = null;
  CACHE.sourceData = null;
  CACHE.sourceDataMap = null;
  CACHE.staffConfigData = null;

  console.log("所有快取已清除");
  SpreadsheetApp.getActiveSpreadsheet().toast("快取已清除", "提示", 2);
}

// ===== 新增到 Google 日曆功能 =====
function addToCalendar() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    spreadsheet.toast("正在新增到 Google 日曆...", "處理中", 3);
    
    // 取得日曆
    let calendar;
    try {
      calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
      if (!calendar) throw new Error("無法存取指定的日曆");
    } catch (calendarError) {
      console.warn("無法存取特定日曆，改用主要日曆:", calendarError);
      SpreadsheetApp.getUi().alert(
        "提醒", 
        "無法存取指定的日曆，將使用您的主要日曆。\n\n如需使用特定日曆，請確認該日曆已分享給您並具有編輯權限。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      calendar = CalendarApp.getDefaultCalendar();
    }
    
    // 批次取得命名範圍
    const ranges = {};
    const rangeNames = ["參訪日期", "參訪時段", "參訪單位", "選擇方案"];
    
    for (const name of rangeNames) {
      try {
        ranges[name] = spreadsheet.getRangeByName(name);
        if (!ranges[name]) throw new Error(`找不到命名範圍：${name}`);
      } catch (error) {
        if (name !== "選擇方案") {  // 選擇方案是可選的
          SpreadsheetApp.getUi().alert(
            "錯誤", 
            `找不到必要的命名範圍「${name}」，請確認已正確設定。`, 
            SpreadsheetApp.getUi().ButtonSet.OK
          );
          return;
        }
      }
    }
    
    // 取得資料
    const visitDate = ranges["參訪日期"].getValue();
    const visitTime = ranges["參訪時段"].getValue();
    const visitUnit = ranges["參訪單位"].getValue();
    const selectedPlan = ranges["選擇方案"] ? ranges["選擇方案"].getValue() : "";
    
    // 驗證資料
    if (!visitDate || !visitTime || !visitUnit) {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        "參訪資料不完整，請確認已填入：\n• 參訪日期\n• 參訪時段\n• 參訪單位", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 解析日期和時間
    const parsedDate = visitDate instanceof Date ? new Date(visitDate) : new Date(visitDate.toString());
    if (isNaN(parsedDate.getTime())) {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        "參訪日期格式錯誤，請確認日期格式正確（例如：2025/9/24）。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 解析時間範圍
    const [startTime, endTime] = visitTime.toString().split("-").map(t => t.trim());
    const timePattern = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;
    
    if (!startTime || !endTime || !timePattern.test(startTime) || !timePattern.test(endTime)) {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        "參訪時段格式錯誤，請確認格式為 HH:MM-HH:MM\n（例如：14:00-17:00）。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 建立開始和結束時間
    const startDateTime = new Date(parsedDate);
    const endDateTime = new Date(parsedDate);
    
    const [startHour, startMin] = startTime.split(":").map(Number);
    const [endHour, endMin] = endTime.split(":").map(Number);
    
    startDateTime.setHours(startHour, startMin, 0, 0);
    endDateTime.setHours(endHour, endMin, 0, 0);
    
    if (startDateTime >= endDateTime) {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        "結束時間必須晚於開始時間，請檢查參訪時段設定。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 建立活動
    const eventTitle = `${visitUnit}－到館學習`;
    const orderNumber = spreadsheet.getSheetByName(CONFIG.mainSheetName).getRange("A4").getValue() || "";
    
    let eventDescription = `參訪單位：${visitUnit}\n參訪時間：${visitTime}\n活動類型：檔案素養教育－到館學習`;
    if (selectedPlan) eventDescription += `\n方案：${selectedPlan}`;
    if (orderNumber) eventDescription += `\n單號：${orderNumber}`;
    eventDescription += `\n\n※ 此活動由 Google Apps Script 自動建立`;
    
    // 檢查重複活動
    const existingEvents = calendar.getEventsForDay(parsedDate);
    const duplicateEvent = existingEvents.find(event => 
      event.getTitle() === eventTitle && 
      event.getStartTime().getTime() === startDateTime.getTime()
    );
    
    if (duplicateEvent) {
      const response = SpreadsheetApp.getUi().alert(
        "重複活動", 
        `在 ${Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy/MM/dd")} 已存在相同的活動：\n「${eventTitle}」\n\n是否仍要建立新活動？`, 
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      
      if (response !== SpreadsheetApp.getUi().Button.YES) return;
    }
    
    // 新增到日曆
    calendar.createEvent(eventTitle, startDateTime, endDateTime, {
      description: eventDescription,
      location: "國家發展委員會檔案管理局\n244013 新北市林口區檔案館路1號"
    });
    
    const successMessage = `活動已成功新增到 Google 日曆！\n\n活動標題：${eventTitle}\n日期：${Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy/MM/dd")}\n時間：${startTime} - ${endTime}\n地點：國家發展委員會檔案管理局`;
    
    SpreadsheetApp.getUi().alert("成功", successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    spreadsheet.toast("活動已成功新增到 Google 日曆！", "完成", 3);
    
  } catch (error) {
    console.error("新增日曆事件時發生錯誤:", error);
    SpreadsheetApp.getUi().alert(
      "錯誤", 
      `新增日曆事件時發生錯誤：\n${error.message}\n\n請檢查：\n1. 網路連線是否正常\n2. 是否有 Google 日曆的存取權限\n3. 資料格式是否正確`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== PDF 匯出重試函數 =====
function fetchPDFWithRetry(url, token, maxRetries = 3, delay = 2000) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`PDF 匯出嘗試 ${attempt}/${maxRetries}`);
      
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        console.log(`PDF 匯出成功（嘗試 ${attempt}）`);
        return response;
      } else {
        console.error(`PDF 匯出失敗（嘗試 ${attempt}）：HTTP ${response.getResponseCode()}`);
        if (attempt === maxRetries) {
          throw new Error(`PDF 匯出失敗：HTTP ${response.getResponseCode()}`);
        }
      }
    } catch (error) {
      console.error(`PDF 匯出錯誤（嘗試 ${attempt}）：${error.message}`);
      if (attempt === maxRetries) throw error;
    }
    
    if (attempt < maxRetries) {
      console.log(`等待 ${delay}ms 後重試...`);
      Utilities.sleep(delay);
      delay *= 1.5;
    }
  }
}

// ===== 匯出並寄送 PDF（更新版） =====
function exportAndSendPDF() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 確認對話框
  const response = SpreadsheetApp.getUi().alert(
    "確認寄送", 
    "確定要生成 PDF 並寄送工作配置提醒郵件嗎？\n\n此操作將會：\n• 生成工作配置單 PDF\n• 生成參訪程序表 PDF\n• 寄送郵件給相關人員", 
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert("取消", "已取消寄送操作。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getUi().alert("處理中", "正在生成 PDF，請稍候...", SpreadsheetApp.getUi().ButtonSet.OK);

  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const emailSheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  let scheduleSheetWasHidden = false;

  try {
    // 取得基本資料
    const orderName = mainSheet.getRange("A4").getValue() || "未命名";
    const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
    
    // 取得日期和時段
    let visitDate = "未定";
    let visitTime = "未定";
    
    try {
      const visitDateValue = spreadsheet.getRangeByName("參訪日期").getValue();
      visitDate = visitDateValue instanceof Date ? 
        Utilities.formatDate(visitDateValue, Session.getScriptTimeZone(), "yyyy/MM/dd") : 
        visitDateValue.toString();
      
      visitTime = spreadsheet.getRangeByName("參訪時段").getValue().toString();
    } catch (error) {
      console.warn("無法取得參訪日期或時段:", error);
    }
    
    const emailSubject = `【通知】檔案素養教育－到館學習：工作配置及參訪程序提醒（${visitDate}，${visitTime}）`;
    
    // 決定參訪程序表
    let scheduleSheet, scheduleRange;
    if (selectedPlan === "2小時方案") {
      scheduleSheet = spreadsheet.getSheetByName("（勿動）參訪程序表-2小時");
      scheduleRange = "A1:E12";
    } else if (selectedPlan === "3小時方案") {
      scheduleSheet = spreadsheet.getSheetByName("（勿動）參訪程序表-3小時");
      scheduleRange = "A1:E13";
    } else {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        `無法識別的選擇方案：${selectedPlan}。請確認選擇方案為「2小時方案」或「3小時方案」。`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 檢查並顯示隱藏的工作表
    if (scheduleSheet.isSheetHidden()) {
      scheduleSheetWasHidden = true;
      scheduleSheet.showSheet();
      SpreadsheetApp.flush();
    }
    
    // 生成 PDF
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    // 主工作表 PDF
    const mainSheetUrl = `${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`;
    const mainSheetBlob = fetchPDFWithRetry(mainSheetUrl, token)
      .getBlob()
      .setName(`${orderName}_工作配置單.pdf`);
    
    // 參訪程序表 PDF
    const scheduleSheetUrl = `${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}&range=${scheduleRange}`;
    const scheduleSheetBlob = fetchPDFWithRetry(scheduleSheetUrl, token)
      .getBlob()
      .setName(`${orderName}_參訪程序表_${selectedPlan}.pdf`);
    
    // 批次取得收件人清單（更新版）
    const namesInSheet = mainSheet.getRange("B14:D16").getValues().flat().filter(n => n);
    const emailData = emailSheet.getRange("A2:B").getValues();
    
    const matchedEmails = emailData
      .filter(row => row[1] && namesInSheet.includes(row[1]))
      .map(row => row[0])
      .filter(email => email);
    
    const mainRecipient = "yyhuang@archives.gov.tw";
    const defaultCcEmails = [
      "hcsun@archives.gov.tw",
      "yhwei@archives.gov.tw", 
      "sihung@archives.gov.tw",
      "tjwu@archives.gov.tw"
    ];
    
    const allCcEmails = [...new Set([...matchedEmails, ...defaultCcEmails])].filter(email => email);
    
    // 寄送郵件
    GmailApp.sendEmail(mainRecipient, emailSubject, 
      "請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/WO73VZ", 
      {
        attachments: [mainSheetBlob, scheduleSheetBlob],
        cc: allCcEmails.join(",")
      }
    );
    
    SpreadsheetApp.getUi().alert(
      "成功", 
      `工作配置提醒PDF（含${selectedPlan}參訪程序表）已生成並寄送完成！`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error("PDF 匯出或寄送過程中發生錯誤:", error);
    SpreadsheetApp.getUi().alert(
      "錯誤", 
      `處理過程中發生錯誤：${error.message}\n\n請稍後再試，或聯繫系統管理員。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } finally {
    // 恢復隱藏狀態
    if (scheduleSheetWasHidden && scheduleSheet) {
      scheduleSheet.hideSheet();
    }
    SpreadsheetApp.flush();
  }
}

// ===== 匯出並寄送 PDF（手動編修版） =====
function exportAndSendPDFmanual() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 確認對話框
  const response = SpreadsheetApp.getUi().alert(
    "確認寄送（手動編修版）", 
    "確定要生成 PDF 並寄送工作配置提醒郵件嗎？\n\n此操作將會：\n• 生成工作配置單 PDF\n• 生成參訪程序表 PDF（手動編修版）\n• 寄送郵件給相關人員", 
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert("取消", "已取消寄送操作。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getUi().alert("處理中", "正在生成 PDF（手動編修版），請稍候...", SpreadsheetApp.getUi().ButtonSet.OK);

  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const emailSheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  let scheduleSheetWasHidden = false;

  try {
    // 取得基本資料
    const orderName = mainSheet.getRange("A4").getValue() || "未命名";
    const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
    
    // 取得日期和時段
    let visitDate = "未定";
    let visitTime = "未定";
    
    try {
      const visitDateValue = spreadsheet.getRangeByName("參訪日期").getValue();
      visitDate = visitDateValue instanceof Date ? 
        Utilities.formatDate(visitDateValue, Session.getScriptTimeZone(), "yyyy/MM/dd") : 
        visitDateValue.toString();
      
      visitTime = spreadsheet.getRangeByName("參訪時段").getValue().toString();
    } catch (error) {
      console.warn("無法取得參訪日期或時段:", error);
    }
    
    const emailSubject = `【通知】檔案素養教育－到館學習：工作配置及參訪程序提醒（${visitDate}，${visitTime}）`;
    
    // 決定參訪程序表（手動編修版）
    let scheduleSheet, scheduleRange;
    if (selectedPlan === "2小時方案") {
      scheduleSheet = spreadsheet.getSheetByName("參訪程序表-2小時（手動編修版）");
      scheduleRange = "A1:F12";
    } else if (selectedPlan === "3小時方案") {
      scheduleSheet = spreadsheet.getSheetByName("參訪程序表-3小時（手動編修版）");
      scheduleRange = "A1:E13";
    } else {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        `無法識別的選擇方案：${selectedPlan}。請確認選擇方案為「2小時方案」或「3小時方案」。`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 檢查手動編修版工作表是否存在
    if (!scheduleSheet) {
      SpreadsheetApp.getUi().alert(
        "錯誤", 
        `找不到工作表：${selectedPlan === "2小時方案" ? "參訪程序表-2小時（手動編修版）" : "參訪程序表-3小時（手動編修版）"}\n\n請確認該工作表已建立。`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 檢查並顯示隱藏的工作表
    if (scheduleSheet.isSheetHidden()) {
      scheduleSheetWasHidden = true;
      scheduleSheet.showSheet();
      SpreadsheetApp.flush();
    }
    
    // 生成 PDF
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    // 主工作表 PDF
    const mainSheetUrl = `${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`;
    const mainSheetBlob = fetchPDFWithRetry(mainSheetUrl, token)
      .getBlob()
      .setName(`${orderName}_工作配置單.pdf`);
    
    // 參訪程序表 PDF（手動編修版）
    const scheduleSheetUrl = `${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}&range=${scheduleRange}`;
    const scheduleSheetBlob = fetchPDFWithRetry(scheduleSheetUrl, token)
      .getBlob()
      .setName(`${orderName}_參訪程序表_${selectedPlan}（手動編修版）.pdf`);
    
    // 批次取得收件人清單（更新版）
    const namesInSheet = mainSheet.getRange("B14:D16").getValues().flat().filter(n => n);
    const emailData = emailSheet.getRange("A2:B").getValues();
    
    const matchedEmails = emailData
      .filter(row => row[1] && namesInSheet.includes(row[1]))
      .map(row => row[0])
      .filter(email => email);
    
    const mainRecipient = "yyhuang@archives.gov.tw";
    const defaultCcEmails = [
      "hcsun@archives.gov.tw",
      "yhwei@archives.gov.tw", 
      "sihung@archives.gov.tw",
      "mjjheng@archives.gov.tw"
    ];
    
    const allCcEmails = [...new Set([...matchedEmails, ...defaultCcEmails])].filter(email => email);
    
    // 寄送郵件
    GmailApp.sendEmail(mainRecipient, emailSubject, 
      "請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/WO73VZ\n\n※ 此版本使用手動編修版參訪程序表", 
      {
        attachments: [mainSheetBlob, scheduleSheetBlob],
        cc: allCcEmails.join(",")
      }
    );
    
    SpreadsheetApp.getUi().alert(
      "成功", 
      `工作配置提醒PDF（含${selectedPlan}參訪程序表手動編修版）已生成並寄送完成！`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error("PDF 匯出或寄送過程中發生錯誤:", error);
    SpreadsheetApp.getUi().alert(
      "錯誤", 
      `處理過程中發生錯誤：${error.message}\n\n請稍後再試，或聯繫系統管理員。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } finally {
    // 恢復隱藏狀態
    if (scheduleSheetWasHidden && scheduleSheet) {
      scheduleSheet.hideSheet();
    }
    SpreadsheetApp.flush();
  }
}

// ===== 自訂選單建立函數 =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 工作配置管理')
    .addItem('🔄 手動刷新資料', 'refreshCurrentOrder')
    .addItem('🗑️ 清除快取', 'clearAllCache')
    .addSeparator()
    .addItem('📅 新增到 Google 日曆', 'addToCalendar')
    .addSeparator()
    .addItem('📧 匯出並寄送 PDF', 'exportAndSendPDF')
    .addItem('📧 匯出並寄送 PDF（手動編修版）', 'exportAndSendPDFmanual')
    .addSeparator()
    .addItem('ℹ️ 關於', 'showAbout')
    .addToUi();
}

// ===== 關於對話框 =====
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const aboutText = `
工作配置管理系統 v2.0
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

主要功能：
• 自動同步工作安排資料
• 動態人員配置管理
• Google 日曆整合
• PDF 自動生成與郵件寄送

更新說明（v2.0）：
✓ 改用 museum_id（單號）作為查詢鍵
✓ 解決動態排序導致的資料錯位問題
✓ 優化快取機制，提升效能
✓ 支援動態人員清單配置

使用說明：
1. 在 A4 儲存格選擇單號
2. 系統自動載入對應資料
3. 編輯工作配置後自動儲存
4. 使用選單功能匯出或寄送

注意事項：
• 請勿手動修改「工作安排」表的排序
• 如遇資料不同步，請使用「手動刷新資料」
• 如遇效能問題，請使用「清除快取」

開發者：國發會檔案管理局
版本：2.0.0（2025-11-03）
  `;
  
  ui.alert('關於工作配置管理系統', aboutText, ui.ButtonSet.OK);
}

// ===== 測試函數：驗證 museum_id 查詢功能 =====
function testFindMatchingRow() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const selectedOrder = mainSheet.getRange("A4").getValue();
  
  console.log("=== 測試 museum_id 查詢功能 ===");
  console.log(`選擇的單號：${selectedOrder}`);
  
  // 初始化 targetSheet
  CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  
  // 測試查詢
  const row = findMatchingRow(selectedOrder);
  console.log(`查詢結果：行號 ${row}`);
  
  if (row !== -1) {
    // 測試讀取資料
    const data = getWorkArrangementData(selectedOrder);
    console.log("讀取的工作安排資料：", data);
    
    // 測試讀取特定欄位
    const lecturer1 = getTargetValueBatch(selectedOrder, 0);
    console.log(`講座人員1：${lecturer1}`);
  } else {
    console.error("找不到對應的單號");
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("測試完成，請查看日誌", "測試", 3);
}

// ===== 測試函數：驗證批次寫入功能 =====
function testBatchWrite() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const selectedOrder = mainSheet.getRange("A4").getValue();
  
  console.log("=== 測試批次寫入功能 ===");
  console.log(`選擇的單號：${selectedOrder}`);
  
  // 初始化 targetSheet
  CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  
  // 準備測試資料
  const testUpdates = [
    { columnIndex: 0, value: "測試講座1" },
    { columnIndex: 9, value: "測試備註" }
  ];
  
  console.log("測試更新資料：", testUpdates);
  
  // 執行批次寫入
  batchWriteToWorkArrangement(selectedOrder, testUpdates);
  
  // 驗證寫入結果
  const data = getWorkArrangementData(selectedOrder);
  console.log("寫入後的資料：", data);
  console.log(`講座人員1：${data[0]}`);
  console.log(`備註：${data[9]}`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("測試完成，請查看日誌", "測試", 3);
}

// ===== 測試函數：驗證人員配置載入 =====
function testStaffConfig() {
  console.log("=== 測試人員配置載入 ===");
  
  // 清除快取
  CACHE.staffConfigData = null;
  
  // 載入人員配置
  const staffData = getStaffConfigData();
  
  console.log("講座人員清單：", staffData.lecturer);
  console.log("導覽人員清單：", staffData.guide);
  console.log("引導人員清單：", staffData.facilitator);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("測試完成，請查看日誌", "測試", 3);
}

// ===== 測試函數：完整流程測試 =====
function testFullWorkflow() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  
  console.log("=== 完整流程測試 ===");
  
  // 1. 測試人員配置載入
  console.log("\n1. 測試人員配置載入...");
  CACHE.staffConfigData = null;
  const staffData = getStaffConfigData();
  console.log(`講座人員數量：${staffData.lecturer.length}`);
  console.log(`導覽人員數量：${staffData.guide.length}`);
  console.log(`引導人員數量：${staffData.facilitator.length}`);
  
  // 2. 測試原始資料載入
  console.log("\n2. 測試原始資料載入...");
  CACHE.sourceData = null;
  CACHE.sourceDataMap = null;
  const sourceData = getSourceData();
  console.log(`原始資料筆數：${sourceData ? sourceData.length - 1 : 0}`);
  
  // 3. 測試單號查詢
  console.log("\n3. 測試單號查詢...");
  const selectedOrder = mainSheet.getRange("A4").getValue();
  console.log(`選擇的單號：${selectedOrder}`);
  
  CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  const row = findMatchingRow(selectedOrder);
  console.log(`查詢到的行號：${row}`);
  
  // 4. 測試資料讀取
  if (row !== -1) {
    console.log("\n4. 測試資料讀取...");
    const data = getWorkArrangementData(selectedOrder);
    console.log(`工作安排資料：`, data);
    
    // 5. 測試 VLOOKUP 更新
    console.log("\n5. 測試 VLOOKUP 更新...");
    const sourceRow = getSourceDataByKey(selectedOrder);
    if (sourceRow) {
      console.log(`申請人／單位：${sourceRow[17]}`); // 第18欄
      console.log(`參訪人數：${sourceRow[21]}`);     // 第22欄
      console.log(`預計參訪日期：${sourceRow[15]}`); // 第16欄
    }
  }
  
  console.log("\n=== 測試完成 ===");
  SpreadsheetApp.getActiveSpreadsheet().toast("完整流程測試完成，請查看日誌", "測試", 5);
}

// ===== 診斷函數：檢查系統狀態 =====
function diagnoseSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  console.log("=== 系統診斷開始 ===");
  
  let diagnosticReport = "系統診斷報告\n";
  diagnosticReport += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  
  // 1. 檢查工作表
  diagnosticReport += "1. 工作表檢查：\n";
  const requiredSheets = [
    CONFIG.mainSheetName,
    CONFIG.targetSheetName,
    CONFIG.emailSheetName,
    CONFIG.sourceSheetName,
    CONFIG.staffConfigSheetName
  ];
  
  let allSheetsExist = true;
  requiredSheets.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      diagnosticReport += `   ✓ ${sheetName}\n`;
      console.log(`✓ 工作表存在：${sheetName}`);
    } else {
      diagnosticReport += `   ✗ ${sheetName} (缺少)\n`;
      console.error(`✗ 工作表缺少：${sheetName}`);
      allSheetsExist = false;
    }
  });
  
  // 2. 檢查命名範圍
  diagnosticReport += "\n2. 命名範圍檢查：\n";
  const requiredRanges = ["參訪日期", "參訪時段", "參訪單位", "選擇方案"];
  
  let allRangesExist = true;
  requiredRanges.forEach(rangeName => {
    const range = spreadsheet.getRangeByName(rangeName);
    if (range) {
      diagnosticReport += `   ✓ ${rangeName}\n`;
      console.log(`✓ 命名範圍存在：${rangeName}`);
    } else {
      diagnosticReport += `   ✗ ${rangeName} (缺少)\n`;
      console.error(`✗ 命名範圍缺少：${rangeName}`);
      allRangesExist = false;
    }
  });
  
  // 3. 檢查人員配置資料
  diagnosticReport += "\n3. 人員配置檢查：\n";
  CACHE.staffConfigData = null;
  const staffData = getStaffConfigData();
  
  if (staffData) {
    diagnosticReport += `   ✓ 講座人員：${staffData.lecturer.length} 人\n`;
    diagnosticReport += `   ✓ 導覽人員：${staffData.guide.length} 人\n`;
    diagnosticReport += `   ✓ 引導人員：${staffData.facilitator.length} 人\n`;
    console.log(`✓ 人員配置載入成功`);
  } else {
    diagnosticReport += `   ✗ 人員配置載入失敗\n`;
    console.error(`✗ 人員配置載入失敗`);
  }
  
  // 4. 檢查原始資料
  diagnosticReport += "\n4. 原始資料檢查：\n";
  CACHE.sourceData = null;
  CACHE.sourceDataMap = null;
  const sourceData = getSourceData();
  
  if (sourceData && sourceData.length > 1) {
    diagnosticReport += `   ✓ 資料筆數：${sourceData.length - 1} 筆\n`;
    console.log(`✓ 原始資料載入成功：${sourceData.length - 1} 筆`);
  } else {
    diagnosticReport += `   ✗ 無資料或載入失敗\n`;
    console.error(`✗ 原始資料載入失敗`);
  }
  
  // 5. 檢查當前選擇的單號
  diagnosticReport += "\n5. 當前單號檢查：\n";
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const selectedOrder = mainSheet.getRange("A4").getValue();
  
  if (selectedOrder) {
    diagnosticReport += `   ✓ 選擇的單號：${selectedOrder}\n`;
    
    CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    const row = findMatchingRow(selectedOrder);
    
    if (row !== -1) {
      diagnosticReport += `   ✓ 查詢到的行號：${row}\n`;
      console.log(`✓ 單號查詢成功：${selectedOrder} -> 行號 ${row}`);
    } else {
      diagnosticReport += `   ✗ 找不到對應的行號\n`;
      console.error(`✗ 單號查詢失敗：${selectedOrder}`);
    }
  } else {
    diagnosticReport += `   - 未選擇單號\n`;
    console.log(`- 未選擇單號`);
  }
  
  // 6. 檢查日曆權限
  diagnosticReport += "\n6. Google 日曆權限檢查：\n";
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
    if (calendar) {
      diagnosticReport += `   ✓ 可存取指定日曆\n`;
      console.log(`✓ 日曆存取正常`);
    } else {
      diagnosticReport += `   ⚠ 無法存取指定日曆（將使用預設日曆）\n`;
      console.warn(`⚠ 無法存取指定日曆`);
    }
  } catch (error) {
    diagnosticReport += `   ⚠ 日曆權限檢查失敗（將使用預設日曆）\n`;
    console.warn(`⚠ 日曆權限檢查失敗：${error.message}`);
  }
  
  // 7. 總結
  diagnosticReport += "\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  if (allSheetsExist && allRangesExist && staffData && sourceData) {
    diagnosticReport += "診斷結果：✓ 系統狀態正常\n";
    console.log("=== 診斷結果：系統狀態正常 ===");
  } else {
    diagnosticReport += "診斷結果：✗ 發現問題，請檢查上述項目\n";
    console.error("=== 診斷結果：發現問題 ===");
  }
  
  // 顯示診斷報告
  ui.alert("系統診斷報告", diagnosticReport, ui.ButtonSet.OK);
  console.log(diagnosticReport);
}

// ===== 輔助函數：批次清空儲存格 =====
function clearCellsBatch(sheet, ranges) {
  ranges.forEach(range => {
    try {
      sheet.getRange(range).clearContent();
    } catch (error) {
      console.error(`清空儲存格 ${range} 時發生錯誤:`, error);
    }
  });
}

// ===== 輔助函數：格式化日期 =====
function formatDate(date) {
  if (!date) return "";
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  return date.toString();
}

// ===== 輔助函數：驗證電子郵件格式 =====
function isValidEmail(email) {
  if (!email) return false;
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email.toString().trim());
}

// ===== 輔助函數：取得工作表資料範圍 =====
function getSheetDataRange(sheetName, startRow = 2) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error(`找不到工作表：${sheetName}`);
    return null;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    console.warn(`工作表 ${sheetName} 沒有資料`);
    return null;
  }
  
  const lastColumn = sheet.getLastColumn();
  return sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn);
}

// ===== 輔助函數：安全取得儲存格值 =====
function getSafeValue(sheet, range, defaultValue = "") {
  try {
    const value = sheet.getRange(range).getValue();
    return value !== null && value !== undefined ? value : defaultValue;
  } catch (error) {
    console.error(`取得儲存格 ${range} 值時發生錯誤:`, error);
    return defaultValue;
  }
}

// ===== 輔助函數：安全設定儲存格值 =====
function setSafeValue(sheet, range, value) {
  try {
    sheet.getRange(range).setValue(value);
    return true;
  } catch (error) {
    console.error(`設定儲存格 ${range} 值時發生錯誤:`, error);
    return false;
  }
}

// ===== 效能監控函數 =====
function measurePerformance(functionName, func) {
  const startTime = new Date().getTime();
  const result = func();
  const endTime = new Date().getTime();
  const duration = endTime - startTime;
  
  console.log(`[效能] ${functionName} 執行時間：${duration}ms`);
  
  if (duration > 5000) {
    console.warn(`[效能警告] ${functionName} 執行時間過長：${duration}ms`);
  }
  
  return result;
}

// ===== 錯誤處理包裝函數 =====
function safeExecute(functionName, func, errorCallback = null) {
  try {
    console.log(`[執行] ${functionName} 開始`);
    const result = func();
    console.log(`[執行] ${functionName} 完成`);
    return result;
  } catch (error) {
    console.error(`[錯誤] ${functionName} 執行失敗:`, error);
    
    if (errorCallback) {
      errorCallback(error);
    } else {
      SpreadsheetApp.getUi().alert(
        "執行錯誤",
        `${functionName} 執行時發生錯誤：\n${error.message}\n\n請查看日誌以獲取更多資訊。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return null;
  }
}

// ===== 日誌記錄函數 =====
function logOperation(operation, details) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const logMessage = `[${timestamp}] ${operation}: ${JSON.stringify(details)}`;
  console.log(logMessage);
  
  // 可選：將日誌寫入專用工作表
  // writeToLogSheet(timestamp, operation, details);
}

// ===== 可選：寫入日誌工作表 =====
function writeToLogSheet(timestamp, operation, details) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName("系統日誌");
    
    // 如果日誌工作表不存在，則建立
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet("系統日誌");
      logSheet.getRange("A1:C1").setValues([["時間", "操作", "詳細資訊"]]);
      logSheet.getRange("A1:C1").setFontWeight("bold");
    }
    
    // 寫入日誌
    logSheet.appendRow([timestamp, operation, JSON.stringify(details)]);
    
    // 限制日誌行數（保留最近1000筆）
    const maxRows = 1000;
    if (logSheet.getLastRow() > maxRows + 1) {
      logSheet.deleteRows(2, logSheet.getLastRow() - maxRows - 1);
    }
  } catch (error) {
    console.error("寫入日誌工作表時發生錯誤:", error);
  }
}

// ===== 備份函數：備份當前設定 =====
function backupCurrentSettings() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const selectedOrder = mainSheet.getRange("A4").getValue();
  
  if (!selectedOrder) {
    SpreadsheetApp.getUi().alert("錯誤", "請先選擇單號", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const backupData = {
      order: selectedOrder,
      timestamp: timestamp,
      dropdowns: {},
      checkboxes: {},
      texts: {}
    };
    
    // 備份下拉式選單
    CONFIG.cells.dropdown.forEach(config => {
      backupData.dropdowns[config.range] = mainSheet.getRange(config.range).getValue();
    });
    
    // 備份勾選框
    CONFIG.cells.checkbox.forEach(config => {
      backupData.checkboxes[config.range] = mainSheet.getRange(config.range).getValue();
    });
    
    // 備份文字輸入
    CONFIG.cells.text.forEach(config => {
      backupData.texts[config.range] = mainSheet.getRange(config.range).getValue();
    });
    
    console.log(`備份完成：${JSON.stringify(backupData)}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`設定已備份（${timestamp}）`, "備份完成", 3);
    
    return backupData;
  } catch (error) {
    console.error("備份設定時發生錯誤:", error);
    SpreadsheetApp.getUi().alert("錯誤", `備份失敗：${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== 版本資訊 =====
const VERSION_INFO = {
  version: "2.0.0",
  releaseDate: "2025-11-03",
  changes: [
    "改用 museum_id（單號）作為查詢鍵",
    "解決動態排序導致的資料錯位問題",
    "優化快取機制，提升效能",
    "支援動態人員清單配置",
    "新增系統診斷功能",
    "改進錯誤處理機制",
    "新增效能監控功能"
  ]
};

// ===== 顯示版本資訊 =====
function showVersionInfo() {
  let versionText = `版本：${VERSION_INFO.version}\n`;
  versionText += `發布日期：${VERSION_INFO.releaseDate}\n\n`;
  versionText += `更新內容：\n`;
  VERSION_INFO.changes.forEach((change, index) => {
    versionText += `${index + 1}. ${change}\n`;
  });
  
  console.log(versionText);
  SpreadsheetApp.getUi().alert("版本資訊", versionText, SpreadsheetApp.getUi().ButtonSet.OK);
}

