// ===== 1. 網頁入口 =====
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('檔案素養教育工作配置系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== 2. 取得初始資料 (修改：加入「已安排」狀態判斷) =====
function getInitialDataAPI() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);

  const targetSheet = CACHE.targetSheet;
  const lastRow = targetSheet.getLastRow();
  
  const orderIds = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const checkedStatuses = targetSheet.getRange(2, 13, lastRow - 1, 1).getValues().flat(); 
  
  const orders = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 設定為當天凌晨，只比對日期

  orderIds.forEach((id, index) => {
    if (!id) return;
    let year = "未知年份";
    const idStr = String(id);
    if (idStr.startsWith("MS") && idStr.length >= 6) year = idStr.substring(2, 6);

    let isClosed = false;
    let isArranged = false; // 新增：已安排狀態
    
    const sourceRow = getSourceDataByKey(id); 
    if (sourceRow) {
      const visitDate = sourceRow[15];
      const isChecked = checkedStatuses[index] === true;
      
      if (visitDate instanceof Date) {
        // 邏輯判斷：
        // 1. 日期已過 且 已核勾選 -> 已結案 (Closed)
        // 2. 日期未過(含當天) 且 已核勾選 -> 已安排 (Arranged)
        if (visitDate < today && isChecked) {
          isClosed = true;
        } else if (visitDate >= today && isChecked) {
          isArranged = true;
        }
      }
    }

    orders.push({ 
      id: idStr, 
      year: year, 
      isClosed: isClosed,
      isArranged: isArranged // 回傳給前端
    });
  });

  const staffConfig = getStaffConfigData();
  const stats = getStaffStatistics();

  return {
    orders: orders,
    staff: staffConfig,
    stats: stats,
    serverYear: new Date().getFullYear().toString()
  };
}

// ===== 3. 讀取單筆詳細資料 =====
function getOrderDataAPI(orderId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    
    const lastRow = CACHE.targetSheet.getLastRow();
    const orderList = CACHE.targetSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    let row = -1;
    const targetId = String(orderId).trim();

    for (let i = 0; i < orderList.length; i++) {
      if (String(orderList[i]).trim() === targetId) {
        row = i + 2;
        break;
      }
    }

    if (row === -1) return { error: `找不到單號 (${orderId})` };

    const workData = CACHE.targetSheet.getRange(row, 3, 1, 12).getValues()[0];
    const sourceRow = getSourceDataByKey(orderId); 
    
    return {
      applicant: {
        unit: sourceRow ? sourceRow[17] : "(查無資料)",  
        date: sourceRow ? formatDate(sourceRow[15]) : "",
        time: sourceRow ? sourceRow[16] : "",
        count: sourceRow ? sourceRow[21] : "",
        leader: sourceRow ? (sourceRow[19] || "") + " " + (sourceRow[20] || "") : "",
        phone: sourceRow ? sourceRow[23] : "",
        course: sourceRow ? sourceRow[18] : "",
        contact: sourceRow ? sourceRow[22] : "",
        email: sourceRow ? sourceRow[24] : "",
        note: sourceRow ? sourceRow[25] : "",
        applyTime: sourceRow ? formatDate(sourceRow[26]) : ""
      },
      work: {
        lecturers: [workData[0], workData[1], workData[2]],
        guides: [workData[3], workData[4], workData[5]],
        facilitators: [workData[6], workData[7], workData[8]],
        promo: workData[9],
        checked: workData[10] === true,
        note: workData[11]
      }
    };
  } catch (e) {
    return { error: "讀取錯誤：" + e.message };
  }
}

// ===== 4. 儲存資料 =====
function saveOrderDataAPI(form) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);

    const orderId = form.orderId;
    const updates = [];

    updates.push({ columnIndex: 0, value: form.lecturers[0] });
    updates.push({ columnIndex: 1, value: form.lecturers[1] });
    updates.push({ columnIndex: 2, value: form.lecturers[2] });
    updates.push({ columnIndex: 3, value: form.guides[0] });
    updates.push({ columnIndex: 4, value: form.guides[1] });
    updates.push({ columnIndex: 5, value: form.guides[2] });
    updates.push({ columnIndex: 6, value: form.facilitators[0] });
    updates.push({ columnIndex: 7, value: form.facilitators[1] });
    updates.push({ columnIndex: 8, value: form.facilitators[2] });
    updates.push({ columnIndex: 9, value: form.promo });
    updates.push({ columnIndex: 10, value: form.checked });
    updates.push({ columnIndex: 11, value: form.note });

    batchWriteToWorkArrangement(orderId, updates);
    
    const newStats = getStaffStatistics();
    return { success: true, message: "✅ 儲存成功！", newStats: newStats };
  } catch (e) {
    throw new Error("儲存失敗：" + e.message);
  }
}

// ===== 5. 管理功能 API =====
function getManagementDataAPI() {
  const staff = getStaffConfigData(); 
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  let emailList = [];
  
  if (emailSheet) {
    const lastRow = emailSheet.getLastRow();
    if (lastRow >= 2) {
      emailList = emailSheet.getRange(2, 1, lastRow - 1, 2).getValues()
        .map(r => ({ email: r[0], name: r[1] }))
        .filter(r => r.email);
    }
  }
  return { staff: staff, emails: emailList };
}

function updateStaffAPI(action, type, name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.staffConfigSheetName);
  let colIndex = 0;
  
  if (type === 'lecturer') colIndex = 1;
  else if (type === 'guide') colIndex = 2;
  else if (type === 'facilitator') colIndex = 3;
  
  if (colIndex === 0) throw new Error("無效的人員類型");

  if (action === 'add') {
    const lastRow = sheet.getRange(1, colIndex, sheet.getMaxRows(), 1).getValues().filter(String).length;
    sheet.getRange(lastRow + 1, colIndex).setValue(name);
    return "新增成功";
  } 
  else if (action === 'delete') {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "無資料可刪除";

    const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const values = range.getValues().flat();
    const newValues = values.filter(v => v !== name && v !== "");
    
    range.clearContent();
    if (newValues.length > 0) {
      sheet.getRange(2, colIndex, newValues.length, 1).setValues(newValues.map(v => [v]));
    }
    return "刪除成功";
  }
}

function updateEmailAPI(action, email, name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  
  if (action === 'add') {
    sheet.appendRow([email, name]);
    return "新增成功";
  } 
  else if (action === 'delete') {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) { 
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        return "刪除成功";
      }
    }
    throw new Error("找不到該 Email");
  }
}

// ===== 6. 執行動作 =====
function runBackendActionAPI(action, orderId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  
  mainSheet.getRange("A4").setValue(orderId);
  SpreadsheetApp.flush();
  refreshCurrentOrder(); 

  try {
    if (action === 'notifyUser') return webSendApplicantNotification(); 
    else if (action === 'addToCalendar') return webAddToCalendar(orderId);
    else if (action === 'sendPDF') return webExportAndSendPDF(orderId);
    else if (action === 'sendPDF_Custom') return webExportAndSendPDF_Custom(orderId);
    return "未知的指令";
  } catch (e) {
    throw new Error("執行失敗：" + e.message);
  }
}

// ===== Web 版專用邏輯 =====

function webSendApplicantNotification() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const applicantEmail = spreadsheet.getRangeByName("申請人電子郵件").getValue();
  if (!applicantEmail) throw new Error("找不到申請人電子郵件");
  
  const subject = "檔案素養教育（開箱檔案館–到館學習）系統已受理（待審核，尚未完成報名），請靜候承辦人通知說明";
  const body = `您好！您的「開箱檔案館–到館學習」申請案已受理。案件仍在處理中，本局將依申請提出順序及各時段受理情形，安排並告知申請結果，感謝您的耐心等候。

黃宇暘敬上

國家發展委員會檔案管理局
應用服務組約用研究員　黃宇暘
TEL：(02)8995-3979
https://www.archives.gov.tw
244013　新北市林口區檔案館路1號`;
  
  GmailApp.sendEmail(applicantEmail, subject, body);
  return `✅ 受理通知已寄送給：${applicantEmail}`;
}

function webAddToCalendar(orderId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ranges = {
    date: spreadsheet.getRangeByName("參訪日期").getValue(),
    time: spreadsheet.getRangeByName("參訪時段").getValue(),
    unit: spreadsheet.getRangeByName("參訪單位").getValue(),
    plan: spreadsheet.getRangeByName("選擇方案").getValue()
  };
  
  if (!ranges.date || !ranges.time) throw new Error("日期或時間資料不全");
  
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId) || CalendarApp.getDefaultCalendar();
  const [startStr, endStr] = ranges.time.split("-");
  const date = new Date(ranges.date);
  
  const start = new Date(date);
  const [sH, sM] = startStr.split(":");
  start.setHours(sH, sM, 0);
  
  const end = new Date(date);
  const [eH, eM] = endStr.split(":");
  end.setHours(eH, eM, 0);
  
  const title = `${ranges.unit}－到館學習`;
  const desc = `參訪單位：${ranges.unit}\n` +
               `參訪時間：${ranges.time}\n` +
               `活動類型：檔案素養教育－到館學習\n` +
               `方案：${ranges.plan}\n` +
               `單號：${orderId}\n\n` +
               `※ 此活動由 Google Apps Script 自動建立`;
  
  const location = "國家發展委員會檔案管理局 244013 新北市林口區檔案館路1號";

  const existing = calendar.getEvents(start, end, {search: title});
  if (existing.length > 0) return "⚠️ 行事曆已有相同活動，已略過建立。";
  
  calendar.createEvent(title, start, end, {
    description: desc, 
    location: location
  });
  
  return "✅ 已加入 Google 行事曆！";
}

function webExportAndSendPDF(orderId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const orderName = mainSheet.getRange("A4").getValue();
  const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
  
  const visitDateVal = spreadsheet.getRangeByName("參訪日期").getValue();
  const visitTimeVal = spreadsheet.getRangeByName("參訪時段").getValue();
  const visitDateStr = visitDateVal instanceof Date ? Utilities.formatDate(visitDateVal, Session.getScriptTimeZone(), "yyyy/MM/dd") : visitDateVal;
  
  let scheduleSheetName = selectedPlan === "2小時方案" ? "（勿動）參訪程序表-2小時" : "（勿動）參訪程序表-3小時";
  const scheduleSheet = spreadsheet.getSheetByName(scheduleSheetName);
  if (!scheduleSheet) throw new Error(`找不到對應的程序表：${scheduleSheetName}`);
  
  scheduleSheet.showSheet();
  SpreadsheetApp.flush();
  
  let mainBlob, schBlob;
  try {
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    mainBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_工作配置單.pdf`);
    
    schBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}&range=A1:E13`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_參訪程序表.pdf`);
  } finally {
    scheduleSheet.hideSheet();
  }

  const recipient = "yyhuang@archives.gov.tw";
  const ccList = ["hcsun@archives.gov.tw", "yhwei@archives.gov.tw", "sihung@archives.gov.tw"]; 
  
  const workData = getWorkArrangementData(orderId); 
  const assignedNames = workData.slice(0, 9).filter(name => name && name !== "" && name !== "未定（無）");
  
  const emailSheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  let matchedEmails = [];
  
  if (emailSheet && assignedNames.length > 0) {
    const lastRow = emailSheet.getLastRow();
    if (lastRow >= 2) {
      const emailData = emailSheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
      assignedNames.forEach(name => {
        const found = emailData.find(row => row[1] === name);
        if (found && found[0]) matchedEmails.push(found[0]);
      });
    }
  }
  
  const finalCc = [...new Set([...ccList, ...matchedEmails])].join(",");

  const subject = `檔案素養教育－到館學習：工作配置及參訪程序提醒（${visitDateStr}，${visitTimeVal}）`;
  const body = `大家好，工作排程請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/WO73VZ\n\n(本郵件副本已自動發送給：${matchedEmails.join(", ") || "無相關人員"})`;

  GmailApp.sendEmail(recipient, subject, body, { 
    attachments: [mainBlob, schBlob], 
    cc: finalCc 
  });
  
  return `✅ PDF 已生成並寄出！\n副本已同步發送給 ${matchedEmails.length} 位工作人員。`;
}

function webExportAndSendPDF_Custom(orderId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const orderName = mainSheet.getRange("A4").getValue();
  const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
  
  const visitDateVal = spreadsheet.getRangeByName("參訪日期").getValue();
  const visitTimeVal = spreadsheet.getRangeByName("參訪時段").getValue();
  const visitDateStr = visitDateVal instanceof Date ? Utilities.formatDate(visitDateVal, Session.getScriptTimeZone(), "yyyy/MM/dd") : visitDateVal;
  
  let scheduleSheetName = "";
  let rangeParams = "";
  
  if (selectedPlan === "2小時方案") {
    scheduleSheetName = "參訪程序表-2小時（手動編修版）";
    rangeParams = "&range=A1:F12"; 
  } else {
    scheduleSheetName = "參訪程序表-3小時（手動編修版）";
    rangeParams = "&range=A1:E13";
  }
  
  const scheduleSheet = spreadsheet.getSheetByName(scheduleSheetName);
  if (!scheduleSheet) throw new Error(`找不到客製化程序表：${scheduleSheetName}`);
  
  scheduleSheet.showSheet();
  SpreadsheetApp.flush();
  
  let mainBlob, schBlob;
  try {
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    mainBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_工作配置單.pdf`);
    
    schBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}${rangeParams}`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_客製化參訪程序表.pdf`);
  } finally {
    scheduleSheet.hideSheet();
  }

  const recipient = "yyhuang@archives.gov.tw";
  const ccList = ["hcsun@archives.gov.tw", "yhwei@archives.gov.tw", "sihung@archives.gov.tw"]; 
  
  const workData = getWorkArrangementData(orderId); 
  const assignedNames = workData.slice(0, 9).filter(name => name && name !== "" && name !== "未定（無）");
  
  const emailSheet = spreadsheet.getSheetByName(CONFIG.emailSheetName);
  let matchedEmails = [];
  
  if (emailSheet && assignedNames.length > 0) {
    const lastRow = emailSheet.getLastRow();
    if (lastRow >= 2) {
      const emailData = emailSheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
      assignedNames.forEach(name => {
        const found = emailData.find(row => row[1] === name);
        if (found && found[0]) matchedEmails.push(found[0]);
      });
    }
  }
  
  const finalCc = [...new Set([...ccList, ...matchedEmails])].join(",");

  const subject = `檔案素養教育－到館學習：工作配置及參訪程序提醒（${visitDateStr}，${visitTimeVal}）- 客製化行程`;
  const body = `大家好，工作排程請參附件（包含客製化行程），如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/WO73VZ\n\n(本郵件副本已自動發送給：${matchedEmails.join(", ") || "無相關人員"})`;

  GmailApp.sendEmail(recipient, subject, body, { 
    attachments: [mainBlob, schBlob], 
    cc: finalCc 
  });
  
  return `✅ 客製化 PDF 已生成並寄出！\n副本已同步發送給 ${matchedEmails.length} 位工作人員。`;
}

// ===== 輔助函數 =====
function getStaffStatistics() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  const sheet = CACHE.targetSheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lecturer: [], guide: [], facilitator: [] };

  const data = sheet.getRange(2, 3, lastRow - 1, 9).getValues();
  const counts = { lecturer: {}, guide: {}, facilitator: {} };

  data.forEach(row => {
    [0, 1, 2].forEach(i => addCount(counts.lecturer, row[i]));
    [3, 4, 5].forEach(i => addCount(counts.guide, row[i]));
    [6, 7, 8].forEach(i => addCount(counts.facilitator, row[i]));
  });

  return {
    lecturer: formatStatsHTML(counts.lecturer),
    guide: formatStatsHTML(counts.guide),
    facilitator: formatStatsHTML(counts.facilitator)
  };
}

function addCount(obj, name) {
  if (!name || name === "" || name === "未定（無）") return;
  obj[name] = (obj[name] || 0) + 1;
}

function formatStatsHTML(obj) {
  return Object.entries(obj)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => `<li class="list-group-item d-flex justify-content-between align-items-center">${name} <span class="badge bg-primary rounded-pill">${count} 次</span></li>`)
    .join('');
}