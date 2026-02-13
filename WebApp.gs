// ===== 1. 網頁入口與路由 =====
function doGet(e) {
  const page = e.parameter.page;
  if (page === 'view') {
    return HtmlService.createTemplateFromFile('index_view')
        .evaluate().setTitle('檔案素養教育工作配置系統 (檢視模式)')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'schedule') {
    return HtmlService.createTemplateFromFile('schedule')
        .evaluate().setTitle('參訪行程表編輯器')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'school') {
    return HtmlService.createTemplateFromFile('index_school')
        .evaluate().setTitle('檔案素養教育－到校服務管理系統')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createTemplateFromFile('index')
        .evaluate().setTitle('檔案素養教育工作配置系統')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function getCurrentSelectedOrderAPI() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  return mainSheet.getRange("A4").getValue();
}

// ===== 2. 取得初始資料 (首頁用) =====
function getInitialDataAPI() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  const targetSheet = CACHE.targetSheet;
  const lastRow = targetSheet.getLastRow();
  
  const orderIds = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const cancelStatuses = targetSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat(); 
  const checkedStatuses = targetSheet.getRange(2, 13, lastRow - 1, 1).getValues().flat();
  
  const orders = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  orderIds.forEach((id, index) => {
    if (!id) return;
    let year = "未知年份";
    const idStr = String(id);
    if (idStr.startsWith("MS") && idStr.length >= 6) year = idStr.substring(2, 6);

    let isClosed = false;
    let isArranged = false;
    const isCancelled = cancelStatuses[index] === true || cancelStatuses[index] === "TRUE"; 
    
    if (!isCancelled) {
      const sourceRow = getSourceDataByKey(id); 
      if (sourceRow) {
        const visitDate = sourceRow[15];
        const isChecked = checkedStatuses[index] === true;
        
        if (visitDate instanceof Date) {
          if (visitDate < today && isChecked) {
            isClosed = true;
          } else if (visitDate >= today && isChecked) {
            isArranged = true;
          }
        }
      }
    }

    orders.push({ 
      id: idStr, 
      year: year, 
      isClosed: isClosed,
      isArranged: isArranged,
      isCancelled: isCancelled
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

// ===== 3. 讀取單筆詳細資料 (首頁用) =====
function getOrderDataAPI(orderId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    
    const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
    mainSheet.getRange("A4").setValue(orderId);

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

    const isCancelled = CACHE.targetSheet.getRange(row, 2).getValue() === true;
    const workData = CACHE.targetSheet.getRange(row, 3, 1, 12).getValues()[0];
    const sourceRow = getSourceDataByKey(orderId);

    return {
      isCancelled: isCancelled,
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
        applyTime: sourceRow ? formatDate(sourceRow[26]) : "",
        plan: sourceRow ? sourceRow[14] : "" 
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

// ===== 4. 行程編輯器專用 API (Schedule Editor) =====

function getScheduleEditorDataAPI(orderId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    
    const row = findMatchingRow(orderId);
    if (row === -1) throw new Error("找不到單號");
    
    const sourceRow = getSourceDataByKey(orderId);
    const workData = CACHE.targetSheet.getRange(`C${row}:N${row}`).getValues()[0];
    
    const savedScheduleJson = CACHE.targetSheet.getRange(row, 28).getValue();
    
    const staff = {
      lecturers: [workData[0], workData[1], workData[2]].filter(n => n && n !== "未定（無）"),
      guides: [workData[3], workData[4], workData[5]].filter(n => n && n !== "未定（無）"),
      facilitators: [workData[6], workData[7], workData[8]].filter(n => n && n !== "未定（無）")
    };

    const leaderName = sourceRow ? (sourceRow[19] || "") + (sourceRow[20] ? " " + sourceRow[20] : "") : "";

    const basicInfo = {
      unit: sourceRow ? sourceRow[17] : "",
      applicantName: leaderName, 
      date: sourceRow ? formatDate(sourceRow[15]) : "",
      time: sourceRow ? sourceRow[16] : "", 
      count: sourceRow ? sourceRow[21] : "",
      plan: sourceRow ? sourceRow[14] : "" 
    };

    let scheduleData = [];
    if (savedScheduleJson && typeof savedScheduleJson === 'string' && savedScheduleJson.length > 5) {
      try {
        scheduleData = JSON.parse(savedScheduleJson);
      } catch (e) {
        console.warn("JSON 解析失敗，重新產生預設行程");
        scheduleData = generateDefaultSchedule(basicInfo, staff);
      }
    } else {
      scheduleData = generateDefaultSchedule(basicInfo, staff);
    }

    return {
      success: true,
      info: basicInfo,
      staff: staff,
      schedule: scheduleData
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveScheduleEditorDataAPI(orderId, scheduleData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    
    const row = findMatchingRow(orderId);
    if (row === -1) throw new Error("找不到單號");

    const jsonString = JSON.stringify(scheduleData);
    CACHE.targetSheet.getRange(row, 28).setValue(jsonString);

    return { success: true, message: "行程表已儲存" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function generateDefaultSchedule(info, staff) {
  const plan = info.plan; 
  const timeStr = info.time || "10:00-12:00"; 
  const isMorning = timeStr.includes("09:") || timeStr.includes("10:") || timeStr.includes("11:") || timeStr.includes("12:");
  
  let startHour = 10, startMin = 0;
  if (isMorning) {
    if (plan === "3小時方案") { startHour = 9; startMin = 0; }
    else { startHour = 10; startMin = 0; }
  } else {
    startHour = 14; startMin = 0;
  }

  const txtFacilitator = staff.facilitators.join("、") || "未定";
  const txtLecturer = staff.lecturers.join("、") || "未定";
  const txtGuide = staff.guides.join("、") || "未定";

  let items = [];

  let currentBasis = new Date();
  currentBasis.setHours(startHour, startMin, 0);
  currentBasis.setMinutes(currentBasis.getMinutes() - 10);
  
  // 1. 集合
  let duration = 10;
  items.push({
    startTime: formatHM(currentBasis),
    duration: duration,
    content: "1樓集合並引導至檔案教室",
    staff: txtFacilitator,
    location: "1樓檔案大廳"
  });
  
  // 推算下一項
  currentBasis.setMinutes(currentBasis.getMinutes() + duration);

  // 2. 解說
  const lectureDuration = (plan === "3小時方案") ? 90 : 60;
  const lectureContent = (plan === "3小時方案") 
    ? "「認識國家檔案」解說（檔案數位資源及檢索應用、特定主題檔案介紹）／政治檔案探究與應用、檔案小論文寫作入門"
    : "「認識國家檔案」解說（檔案數位資源及檢索應用、特定主題檔案介紹）";
  
  items.push({
    startTime: formatHM(currentBasis), 
    duration: lectureDuration,
    content: lectureContent,
    staff: txtLecturer,
    location: "1樓檔案教室"
  });

  currentBasis.setMinutes(currentBasis.getMinutes() + lectureDuration);

  // 3. 休息
  duration = 10;
  items.push({
    startTime: formatHM(currentBasis), 
    duration: duration,
    content: "休息\n轉場上樓",
    staff: txtFacilitator,
    location: ""
  });

  currentBasis.setMinutes(currentBasis.getMinutes() + duration);

  // 4. 常設展
  duration = 50;
  items.push({
    startTime: formatHM(currentBasis), 
    duration: duration,
    content: "「島讀•我們的故事」常設展導覽",
    staff: txtGuide,
    location: "2樓常設展廳"
  });

  currentBasis.setMinutes(currentBasis.getMinutes() + duration);

  if (plan === "3小時方案") {
    // 特展
    duration = 30;
    items.push({
      startTime: formatHM(currentBasis), 
      duration: duration,
      content: "「重返1987─解嚴檔案」特展導覽",
      staff: txtGuide,
      location: "2樓特展廳"
    });
    
    currentBasis.setMinutes(currentBasis.getMinutes() + duration);

    // 合影
    items.push({
      startTime: formatHM(currentBasis), 
      duration: 0, 
      content: "合影\n結束參訪",
      staff: txtFacilitator,
      location: "2樓展場櫃檯"
    });
  } else {
    // 2小時方案
    items.push({
      startTime: formatHM(currentBasis), 
      duration: 0,
      content: "合影\n結束參訪，回檔案教室拿包包",
      staff: txtFacilitator,
      location: "2樓展場櫃檯"
    });
  }
  return items;
}

function formatHM(date) {
  let h = date.getHours().toString().padStart(2, '0');
  let m = date.getMinutes().toString().padStart(2, '0');
  return `${h}:${m}`;
}

// ===== 5. 執行動作 (Main Actions) =====
function runBackendActionAPI(action, orderId, isTestMode) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  
  mainSheet.getRange("A4").setValue(orderId);
  SpreadsheetApp.flush();
  refreshCurrentOrder();

  try {
    if (action === 'notifyUser') return webSendApplicantNotification(isTestMode); 
    else if (action === 'addToCalendar') return webAddToCalendar(orderId); 
    else if (action === 'sendPDF') return webExportAndSendPDF(orderId, isTestMode); // 統一入口
    return "未知的指令";
  } catch (e) {
    throw new Error("執行失敗：" + e.message);
  }
}

// ===== 6. 其他儲存與取消 API =====
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

function setOrderCancellationAPI(orderId, isCancelled) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    
    const row = findMatchingRow(orderId);
    if (row === -1) throw new Error("找不到單號");

    CACHE.targetSheet.getRange(row, 2).setValue(isCancelled);

    return { 
      success: true, 
      message: isCancelled ? "已標記為【取消】" : "已恢復為【正常預約】",
      isCancelled: isCancelled
    };
  } catch (e) {
    throw new Error("更新狀態失敗：" + e.message);
  }
}

// ===== 7. 管理功能 API (Staff & Email) =====
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
    const values = range.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
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

// ===== 8. Web版 功能邏輯 =====

function webSendApplicantNotification(isTestMode) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let recipient;
  let subjectPrefix = "";
  
  if (isTestMode) {
    recipient = "htm109@gmail.com";
    subjectPrefix = "（測試信，請勿理會）";
  } else {
    recipient = spreadsheet.getRangeByName("申請人電子郵件").getValue();
    if (!recipient) throw new Error("找不到申請人電子郵件");
  }

  const subject = `${subjectPrefix}檔案素養教育（開箱檔案館–到館學習）系統已受理（待審核，尚未完成報名），請靜候承辦人通知說明`;
  const body = `您好！您的「開箱檔案館–到館學習」申請案已受理。案件仍在處理中，本局將依申請提出順序及各時段受理情形，安排並告知申請結果，感謝您的耐心等候。

黃宇暘敬上

國家發展委員會檔案管理局
應用服務組約用研究員　黃宇暘
TEL：(02)8995-3979
https://www.archives.gov.tw
244013　新北市林口區檔案館路1號`;
  
  GmailApp.sendEmail(recipient, subject, body);
  return isTestMode ? `✅ [測試模式] 受理通知已寄送給：${recipient}` : `✅ 受理通知已寄送給：${recipient}`;
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
  const desc = `參訪單位：${ranges.unit}\n參訪時間：${ranges.time}\n活動類型：檔案素養教育－到館學習\n方案：${ranges.plan}\n單號：${orderId}\n\n※ 此活動由 Google Apps Script 自動建立`;
  const location = "國家發展委員會檔案管理局 244013 新北市林口區檔案館路1號";

  const existing = calendar.getEvents(start, end, {search: title});
  if (existing.length > 0) return "⚠️ 行事曆已有相同活動，已略過建立。";
  calendar.createEvent(title, start, end, { description: desc, location: location });
  return "✅ 已加入 Google 行事曆！";
}

// 統一 PDF 匯出函式 (一律走系統生成表)
function webExportAndSendPDF(orderId, isTestMode) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const orderName = mainSheet.getRange("A4").getValue();
  const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();

  const visitDateVal = spreadsheet.getRangeByName("參訪日期").getValue();
  const visitTimeVal = spreadsheet.getRangeByName("參訪時段").getValue();
  const visitDateStr = visitDateVal instanceof Date ? Utilities.formatDate(visitDateVal, Session.getScriptTimeZone(), "yyyy/MM/dd") : visitDateVal;

  // 1. 取得 JSON 資料 (若無則自動產生預設)
  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  const row = findMatchingRow(orderId);
  const savedScheduleJson = CACHE.targetSheet.getRange(row, 28).getValue(); // AB欄

  let scheduleData = [];
  try {
    if (savedScheduleJson && savedScheduleJson.length > 5) {
      scheduleData = JSON.parse(savedScheduleJson);
    } else {
      // 若無客製化資料，使用預設邏輯產生
      const result = getScheduleEditorDataAPI(orderId); 
      scheduleData = result.schedule;
    }
  } catch(e) {
    throw new Error("行程資料解析失敗");
  }

  // 2. 將資料寫入「系統生成工作配置表」
  const genSheetName = "系統生成工作配置表";
  let genSheet = spreadsheet.getSheetByName(genSheetName);
  if (!genSheet) {
    genSheet = spreadsheet.insertSheet(genSheetName);
  }
  
  // 準備基本資訊
  let applicantInfo = spreadsheet.getRange("B8").getValue() || "";
  let unitInfo = spreadsheet.getRangeByName("參訪單位").getValue();
  let rawCount = spreadsheet.getRangeByName("參訪人數").getValue();
  let countStr = rawCount.toString().includes("人") ? rawCount : rawCount + "人";

  // 計算時間後綴 (上午/下午)
  let timeSuffix = "";
  try {
     const startHourStr = visitTimeVal.toString().split(":")[0];
     const startHour = parseInt(startHourStr, 10);
     if (!isNaN(startHour)) {
        timeSuffix = startHour < 12 ? " （上午）" : " （下午）";
     }
  } catch(e) {
     console.log("解析時間錯誤: " + e);
  }

  const info = {
    unit: unitInfo + (applicantInfo ? "／" + applicantInfo : ""), 
    timeStr: visitDateStr + " " + visitTimeVal + timeSuffix, // 加入後綴
    count: countStr,
    plan: selectedPlan // 新增方案
  };
  
  // 執行寫入 (設定滿版 A4 參數)
  writeToGeneratedSheet(genSheet, info, scheduleData);
  SpreadsheetApp.flush();

  // 3. 匯出 PDF
  let mainBlob, schBlob;
  try {
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    mainBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_工作配置單.pdf`);
    
    const lastRow = genSheet.getLastRow();
    const generatedPdfParams = pdfParams + "&fitw=true"; 
    schBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${genSheet.getSheetId()}${generatedPdfParams}&range=A1:D${lastRow+1}`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_參訪程序表.pdf`);
    
  } catch (e) {
    throw new Error("PDF 生成失敗：" + e.message);
  }

  return sendWorkEmail(orderId, isTestMode, visitDateStr, visitTimeVal, mainBlob, schBlob);
}

// 輔助：寫入系統生成表 (優化 A4 滿版排版)
function writeToGeneratedSheet(sheet, info, scheduleData) {
  sheet.clear(); 
  
  // 設定欄寬
  sheet.setColumnWidth(1, 120); 
  sheet.setColumnWidth(2, 350); 
  sheet.setColumnWidth(3, 130); 
  sheet.setColumnWidth(4, 130); 

  // 標題與資訊
  const titleRange = sheet.getRange("A1:D1");
  titleRange.merge().setValue("國家發展委員會檔案管理局\n開箱檔案館–到館學習\n程序表");
  titleRange.setHorizontalAlignment("center").setVerticalAlignment("middle")
            .setFontWeight("bold").setFontSize(18).setWrap(true);
  sheet.setRowHeight(1, 100);

  // 基本資訊 (字體 12)
  // 修改：將 "流程：" 也放入表格結構中，確保對齊
  
  sheet.getRange("A2").setValue("申請單位:"); 
  sheet.getRange("B2:D2").merge().setValue(info.unit);
  
  sheet.getRange("A3").setValue("參訪時間:"); 
  sheet.getRange("B3:D3").merge().setValue(info.timeStr);
  
  sheet.getRange("A4").setValue("參訪方案:"); 
  sheet.getRange("B4:D4").merge().setValue(info.plan);

  sheet.getRange("A5").setValue("到館人數:"); 
  sheet.getRange("B5:D5").merge().setValue(info.count);
  
  sheet.getRange("A6").setValue("流程:");
  // A6:D6 merge會導致靠左對齊，保持不merge或合併後設定靠左
  sheet.getRange("A6:D6").merge().setValue("流程:");
  
  sheet.getRange("A2:A6").setFontWeight("bold").setFontSize(12);
  sheet.getRange("B2:D5").setFontSize(12);
  
  // 表頭 (行號順移到 8)
  const headers = ["時間", "內容", "安排人員", "地點"];
  const headerRange = sheet.getRange("A8:D8");
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold").setFontSize(12).setBackground("#efefef")
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(8, 35);

  // 內容
  let startRow = 9; // 從第 9 行開始
  scheduleData.forEach((item, index) => {
    let timeStr = "";
    if (item.duration > 0) {
      let nextStart = addMinutesToTime(item.startTime, item.duration);
      timeStr = `${item.startTime}-${nextStart}\n(${item.duration}分鐘)`;
    } else {
      timeStr = item.startTime;
    }

    const rowData = [
      timeStr,
      item.content,
      item.staff,
      item.location
    ];
    
    const currentRow = startRow + index;
    const rowRange = sheet.getRange(currentRow, 1, 1, 4);
    rowRange.setValues([rowData]);
    rowRange.setFontSize(12).setVerticalAlignment("middle").setWrap(true);
    
    // 設定列高為 65
    sheet.setRowHeight(currentRow, 65);
    
    sheet.getRange(currentRow, 1).setHorizontalAlignment("center");
    sheet.getRange(currentRow, 3, 1, 2).setHorizontalAlignment("center");
  });

  sheet.getRange(8, 1, scheduleData.length + 1, 4).setBorder(true, true, true, true, true, true);
}

function addMinutesToTime(timeStr, mins) {
  if (!timeStr) return "";
  const [h, m] = timeStr.split(':').map(Number);
  const d = new Date();
  d.setHours(h, m + mins, 0);
  return `${d.getHours().toString().padStart(2,'0')}:${d.getMinutes().toString().padStart(2,'0')}`;
}

// 補回遺失的 sendWorkEmail 函式
function sendWorkEmail(orderId, isTestMode, visitDateStr, visitTimeVal, mainBlob, schBlob) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  let recipient, ccListStr, subjectPrefix = "";
  
  if (isTestMode) {
    recipient = "htm109@gmail.com";
    ccListStr = ""; 
    subjectPrefix = "（測試信，請勿理會）";
  } else {
    recipient = "yyhuang@archives.gov.tw";
    const ccList = ["hcsun@archives.gov.tw", "yhwei@archives.gov.tw", "sihung@archives.gov.tw"];
    
    const workData = getWorkArrangementData(orderId);
    if(workData) {
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
      ccListStr = [...new Set([...ccList, ...matchedEmails])].join(",");
    } else {
      ccListStr = ccList.join(",");
    }
  }

  const subject = `${subjectPrefix}檔案素養教育－到館學習：工作配置及參訪程序提醒（${visitDateStr}，${visitTimeVal}）`;
  
  let body;
  const commonLink = "https://reurl.cc/WO73VZ";
  if (isTestMode) {
      body = `【這是測試信】\n\n大家好，工作排程請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：${commonLink}`;
  } else {
      body = `大家好，工作排程請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：${commonLink}\n\n(本郵件副本已自動發送給相關工作人員)`;
  }

  GmailApp.sendEmail(recipient, subject, body, { 
    attachments: [mainBlob, schBlob], 
    cc: ccListStr 
  });
  
  if(isTestMode) {
      return `✅ [測試模式] PDF 已生成並寄出！\n只寄送給：${recipient}`;
  } else {
      return `✅ PDF 已生成並寄出！\n副本已同步發送給相關工作人員。`;
  }
}

// ===== 輔助函數 (維持不變) =====
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

function FORCE_AUTH_TRIGGER() {
  // 這段程式碼只是為了觸發授權視窗，不需要真的執行什麼
  CalendarApp.getDefaultCalendar();
  GmailApp.getRemainingDailyQuota();
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  console.log("✅ 授權成功！現在權限已更新。");
}