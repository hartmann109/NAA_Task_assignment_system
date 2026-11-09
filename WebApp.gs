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

// ===== 讀寫排程勾選名單 (存在系統屬性中，不干擾試算表) =====
function getLecturerScheduleAPI() {
  const props = PropertiesService.getDocumentProperties();
  const data = props.getProperty('LECTURER_SCHEDULE');
  return data ? JSON.parse(data) : [];
}

function toggleLecturerScheduleAPI(name, isScheduled) {
  const props = PropertiesService.getDocumentProperties();
  let list = getLecturerScheduleAPI();
  
  if (isScheduled) {
    if (!list.includes(name)) list.push(name);
  } else {
    list = list.filter(n => n !== name);
  }
  
  props.setProperty('LECTURER_SCHEDULE', JSON.stringify(list));
  return list;
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
    
    let unitStr = "";
    let timeStr = "";
    let isoDate = "";

    if (!isCancelled) {
      const sourceRow = getSourceDataByKey(id); 
      if (sourceRow) {
        const visitDate = sourceRow[15]; 
        const isChecked = checkedStatuses[index] === true;
        
        unitStr = sourceRow[17] || ""; 
        timeStr = sourceRow[16] || ""; 

        if (visitDate instanceof Date) {
          isoDate = Utilities.formatDate(visitDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (visitDate < today && isChecked) isClosed = true;
          else if (visitDate >= today && isChecked) isArranged = true;
        } else if (visitDate) {
          try {
            let d = new Date(visitDate);
            if (!isNaN(d.getTime())) isoDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } catch(e) {}
        }
      }
    }

    orders.push({ 
      id: idStr, 
      year: year, 
      isClosed: isClosed,
      isArranged: isArranged,
      isCancelled: isCancelled,
      unit: unitStr,
      time: timeStr,
      isoDate: isoDate
    });
  });

  const staffConfig = getStaffConfigData();
  const stats = getStaffStatistics();
  const lecturerSchedule = getLecturerScheduleAPI();

  return {
    orders: orders,
    staff: staffConfig,
    stats: stats,
    serverYear: new Date().getFullYear().toString(),
    lecturerSchedule: lecturerSchedule
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
        plan: sourceRow ? sourceRow[14] : "",
        // 新增：提取 AC 欄 (index 28) 作為來賓年齡層
        ageGroup: sourceRow ? (sourceRow[28] || "無") : "無"
      },
      work: {
        lecturers: [workData[0], workData[1], workData[2]],
        guides: [workData[3], workData[4], workData[5]],
        facilitators: [workData[6], workData[7], workData[8]],
        promo: workData[9],
        checked: workData[10] === true,
        note: workData[11],
        // 新增：讀取 AD 欄的 1小時方案年齡層配置
        ageLevelConfig: CACHE.targetSheet.getRange(row, 30).getValue()
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
    
    // 讀取 AD 欄的配置
    const ageLevelConfig = CACHE.targetSheet.getRange(row, 30).getValue();
    const rawPlan = sourceRow ? sourceRow[14] : "";
    let displayPlan = rawPlan;
    if (rawPlan === "1小時方案") {
        displayPlan = rawPlan + "／" + (ageLevelConfig || "低年級");
    }

    const basicInfo = {
      unit: sourceRow ? sourceRow[17] : "",
      applicantName: leaderName, 
      date: sourceRow ? formatDate(sourceRow[15]) : "",
      time: sourceRow ? sourceRow[16] : "", 
      count: sourceRow ? sourceRow[21] : "",
      plan: displayPlan 
    };

    let scheduleData = [];
    if (savedScheduleJson && typeof savedScheduleJson === 'string' && savedScheduleJson.length > 5) {
      try {
        scheduleData = JSON.parse(savedScheduleJson);
      } catch (e) {
        console.warn("JSON 解析失敗，重新產生預設行程");
        scheduleData = generateDefaultSchedule(basicInfo, staff, rawPlan, ageLevelConfig);
      }
    } else {
      scheduleData = generateDefaultSchedule(basicInfo, staff, rawPlan, ageLevelConfig);
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

function generateDefaultSchedule(info, staff, rawPlan, ageLevelConfig) {
  const plan = rawPlan || info.plan; 
  const timeStr = info.time || "10:00-12:00"; 
  const isMorning = timeStr.includes("08:") || timeStr.includes("09:") || timeStr.includes("10:") || timeStr.includes("11:") || timeStr.includes("12:");
  
  const txtFacilitator = staff.facilitators.join("、") || "未定";
  const txtLecturer = staff.lecturers.join("、") || "未定";
  const txtGuide = staff.guides.join("、") || "未定";

  let items = [];
  let currentBasis = new Date();

  // === 1小時方案邏輯 ===
  if (plan === "1小時方案") {
    currentBasis.setHours(isMorning ? 10 : 13, 50, 0);

    if (!ageLevelConfig || ageLevelConfig === "低年級") {
      items.push({ startTime: formatHM(currentBasis), duration: 10, content: "1樓檔案教室集合、影片欣賞", staff: txtFacilitator, location: "1樓檔案教室" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 10);
      items.push({ startTime: formatHM(currentBasis), duration: 30, content: "展廳介紹（檔案大廳、檔案FUN體驗、阿凱將小屋）", staff: txtGuide, location: "2樓展廳" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 30);
      items.push({ startTime: formatHM(currentBasis), duration: 10, content: "學習單闖關", staff: "", location: "2樓展廳" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 10);
      items.push({ startTime: formatHM(currentBasis), duration: 10, content: "回檔案教室完成學習單、合影\n*含學習單闖關約60分鐘", staff: txtFacilitator, location: "1樓檔案教室" });
    } else { // 中高年級
      items.push({ startTime: formatHM(currentBasis), duration: 10, content: "1樓檔案教室集合、影片欣賞", staff: txtFacilitator, location: "1樓檔案教室" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 10);
      items.push({ startTime: formatHM(currentBasis), duration: 30, content: "展廳介紹（檔案大廳、常設展廳、檔案FUN體驗）", staff: txtGuide, location: "2樓展廳" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 30);
      items.push({ startTime: formatHM(currentBasis), duration: 30, content: "元宇宙沉浸式體驗、發放學習單", staff: "", location: "2樓展廳" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 30);
      items.push({ startTime: formatHM(currentBasis), duration: 20, content: "學習單闖關", staff: "", location: "2樓展廳" });
      currentBasis.setMinutes(currentBasis.getMinutes() + 20);
      items.push({ startTime: formatHM(currentBasis), duration: 10, content: "回檔案教室完成學習單、合影\n*不含學習單闖關約60分鐘", staff: txtFacilitator, location: "1樓檔案教室" });
    }
    return items;
  }

  // === 2/3小時方案邏輯 ===
  let startHour = 10, startMin = 0;
  if (isMorning) {
    if (plan === "3小時方案") { startHour = 9; startMin = 0; }
    else { startHour = 10; startMin = 0; }
  } else {
    startHour = 14; startMin = 0;
  }
  
  currentBasis.setHours(startHour, startMin, 0);
  currentBasis.setMinutes(currentBasis.getMinutes() - 10);
  
  let duration = 10;
  items.push({
    startTime: formatHM(currentBasis),
    duration: duration,
    content: "1樓集合並引導至檔案教室",
    staff: txtFacilitator,
    location: "1樓檔案大廳"
  });
  
  currentBasis.setMinutes(currentBasis.getMinutes() + duration);

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

  duration = 10;
  items.push({
    startTime: formatHM(currentBasis), 
    duration: duration,
    content: "休息\n轉場上樓",
    staff: txtFacilitator,
    location: ""
  });

  currentBasis.setMinutes(currentBasis.getMinutes() + duration);

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
    duration = 30;
    items.push({
      startTime: formatHM(currentBasis), 
      duration: duration,
      content: "「重返1987─解嚴檔案」特展導覽",
      staff: txtGuide,
      location: "2樓特展廳"
    });
    
    currentBasis.setMinutes(currentBasis.getMinutes() + duration);

    items.push({
      startTime: formatHM(currentBasis), 
      duration: 0, 
      content: "合影\n結束參訪",
      staff: txtFacilitator,
      location: "2樓展場櫃檯"
    });
  } else {
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
    else if (action === 'sendPDF') return webExportAndSendPDF(orderId, isTestMode);
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
    
    // 寫入 1小時方案 的年齡層選項 (AD欄)
    if (form.ageLevelConfig !== undefined) {
       const row = findMatchingRow(orderId);
       if (row !== -1) {
          CACHE.targetSheet.getRange(row, 30).setValue(form.ageLevelConfig);
       }
    }
    
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
  return { 
    staff: staff, 
    emails: emailList, 
    lecturerSchedule: getLecturerScheduleAPI()
  };
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
    
    if (type === 'lecturer') toggleLecturerScheduleAPI(name, false);
    
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
應用服務組約用研究員 黃宇暘
TEL：(02)8995-3979
https://www.archives.gov.tw
244013 新北市林口區檔案館路1號`;
  
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

function webExportAndSendPDF(orderId, isTestMode) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const orderName = mainSheet.getRange("A4").getValue();
  const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue(); 

  const visitDateVal = spreadsheet.getRangeByName("參訪日期").getValue();
  const visitTimeVal = spreadsheet.getRangeByName("參訪時段").getValue();
  const visitDateStr = visitDateVal instanceof Date ? Utilities.formatDate(visitDateVal, Session.getScriptTimeZone(), "yyyy/MM/dd") : visitDateVal;

  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
  const row = findMatchingRow(orderId);
  const savedScheduleJson = CACHE.targetSheet.getRange(row, 28).getValue(); 
  const ageLevelConfig = CACHE.targetSheet.getRange(row, 30).getValue();

  let scheduleData = [];
  try {
    if (savedScheduleJson && savedScheduleJson.length > 5) {
      scheduleData = JSON.parse(savedScheduleJson);
    } else {
      const result = getScheduleEditorDataAPI(orderId); 
      scheduleData = result.schedule;
    }
  } catch(e) {
    throw new Error("行程資料解析失敗");
  }

  const genSheetName = "系統生成工作配置表";
  let genSheet = spreadsheet.getSheetByName(genSheetName);
  if (!genSheet) {
    genSheet = spreadsheet.insertSheet(genSheetName);
  }
  
  let applicantInfo = spreadsheet.getRange("B8").getValue() || "";
  let unitInfo = spreadsheet.getRangeByName("參訪單位").getValue();
  let rawCount = spreadsheet.getRangeByName("參訪人數").getValue();
  let countStr = rawCount.toString().includes("人") ? rawCount : rawCount + "人";

  let timeSuffix = "";
  try {
     const startHourStr = visitTimeVal.toString().split(":")[0];
     const startHour = parseInt(startHourStr, 10);
     if (!isNaN(startHour)) {
        timeSuffix = startHour < 12 ? " （上午）" : " （下午）";
     }
  } catch(e) {}

  let displayPlan = selectedPlan;
  if (selectedPlan === "1小時方案") {
      displayPlan = selectedPlan + "／" + (ageLevelConfig || "低年級");
  }

  const info = {
    unit: unitInfo + (applicantInfo ? "／" + applicantInfo : ""), 
    timeStr: visitDateStr + " " + visitTimeVal + timeSuffix, 
    count: countStr,
    plan: displayPlan 
  };
  
  writeToGeneratedSheet(genSheet, info, scheduleData);
  SpreadsheetApp.flush();

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

function writeToGeneratedSheet(sheet, info, scheduleData) {
  sheet.clear(); 
  
  sheet.setColumnWidth(1, 120); 
  sheet.setColumnWidth(2, 350); 
  sheet.setColumnWidth(3, 130); 
  sheet.setColumnWidth(4, 130); 

  const titleRange = sheet.getRange("A1:D1");
  titleRange.merge().setValue("國家發展委員會檔案管理局\n開箱檔案館–到館學習\n程序表");
  titleRange.setHorizontalAlignment("center").setVerticalAlignment("middle")
            .setFontWeight("bold").setFontSize(18).setWrap(true);
  sheet.setRowHeight(1, 100);

  sheet.getRange("A2").setValue("申請單位:"); 
  sheet.getRange("B2:D2").merge().setValue(info.unit);
  
  sheet.getRange("A3").setValue("參訪時間:"); 
  sheet.getRange("B3:D3").merge().setValue(info.timeStr);
  
  sheet.getRange("A4").setValue("參訪方案:"); 
  sheet.getRange("B4:D4").merge().setValue(info.plan);

  sheet.getRange("A5").setValue("到館人數:"); 
  sheet.getRange("B5:D5").merge().setValue(info.count);
  
  sheet.getRange("A6").setValue("流程:");
  sheet.getRange("A6:D6").merge().setValue("流程:"); 
  
  sheet.getRange("A2:A6").setFontWeight("bold").setFontSize(12);
  sheet.getRange("B2:D5").setFontSize(12);
  
  const headers = ["時間", "內容", "安排人員", "地點"];
  const headerRange = sheet.getRange("A8:D8");
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold").setFontSize(12).setBackground("#efefef")
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(8, 35);

  let startRow = 9; 
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

// ===== 工作通知 Email 寄送邏輯 =====
function sendWorkEmail(orderId, isTestMode, visitDateStr, visitTimeVal, mainBlob, schBlob) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  let recipient, ccListStr, subjectPrefix = "";
  
  if (isTestMode) {
    recipient = "htm109@gmail.com";
    ccListStr = ""; 
    subjectPrefix = "（測試信，請勿理會）";
  } else {
    // 正式模式：寄給承辦人 (宇暘)
    recipient = "yyhuang@archives.gov.tw";
    
    // 預設必發副本長官清單 (還原 0729 版本：加入 ychung)
    const ccList = ["hcsun@archives.gov.tw", "yhwei@archives.gov.tw", "sihung@archives.gov.tw", "ychung@archives.gov.tw"];
    
    // 讀取該單號被指派的工作人員
    const workData = getWorkArrangementData(orderId);
    if (workData) {
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

  let options = { attachments: [mainBlob, schBlob] };
  
  try {
      const aliases = GmailApp.getAliases();
      if (aliases.indexOf("yyhuang@archives.gov.tw") > -1) {
          options.from = "yyhuang@archives.gov.tw";
          options.name = "國家發展委員會檔案管理局-黃宇暘";
      }
  } catch(e) {
      console.log("無法獲取或設定 Email 別名：" + e);
  }

  if (ccListStr) options.cc = ccListStr;
  
  GmailApp.sendEmail(recipient, subject, body, options);
  
  if (isTestMode) {
      return `✅ [測試模式] PDF 已生成並寄出！\n只寄送給：${recipient}`;
  } else {
      return `✅ PDF 已生成並寄出！\n副本已同步發送給相關工作人員。`;
  }
}

// ===== 匯出並寄送 PDF（手動編修版） =====
// 完整還原 0729 版本中的手動編修發送功能
function exportAndSendPDFmanual() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

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
    const orderName = mainSheet.getRange("A4").getValue() || "未命名";
    const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
    
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
    } else if (selectedPlan === "1小時方案") {
      // 預留 1小時方案的手動版
      scheduleSheet = spreadsheet.getSheetByName("參訪程序表-1小時（手動編修版）");
      scheduleRange = "A1:F13"; 
    } else {
      SpreadsheetApp.getUi().alert("錯誤", `無法識別的選擇方案：${selectedPlan}。`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (!scheduleSheet) {
      SpreadsheetApp.getUi().alert("錯誤", `找不到工作表：參訪程序表-${selectedPlan.substring(0,2)}（手動編修版）\n\n請確認該工作表已建立。`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (scheduleSheet.isSheetHidden()) {
      scheduleSheetWasHidden = true;
      scheduleSheet.showSheet();
      SpreadsheetApp.flush();
    }
    
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    const mainSheetUrl = `${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`;
    const mainSheetBlob = fetchPDFWithRetry(mainSheetUrl, token).getBlob().setName(`${orderName}_工作配置單.pdf`);
    
    const scheduleSheetUrl = `${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}&range=${scheduleRange}`;
    const scheduleSheetBlob = fetchPDFWithRetry(scheduleSheetUrl, token).getBlob().setName(`${orderName}_參訪程序表_${selectedPlan}（手動編修版）.pdf`);
    
    const namesInSheet = mainSheet.getRange("B14:D16").getValues().flat().filter(n => n);
    const emailData = emailSheet.getRange("A2:B").getValues();
    
    const matchedEmails = emailData
      .filter(row => row[1] && namesInSheet.includes(row[1]))
      .map(row => row[0])
      .filter(email => email);
    
    const mainRecipient = "yyhuang@archives.gov.tw";
    // 完整還原 0729 版本中的手動編修版長官名單 (包含 mjjheng)
    const defaultCcEmails = [
      "hcsun@archives.gov.tw",
      "yhwei@archives.gov.tw", 
      "sihung@archives.gov.tw",
      "mjjheng@archives.gov.tw"
    ];
    
    const allCcEmails = [...new Set([...matchedEmails, ...defaultCcEmails])].filter(email => email);
    
    GmailApp.sendEmail(mainRecipient, emailSubject, 
      "請參附件，如有問題，請詢問宇暘 #3979。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/WO73VZ\n\n※ 此版本使用手動編修版參訪程序表", 
      {
        attachments: [mainSheetBlob, scheduleSheetBlob],
        cc: allCcEmails.join(",")
      }
    );
    
    SpreadsheetApp.getUi().alert("成功", `工作配置提醒PDF（手動編修版）已生成並寄送完成！`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error("PDF 匯出或寄送過程中發生錯誤:", error);
    SpreadsheetApp.getUi().alert("錯誤", `處理過程中發生錯誤：${error.message}\n\n請稍後再試，或聯繫系統管理員。`, SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    if (scheduleSheetWasHidden && scheduleSheet) {
      scheduleSheet.hideSheet();
    }
    SpreadsheetApp.flush();
  }
}

// ===== 9. 同步與排序外部表單資料 =====
function syncAndSortOrdersAPI() {
  try {
    const sourceUrl = 'https://docs.google.com/spreadsheets/d/1afDgJHQxhmcqFIEzN1ZNdXdSVyTyZYOBE8A6MioD0Wg/edit';
    const sourceSS = SpreadsheetApp.openByUrl(sourceUrl);
    const sourceSheet = sourceSS.getSheetByName('museumsheet');
    if (!sourceSheet) throw new Error("找不到來源工作表 (museumsheet)");

    const sourceLastRow = sourceSheet.getLastRow();
    if(sourceLastRow < 3) return { success: true, message: "來源表單目前沒有資料。" };
    const sourceData = sourceSheet.getRange('A3:O' + sourceLastRow).getValues();

    const targetSS = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = targetSS.getSheetByName(CONFIG.targetSheetName); 
    
    const targetLastRow = targetSheet.getLastRow();
    let existingIds = [];
    if (targetLastRow > 1) {
      existingIds = targetSheet.getRange('A2:A' + targetLastRow).getValues().flat().filter(String);
    }

    let addedCount = 0;
    const newRows = [];

    for(let i = 0; i < sourceData.length; i++) {
       const row = sourceData[i];
       const orderId = row[0]; 
       if(!orderId) continue;

       if(!existingIds.includes(orderId)) {
          let newRow = new Array(30).fill(""); // 擴充至 AD 欄 (Index 29)
          newRow[0] = orderId; 
          for(let j = 1; j <= 13; j++) {
             newRow[13 + j] = row[j]; // B~N 對應至 O~AA
          }
          newRow[28] = row[14] || ""; // O欄 (museum_studentage) 對應至 AC 欄
          newRows.push(newRow);
          addedCount++;
       }
    }

    if(newRows.length > 0) {
       targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    const finalLastRow = targetSheet.getLastRow();
    const finalLastCol = targetSheet.getLastColumn(); 
    if (finalLastRow > 1) {
        targetSheet.getRange(2, 1, finalLastRow - 1, finalLastCol).sort({column: 1, ascending: true});
    }

    CACHE.orders = null;
    CACHE.targetSheet = null;

    return { success: true, message: `同步完成！共匯入 ${addedCount} 筆新預約，並已自動排序。` };
  } catch (e) {
    return { success: false, message: `同步失敗：${e.message}` };
  }
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

function FORCE_AUTH_TRIGGER() {
  CalendarApp.getDefaultCalendar();
  GmailApp.getAliases();
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  console.log("✅ 授權成功！現在權限已更新。");
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

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const aboutText = `工作配置管理系統 v2.0\n開發者：國發會檔案管理局`;
  ui.alert('關於工作配置管理系統', aboutText, ui.ButtonSet.OK);
}