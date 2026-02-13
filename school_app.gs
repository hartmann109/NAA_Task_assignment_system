/**
 * 檔案素養教育－到校服務管理系統 (School Service Management System)
 * Developer: 黃宇暘
 * Date: 2026/01/28 - Version 3.0 (Settings & Layout Fix)
 */

var SPREADSHEET_ID = '1SRacE8dHD-fAF6T67mPllPs0eACC9ldksMKHhI3056Q'; 
var SCHOOL_SHEET_NAME = '素養到校工作安排';
var PERSONNEL_CONFIG_SHEET = '人員配置表';      // 用於產生「指派人員下拉選單」 (D欄)
var EMAIL_LIST_SHEET = '電子郵件寄送清單';      // 用於「查找 Email」 (A欄Email, B欄Name)
var CALENDAR_ID = 'c_3fc4a74cc751b2cbe15fcd2b384d078d86fee64693db9f7c7132b20ef6a475fb@group.calendar.google.com';

// function doGet(e) {
//   return HtmlService.createTemplateFromFile('index_school')
//       .evaluate()
//       .setTitle('檔案素養教育－到校服務管理系統')
//       .addMetaTag('viewport', 'width=device-width, initial-scale=1');
// }

/* --- 主要資料讀取 --- */
function getSchoolData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SCHOOL_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  return data.filter(r => r[0] !== "").map(row => ({
      id: row[0],           
      date: row[1] instanceof Date ? Utilities.formatDate(row[1], "GMT+8", "yyyy/MM/dd") : row[1],
      timeSlot: row[2],     
      unit: row[3],         
      location: row[4],     
      applicant: row[5],    
      contact: row[6],      
      phone: row[7],        
      email: row[8],        
      course: row[9],       
      identity: row[10],    
      count: row[12],       
      content: row[13],     
      note: row[15],        
      submitTime: row[16] instanceof Date ? Utilities.formatDate(row[16], "GMT+8", "yyyy/MM/dd HH:mm:ss") : row[16],
      assignedTo: row[17],  
      isCancelled: row[18], 
      isCompleted: row[19]  
  }));
}

/* --- 設定功能：人員配置表 (下拉選單用) --- */
// 讀取 D 欄
function getSchoolPersonnelList() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat().filter(String);
}

// 新增到 D 欄
function addSchoolPerson(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  // 為了避免影響 ABC 欄，我們找出 D 欄最後一個有值的位置
  var colD = sheet.getRange("D:D").getValues();
  var lastRow = colD.filter(String).length; 
  sheet.getRange(lastRow + 1, 4).setValue(name);
  return "OK";
}

// 從 D 欄刪除
function removeSchoolPerson(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  var data = sheet.getRange("D:D").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == name) {
      sheet.getRange(i + 1, 4).clearContent(); // 清除內容
      // 若需要向上填補空缺，可以使用 deleteCells，但 clearContent 較安全
      // 這裡改用 deleteCells 確保列表連續
      sheet.getRange(i + 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS); 
      return "OK";
    }
  }
  return "Not Found";
}

/* --- 設定功能：Email 寄送清單 (後端寄信用) --- */
// 讀取 A(Email), B(Name)
function getEmailListData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  // 回傳 [{email:..., name:...}]
  return data.map(r => ({ email: r[0], name: r[1] })).filter(r => r.email && r.name);
}

// 新增 Email 人員
function addEmailPerson(name, email) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  sheet.appendRow([email, name]); // A:Email, B:Name
  return "OK";
}

// 刪除 Email 人員 (根據 Email)
function removeEmailPerson(email) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  var data = sheet.getRange("A:A").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == email) {
      sheet.deleteRow(i + 1);
      return "OK";
    }
  }
  return "Not Found";
}

/* --- 核心操作 --- */
function saveSchoolConfig(orderId, assignedTo, isCancelled, isCompleted) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SCHOOL_SHEET_NAME);
  var data = sheet.getRange("A:A").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == orderId) {
      var row = i + 1;
      sheet.getRange(row, 18).setValue(assignedTo);
      sheet.getRange(row, 19).setValue(isCancelled ? "TRUE" : "");
      sheet.getRange(row, 20).setValue(isCompleted ? "TRUE" : "");
      return { success: true, message: "配置已更新" };
    }
  }
  return { success: false, message: "找不到單號" };
}

function sendSchoolWorkNotification(info) {
  try {
    if (!info.assignedTo) return { success: false, message: "未指定人員" };
    
    // 跨表查詢 Email
    var assignedEmail = getEmailFromList(info.assignedTo);
    if (!assignedEmail) return { success: false, message: "在「電子郵件寄送清單」中找不到 [" + info.assignedTo + "] 的 Email，請至設定中新增。" };

    var subject = "檔案素養教育（檔案館外送－到校服務）工作通知－" + info.date + "，" + info.timeSlot;
    var body = 
      "您好，檔案館外送－到校服務工作安排如下：\n\n" +
      info.id + "－" + info.date + "－" + info.timeSlot + "\n" +
      "申請單位：" + info.unit + "\n" +
      "地點：" + info.location + "\n" +
      "申請人：" + info.applicant + "\n" +
      "課程名稱：" + info.course + "\n" +
      "聯絡人：" + info.contact + "\n" +
      "電話：" + info.phone + "\n" +
      "聯絡信箱：" + info.email + "\n" +
      "授課對象身分別：" + info.identity + "\n" +
      "課程人數：" + info.count + "\n" +
      "申請課程內容：" + info.content + "\n\n" +
      "如有問題，請詢問紹綱 #3992。活動已同步登錄至行事曆，如尚未訂閱，請點選：https://reurl.cc/7bzaM5\n\n" +
      "徐紹綱敬上\n\n" +
      "國家發展委員會檔案管理局\n應用服務組副研究員　 徐紹綱\nTEL：(02)8995-3992\nhttps://www.archives.gov.tw\n244013　新北市林口區檔案館路1號";

    var aliases = GmailApp.getAliases();
    var fromEmail = "";
    if (aliases.indexOf("skhsu@archives.gov.tw") > -1) fromEmail = "skhsu@archives.gov.tw";
    var options = { name: "國家發展委員會檔案管理局-徐紹綱" };
    if (fromEmail) options.from = fromEmail;

    GmailApp.sendEmail(assignedEmail, subject, body, options);
    return { success: true, message: "工作通知已寄送給：" + info.assignedTo + " (" + assignedEmail + ")" };
  } catch (e) { return { success: false, message: "寄送失敗：" + e.message }; }
}

function getEmailFromList(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == name) return data[i][0]; // Match Name(B), Return Email(A)
  }
  return ""; 
}

// 受理通知 (保持不變)
function sendSchoolAcceptanceEmail(info) {
  try {
    var subject = "檔案素養教育（檔案館外送－到校服務）系統已受理（待審核，尚未完成報名），請靜候承辦人通知說明";
    var body = "您好！您的「檔案館外送－到校服務」申請案已受理。案件仍在處理中，本局將依申請提出順序及各時段受理情形，安排並告知申請結果，感謝您的耐心等候。\n\n徐紹綱敬上\n\n國家發展委員會檔案管理局\n應用服務組副研究員　 徐紹綱\nTEL：(02)8995-3992\nhttps://www.archives.gov.tw\n244013　新北市林口區檔案館路1號";
    var aliases = GmailApp.getAliases();
    var fromEmail = "";
    if (aliases.indexOf("skhsu@archives.gov.tw") > -1) fromEmail = "skhsu@archives.gov.tw";
    var options = { name: "國家發展委員會檔案管理局-徐紹綱" };
    if (fromEmail) options.from = fromEmail;
    GmailApp.sendEmail(info.email, subject, body, options);
    return { success: true, message: "受理通知已寄出" };
  } catch (e) { return { success: false, message: "寄信失敗：" + e.message }; }
}

// 行事曆 (保持不變)
function addSchoolCalendarEvent(info) {
  try {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) return { success: false, message: "找不到指定的行事曆" };
    var dateStr = info.date.replace(/-/g, "/"); 
    var timeSlot = info.timeSlot;
    var startTime, endTime;
    if (timeSlot.indexOf("-") > -1) {
      var times = timeSlot.split("-");
      startTime = new Date(dateStr + " " + times[0].trim());
      endTime = new Date(dateStr + " " + times[1].trim());
    } else {
      startTime = new Date(dateStr + " 09:00");
      endTime = new Date(dateStr + " 17:00");
    }
    var title = info.unit + " - 檔案素養教育(到校)";
    var description = "申請單位：" + info.unit + "\n申請時間：" + info.submitTime + "\n活動類型：檔案素養教育－到校服務\n單號：" + info.id + "\n申請人：" + info.applicant + "\n聯絡電話：" + info.phone;
    calendar.createEvent(title, startTime, endTime, { description: description, location: info.location });
    return { success: true, message: "已加入行事曆" };
  } catch (e) { return { success: false, message: "行事曆加入失敗：" + e.message }; }
}