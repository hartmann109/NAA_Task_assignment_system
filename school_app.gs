/**
 * 檔案素養教育－到校服務管理系統 (School Service Management System)
 * Developer: 黃宇暘
 * Date: 2026/03/19 - Version (實體同步與排序版)
 */

var SPREADSHEET_ID = '1SRacE8dHD-fAF6T67mPllPs0eACC9ldksMKHhI3056Q'; 
var SCHOOL_SHEET_NAME = '素養到校工作安排';
var PERSONNEL_CONFIG_SHEET = '人員配置表';      
var EMAIL_LIST_SHEET = '電子郵件寄送清單';      
var CALENDAR_ID = 'c_3fc4a74cc751b2cbe15fcd2b384d078d86fee64693db9f7c7132b20ef6a475fb@group.calendar.google.com';

/* ========================================================================= */
/* --- 新增：同步表單資料與實體排序 API --- */
/* ========================================================================= */
function syncSchoolDataAPI() {
  try {
    // 原始預約表單的 ID 與分頁名稱
    var SOURCE_SPREADSHEET_ID = '1afDgJHQxhmcqFIEzN1ZNdXdSVyTyZYOBE8A6MioD0Wg';
    var SOURCE_SHEET_NAME = 'schoolsheet';
    
    // 1. 取得來源資料 (A 到 Q 欄，共 17 欄)
    var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    var sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
    var sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow < 2) return { success: true, count: 0, message: "原始表單目前沒有資料。" };
    var sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, 17).getValues();
    
    // 2. 取得目標資料庫
    var targetSS = SpreadsheetApp.openById(SPREADSHEET_ID);
    var targetSheet = targetSS.getSheetByName(SCHOOL_SHEET_NAME);
    var targetLastRow = targetSheet.getLastRow();
    
    // 3. 收集已經存在目標資料庫的單號 (A 欄)
    var existingIds = [];
    if (targetLastRow >= 2) {
      existingIds = targetSheet.getRange(2, 1, targetLastRow - 1, 1).getValues().flat().map(String);
    }
    
    var newDataToAppend = [];
    
    // 4. 比對並篩選出「新單號」
    sourceData.forEach(function(row) {
      var id = String(row[0]).trim();
      // 如果單號不為空，且不在現有清單中，就視為新資料
      if (id && existingIds.indexOf(id) === -1) {
         // 將 A 到 Q 欄提取出來，並補上 R, S, T 三個空欄位，湊齊 20 欄
         var newRow = row.slice(0, 17); 
         newRow.push("", "", ""); // 指派人員(R), 是否取消(S), 是否完成(T)
         newDataToAppend.push(newRow);
      }
    });
    
    // 5. 寫入新資料到最下方
    if (newDataToAppend.length > 0) {
      targetSheet.getRange(targetLastRow + 1, 1, newDataToAppend.length, 20).setValues(newDataToAppend);
    }
    
    // 6. 對整張表進行實體排序 (依據 B 欄：預約時間)
    var newLastRow = targetSheet.getLastRow();
    if (newLastRow >= 2) {
      // 範圍：A2 到 T 最後一列
      var sortRange = targetSheet.getRange(2, 1, newLastRow - 1, 20);
      // column: 2 代表 B 欄，ascending: true 代表遞增排序 (舊到新)
      sortRange.sort({column: 2, ascending: true});
    }
    
    return { 
      success: true, 
      count: newDataToAppend.length, 
      message: newDataToAppend.length > 0 ? "同步完成！新增了 " + newDataToAppend.length + " 筆預約，並已重新排序。" : "目前沒有新的預約需要同步。" 
    };
    
  } catch (e) {
    return { success: false, message: "同步失敗：" + e.message };
  }
}

/* ========================================================================= */

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
function getSchoolPersonnelList() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat().filter(String);
}

function addSchoolPerson(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  var colD = sheet.getRange("D:D").getValues();
  var lastRow = colD.filter(String).length; 
  sheet.getRange(lastRow + 1, 4).setValue(name);
  return "OK";
}

function removeSchoolPerson(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PERSONNEL_CONFIG_SHEET);
  var data = sheet.getRange("D:D").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == name) {
      sheet.getRange(i + 1, 4).clearContent();
      sheet.getRange(i + 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS);
      return "OK";
    }
  }
  return "Not Found";
}

/* --- 設定功能：Email 寄送清單 (後端寄信用) --- */
function getEmailListData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data.map(r => ({ email: r[0], name: r[1] })).filter(r => r.email && r.name);
}

function addEmailPerson(name, email) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMAIL_LIST_SHEET);
  sheet.appendRow([email, name]);
  return "OK";
}

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
    if (data[i][1] == name) return data[i][0];
  }
  return "";
}

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