// ===== 1. 網頁入口 =====
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('檔案素養教育工作配置系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== 2. 取得初始資料 (單號列表 + 人員名單) =====
function getInitialDataAPI() {
  // 取得單號列表 (從工作安排表讀取，避免依賴 A4)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName); // 工作安排
  const lastRow = targetSheet.getLastRow();
  // 假設單號在 A 欄 (從第 2 行開始)
  const orders = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues()
                 .flat()
                 .filter(id => id && id !== ""); // 過濾空值

  // 取得人員清單 (重用您原本的函數)
  const staffConfig = getStaffConfigData();
  const stats = getStaffStatistics();

return {
    orders: orders,
    staff: staffConfig,
    stats: stats // 【新增這行】回傳統計數據
  };
}

// ===== 3. 讀取單筆詳細資料 (修正版) =====
function getOrderDataAPI(orderId) {
  try {
    // 【修正開始】手動初始化 targetSheet 快取
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) {
      CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    }
    // 【修正結束】
    
    // 呼叫您原本的函數來找行號
    const row = findMatchingRow(orderId);
    if (row === -1) return { error: "找不到該單號 (ID: " + orderId + ")，請確認工作安排表是否有資料。" };

    // 取得原始資料 (這部分原本就有防呆，不用改)
    const sourceRow = getSourceDataByKey(orderId); 
    
    // 取得工作安排資料
    const workData = getWorkArrangementData(orderId); 
    
    if (!workData) return { error: "無法讀取工作安排資料 (C~N欄)" };

    return {
      // 申請資訊
      applicant: {
        unit: sourceRow ? sourceRow[17] : "",  
        date: sourceRow ? formatDate(sourceRow[15]) : "",
        time: sourceRow ? sourceRow[16] : "",
        count: sourceRow ? sourceRow[21] : "",
        leader: sourceRow ? (sourceRow[19] || "") + " " + (sourceRow[20] || "") : "",
        phone: sourceRow ? sourceRow[23] : "",
        course: sourceRow ? sourceRow[18] : ""
      },
      // 工作配置
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

// ===== 4. 儲存資料 (修正版) =====
function saveOrderDataAPI(form) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!CACHE.targetSheet) {
      CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);
    }

    const orderId = form.orderId;
    const updates = [];

    // 準備更新陣列
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

    // 呼叫原本的批次寫入函數
    batchWriteToWorkArrangement(orderId, updates);
    
    // 重新計算統計數據
    const newStats = getStaffStatistics();
    
    return { success: true, message: "✅ 儲存成功！", newStats: newStats };
    
  } catch (e) {  // <--- 這裡原本少了一個 } 導致程式崩潰
    throw new Error("儲存失敗：" + e.message);
  }
}

// ===== 5. 執行特殊動作 (轉接器) =====
// 這部分最重要，因為原本的函數會跳出 UI 視窗，網頁版會報錯，所以要用「無 UI 版」的邏輯
function runBackendActionAPI(action, orderId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  
  // 重要：為了讓原本依賴 A4 儲存格的邏輯 (如 PDF 生成) 能運作，
  // 我們先由程式自動將 A4 設定為當前選擇的單號。
  mainSheet.getRange("A4").setValue(orderId);
  SpreadsheetApp.flush(); // 強制更新試算表
  
  // 更新快取，確保資料是最新的
  refreshCurrentOrder(); 

  try {
    if (action === 'notifyUser') {
      // 呼叫原本的寄信函數 (需修改或確保無 UI 阻擋)
      // 建議：直接在這裡寫一段簡化版寄信邏輯，避免 UI 互動
      return webSendApplicantNotification(); 
    } 
    else if (action === 'addToCalendar') {
      // 呼叫原本的日曆函數 (需處理 UI 確認視窗)
      // 這裡呼叫一個修改過的無 UI 版本
      return webAddToCalendar();
    } 
    else if (action === 'sendPDF') {
      // 呼叫 PDF 函數
      // 這裡呼叫一個修改過的無 UI 版本
      return webExportAndSendPDF();
    }
    return "未知的指令";
  } catch (e) {
    throw new Error("執行失敗：" + e.message);
  }
}

// ===== Web 版專用：無 UI 版本的執行函數 =====

function webSendApplicantNotification() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const applicantEmail = spreadsheet.getRangeByName("申請人電子郵件").getValue();
  
  if (!applicantEmail) throw new Error("找不到申請人電子郵件");
  
  const subject = "檔案素養教育（開箱檔案館–到館學習）系統已受理，請靜候承辦人通知說明";
  const body = `您好！您的「開箱檔案館–到館學習」申請案已受理。案件仍在處理中，本局將依申請提出順序及各時段受理情形，安排並告知申請結果，感謝您的耐心等候。\n\n黃宇暘敬上\n國家發展委員會檔案管理局`;
  
  GmailApp.sendEmail(applicantEmail, subject, body);
  return `✅ 受理通知已寄送給：${applicantEmail}`;
}

function webAddToCalendar() {
  // 直接呼叫原本的函數可能會跳確認視窗，所以我們複製核心邏輯簡化之
  // 為了安全，這裡建議直接引用您原本的 addToCalendar，
  // 但您需要去「程式碼.gs」把 addToCalendar 裡面的 ui.alert 註解掉，或者改用 try-catch 包裹
  // 為了不改動您原本代碼太麻煩，我們這裡做一個簡單的 workaround:
  
  // 執行原本的函數，但可能會卡在 ui.alert。
  // 最佳解法：請修改 `程式碼.gs` 的 addToCalendar，將 ui.alert 變成可選的。
  // 但現在，我們先重寫一個簡單版在這裡：
  
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
  const desc = `方案：${ranges.plan}\n由 Web App 自動建立`;
  
  // 檢查重複 (略過確認，直接不建立或建立)
  const existing = calendar.getEvents(start, end, {search: title});
  if (existing.length > 0) {
    return "⚠️ 行事曆已有相同活動，已略過建立。";
  }
  
  calendar.createEvent(title, start, end, {description: desc, location: "國家檔案館"});
  return "✅ 已加入 Google 行事曆！";
}

function webExportAndSendPDF() {
  // 這是最複雜的，因為涉及到匯出 PDF。
  // 原本的 exportAndSendPDF 包含確認視窗。
  // 我們這裡呼叫一個邏輯：
  // 1. 確保 A4 已經被設定 (在 runBackendActionAPI 已做)
  // 2. 執行匯出
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.mainSheetName);
  const orderName = mainSheet.getRange("A4").getValue();
  const selectedPlan = spreadsheet.getRangeByName("選擇方案").getValue();
  
  // 判斷程序表
  let scheduleSheetName = selectedPlan === "2小時方案" ? "（勿動）參訪程序表-2小時" : "（勿動）參訪程序表-3小時";
  const scheduleSheet = spreadsheet.getSheetByName(scheduleSheetName);
  
  if (!scheduleSheet) throw new Error(`找不到對應的程序表：${scheduleSheetName}`);
  
  // 顯示工作表以供截圖 (PDF)
  scheduleSheet.showSheet();
  SpreadsheetApp.flush();
  
  try {
    const token = ScriptApp.getOAuthToken();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf`;
    const pdfParams = '&size=A4&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false';
    
    // 生成 Blob
    const mainBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${mainSheet.getSheetId()}${pdfParams}&range=A1:D19`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_工作配置單.pdf`);
    
    const schBlob = UrlFetchApp.fetch(`${baseUrl}&gid=${scheduleSheet.getSheetId()}${pdfParams}&range=A1:E13`, {
      headers: {Authorization: `Bearer ${token}`}
    }).getBlob().setName(`${orderName}_參訪程序表.pdf`);
    
    // 寄信邏輯 (簡化版)
    const recipient = "yyhuang@archives.gov.tw"; // 主要收件人
    // 這裡可以加入更多收件人邏輯...
    
    GmailApp.sendEmail(recipient, `【通知】工作配置提醒 (${orderName})`, "請參閱附件。", {
      attachments: [mainBlob, schBlob]
    });
    
    return "✅ PDF 已生成並寄出！";
    
  } finally {
    scheduleSheet.hideSheet(); // 記得藏回去
  }
}

// ===== 輔助：計算人員場次統計 =====
function getStaffStatistics() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 確保快取存在
  if (!CACHE.targetSheet) CACHE.targetSheet = spreadsheet.getSheetByName(CONFIG.targetSheetName);

  const sheet = CACHE.targetSheet;
  const lastRow = sheet.getLastRow();

  // 如果沒資料，回傳空值
  if (lastRow < 2) return { lecturer: [], guide: [], facilitator: [] };

  // 讀取 C 欄到 K 欄 (共 9 欄) 的所有資料
  // C-E: 講座 (0,1,2), F-H: 導覽 (3,4,5), I-K: 引導 (6,7,8)
  const data = sheet.getRange(2, 3, lastRow - 1, 9).getValues();

  const counts = { lecturer: {}, guide: {}, facilitator: {} };

  // 遍歷所有資料進行計數
  data.forEach(row => {
    // 講座人員
    [0, 1, 2].forEach(i => addCount(counts.lecturer, row[i]));
    // 導覽人員
    [3, 4, 5].forEach(i => addCount(counts.guide, row[i]));
    // 引導人員
    [6, 7, 8].forEach(i => addCount(counts.facilitator, row[i]));
  });

  // 格式化輸出 (排序：次數多到少)
  return {
    lecturer: formatStatsHTML(counts.lecturer),
    guide: formatStatsHTML(counts.guide),
    facilitator: formatStatsHTML(counts.facilitator)
  };
}

function addCount(obj, name) {
  // 過濾掉空白或 "未定（無）"
  if (!name || name === "" || name === "未定（無）") return;
  obj[name] = (obj[name] || 0) + 1;
}

function formatStatsHTML(obj) {
  return Object.entries(obj)
    .sort((a, b) => b[1] - a[1]) // 依次數由大到小排序
    .map(([name, count]) => `<li class="list-group-item d-flex justify-content-between align-items-center">${name} <span class="badge bg-primary rounded-pill">${count} 次</span></li>`)
    .join('');
}