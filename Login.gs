// ================================================================= //
//                            設定區
// ================================================================= //
const SPREADSHEET_ID = "1xC1Hb3_V4faqL4xJIrbMO3mUQW_rUclca1xU7MrMLuY";
const ADMIN_EMAILS = ['jc8v2hz@gmail.com', 'jn8x2kz@gmail.com', 'd4208diversification@gmail.com', 'kelytiger0556@gmail.com'];
const APPLICANTS_CACHE_KEY = 'allApplicantsData_v3'; 
const CACHE_DURATION_SECONDS = 300; 

// ================================================================= //
//                        核心輔助函式
// ================================================================= //
function getAuthSheet_() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheets()[0]; 
    if (!sheet) { throw new Error("在使用者授權試算表中找不到第一個工作表。"); }
    return sheet;
  } catch (e) {
    throw new Error("讀取使用者授權試算表設定時發生內部錯誤: " + e.message);
  }
}

function getCurrentUserEmail_() {
  try {
    const user = Session.getActiveUser();
    return user ? user.getEmail().toLowerCase() : null;
  } catch (e) { return null; }
}
  
// ================================================================= //
//                        主要路由與頁面服務
// ================================================================= //
function doGet(e) {

  // --- 【全新 API 路由，使用 JSONP 模式】 ---
  if (e.parameter.action === 'getSummary') {
    // 獲取前端指定的 callback 函式名稱，如果沒有則預設為 'callback'
    const callback = e.parameter.callback || 'callback';
    
    // 呼叫總彙整函式，獲取資料
    const reportData = getGrandSummaryReportData();
    const jsonString = JSON.stringify(reportData);
    
    // 【核心修改】將 JSON 字串包裹在 callback 函式中
    const outputString = `${callback}(${jsonString});`;
    
    // 使用 ContentService 回傳，並將 MIME 類型設為 JAVASCRIPT
    return ContentService.createTextOutput(outputString)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }




  const page = e.parameter.page;
  if (page === 'admin') {
    const userEmail = getCurrentUserEmail_();
    if (userEmail && ADMIN_EMAILS.includes(userEmail)) {
      return HtmlService.createTemplateFromFile('Admin').evaluate().setTitle('管理後台');
    } else {
      return HtmlService.createTemplateFromFile('Admin_Unauthorized').evaluate().setTitle('存取遭拒');
    }
  }
  
  if (!checkActivityPeriod_()) {
    return HtmlService.createTemplateFromFile('PeriodClosed').evaluate().setTitle('非活動檔期');
  }
  
  const userEmail = getCurrentUserEmail_();
  if (page === 'pending') {
    return HtmlService.createTemplateFromFile('Pending').evaluate().setTitle('審核中');
  }

  if (!userEmail) {
      return HtmlService.createTemplateFromFile('Unauthorized').evaluate().setTitle('存取遭拒');
  }

  let sheet;
  try {
    sheet = getAuthSheet_();
  } catch (err) {
    return HtmlService.createHtmlOutput("系統錯誤：無法連接到資料庫。");
  }

  if (isUserAuthorized_(userEmail, sheet)) {
    return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('歡迎使用');
  }
  if (isUserPending_(userEmail, sheet)) {
    return HtmlService.createTemplateFromFile('Pending').evaluate().setTitle('審核中');
  }
  return HtmlService.createTemplateFromFile('Unauthorized').evaluate().setTitle('存取遭拒');
}

/**
 * 【已更新】從伺服器獲取指定 HTML 子頁面的內容。
 * @param {string} fileName - HTML 子頁面的檔案名稱 (不含 .html 後綴)
 * @returns {string} HTML 內容字串
 */
function getPartialHtmlFromFile(fileName) {
  try {
    // 【最終修正】確保白名單包含所有 Admin 和 Index 頁面會用到的子頁面
    const allowedPages = [
        // Admin.html 用的子頁面
        'Admin_UserManagement', 
        'Admin_CaseAdminManagement', 
        'Admin_ProductManagement', 
        'Admin_SiteInfoManagement', 
        'Admin_BusinessUnitManagement',
        'Admin_Settings',
        
        // index.html 用的子頁面
        'Page_CaseOrder', 
        'Page_LooseOrder', 
        'Page_CaseSummary', 
        'Page_LooseSummary',
        'Page_GrandSummary',
        'Page_MyProfile',
        'Page_OrderHistory',
        'Page_DeliveryHistory',
        'Page_CaseOrderAdmin',
        'Page_LooseOrderAdmin',
        'Page_IOManagement', 
        'Page_InventoryManagement',
        'Page_LoosePeriodClosed'
    ];
    
    if (!fileName || !allowedPages.includes(fileName)) {
      // 在日誌中記錄被拒絕的檔案名稱，方便偵錯
      Logger.log("getPartialHtmlFromFile: 請求被拒絕，請求的檔案名稱 '" + fileName + "' 不在白名單中。");
      throw new Error("無效的頁面請求。");
    }
    return HtmlService.createHtmlOutputFromFile(fileName).getContent();
  } catch (e) {
    console.error("getPartialHtmlFromFile 錯誤 (請求檔案: " + fileName + "): " + e.toString());
    return '<div class="container"><p style="color:red;">載入頁面內容失敗：' + e.message + '</p></div>';
  }
}
  
function isUserAuthorized_(userEmail, sheet) {
  try {
    if (!userEmail) return false;
    if (sheet.getLastRow() < 2) return false;
    // 【核心修改】讀取範圍從 C 欄到 F 欄 (共 4 欄)
    const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 4).getValues(); // C:Email, D:單位, E:加油站, F:授權
    for (const row of data) {
      // 在這個範圍內，Email 是第0個元素(row[0]), 授權狀態是第3個元素(row[3])
      if (String(row[0] || '').toLowerCase().trim() === userEmail && row[3] === true) {
        return true;
      }
    }
    return false;
  } catch (e) { return false; }
}

// 請替換此函式
function isUserPending_(userEmail, sheet) {
  try {
    if (!userEmail) return false;
    if (sheet.getLastRow() < 2) return false;
    
    // 【核心修改】讀取範圍從 C 欄到 F 欄 (共 4 欄)
    const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 4).getValues();

    for (const row of data) {
      // 在這個範圍內，Email 是第0個元素(row[0]), 授權狀態是第3個元素(row[3])
      if (String(row[0] || '').toLowerCase().trim() === userEmail && row[3] !== true) {
        return true;
      }
    }
    return false;
  } catch (e) { return false; }
}


function requestAuthorization(userInfo) {
  try {
    const userEmail = getCurrentUserEmail_();
    if (!userEmail) throw new Error("無法獲取使用者 Email。");
    const { name, employeeId } = userInfo;
    if (!name || !employeeId) throw new Error("姓名和員工編號不可為空。");
    const sheet = getAuthSheet_();
    
    // 【核心修改】appendRow 的陣列擴充為 7 個元素
    sheet.appendRow([
        name.trim(),                  // A: 姓名
        "'" + String(employeeId).trim(), // B: 員工編號
        userEmail,                    // C: Email
        '',                           // D: 隸屬單位名稱 (預設為空)
        false,                        // E: 是否為加油站人員 (預設為否)
        false,                        // F: 審核 (預設為否)
        ''                            // G: 備註 (預設為空)
    ]);

    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat("@");
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) { return { success: false, error: err.message }; }
}

function checkIfCurrentUserIsApproved() {
  const userEmail = getCurrentUserEmail_();
  if (!userEmail) return false; 
  return isUserAuthorized_(userEmail, getAuthSheet_());
}

function checkUserRoles() {
  try {
    const userEmail = getCurrentUserEmail_();
    if (!userEmail) return { isAuthorized: false, isCaseAdmin: false, isMainAdmin: false, isGasStationStaff: false, email: '' };

    const isMainAdmin = ADMIN_EMAILS.includes(userEmail);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = spreadsheet.getSheetByName("訂購網頁權限"); // 直接用名稱獲取
    const caseAdminSheet = spreadsheet.getSheetByName("成箱管理員權限");

    let isAuthorized = isUserAuthorized_(userEmail, userSheet);
    let isGasStationStaff = false;
    // 從「訂購網頁權限」工作表中，一次性獲取授權狀態和加油站人員狀態
    if (userSheet) {
      const data = userSheet.getRange("C2:F" + userSheet.getLastRow()).getValues(); // C:Email, E:加油站人員, F:審核
      for (const row of data) {
        if (String(row[0]).toLowerCase().trim() === userEmail) {
          if(row[3] === true) isAuthorized = true; // F欄是審核
          if(row[2] === true) isGasStationStaff = true; // E欄是加油站人員
          break; // 找到後即可跳出迴圈
        }
      }
    }

    let isCaseAdmin = false;
    if (caseAdminSheet) {
      isCaseAdmin = isUserAuthorized_(userEmail, caseAdminSheet); // 重複使用 isUserAuthorized_ 邏輯
    }
    
    const finalRoles = { isAuthorized, isCaseAdmin, isMainAdmin, isGasStationStaff, email: userEmail };
    Logger.log("checkUserRoles: 最終回傳的角色物件: " + JSON.stringify(finalRoles));
    return finalRoles;
  } catch (e) {
    return { isAuthorized: false, isCaseAdmin: false, isMainAdmin: false, isGasStationStaff: false, email: '', error: e.message };
  }
}

function getUserProfileData() {
  try {
    const userEmail = getCurrentUserEmail_();
    if (!userEmail) throw new Error("無法獲取使用者資訊。");
    const sheet = getAuthSheet_();
    if (sheet.getLastRow() < 2) return { name: "未找到", employeeId: "N/A", email: userEmail };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (const row of data) {
      if (String(row[2] || '').toLowerCase().trim() === userEmail) {
        return { name: String(row[0]), employeeId: String(row[1]), email: userEmail };
      }
    }
    return { name: "未列於授權清單", employeeId: "N/A", email: userEmail };
  } catch (e) { return { error: "伺服器讀取個人資料時發生錯誤：" + e.message }; }
}
