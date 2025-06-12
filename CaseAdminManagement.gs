// ================================================================= //
//               成箱管理員權限管理 - 後端邏輯
// ================================================================= //

// 注意：此檔案中的 SPREADSHEET_ID 會與 Login.gs 中定義的共用，
// 因為它們操作的是同一個試算表檔案，只是不同的工作表。
const CASE_ADMIN_CACHE_KEY = 'allCaseAdminsData_v1';
const CASE_ADMIN_CACHE_DURATION_SECONDS = 300; // 快取 5 分鐘

/**
 * 輔助函式：獲取「成箱管理員權限」工作表物件 (試算表的第二頁)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 工作表物件
 */
function getCaseAdminSheet_() {
  try {
    // SPREADSHEET_ID 是定義在 Login.gs 中的常數，此處可以直接使用
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); 
    if (!spreadsheet) {
        throw new Error("伺服器錯誤：無法開啟使用者授權清單試算表。");
    }
    // getSheets()[1] 會獲取第二個工作表 (索引從 0 開始)
    const sheet = spreadsheet.getSheets()[1]; 
    if (!sheet || sheet.getName() !== '成箱管理員權限') { // 雙重確認工作表名稱
        Logger.log("getCaseAdminSheet_ 錯誤: 找不到名為 '成箱管理員權限' 的第二個工作表。");
        throw new Error("伺服器錯誤：找不到 '成箱管理員權限' 工作表。");
    }
    return sheet;
  } catch (e) {
    console.error("getCaseAdminSheet_ 捕捉到嚴重錯誤: " + e.toString());
    throw new Error("讀取成箱管理員工作表設定時發生內部錯誤: " + e.message);
  }
}

/**
 * 輔助函式：清除成箱管理員資料快取
 */
function clearCaseAdminCache_() {
  try {
    CacheService.getScriptCache().remove(CASE_ADMIN_CACHE_KEY);
    Logger.log("成箱管理員資料快取已清除。");
  } catch (e) {
    console.error('清除成箱管理員快取失敗: ' + e.toString());
  }
}

/**
 * 獲取所有成箱管理員的資料
 * @returns {Array<Object>} 包含所有成箱管理員資料的陣列。
 */
function getCaseAdminListData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CASE_ADMIN_CACHE_KEY);
  if (cached != null) {
    Logger.log('getCaseAdminListData: 從快取讀取成箱管理員資料。');
    try { return JSON.parse(cached); } catch (e) { Logger.log('解析成箱管理員快取失敗，重新讀取。');}
  }

  Logger.log('getCaseAdminListData: 從試算表讀取成箱管理員資料。');
  try {
    const sheet = getCaseAdminSheet_();
    if (sheet.getLastRow() < 2) { return []; } // 只有標頭或空表

    const numDataRows = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, numDataRows, 5).getValues(); // A到E欄 (姓名, 員工編號, Email, 授權, 備註)
    
    const caseAdmins = values.map(function(row, index) {
      return {
        rowNumber: index + 2,
        name: String(row[0] || ''),
        employeeId: String(row[1] || ''),
        email: String(row[2] || ''),
        isApproved: row[3] === true,
        remarks: String(row[4] || '')
      };
    });
    
    cache.put(CASE_ADMIN_CACHE_KEY, JSON.stringify(caseAdmins), CASE_ADMIN_CACHE_DURATION_SECONDS);
    Logger.log("成箱管理員資料已讀取並快取，共 " + caseAdmins.length + " 筆。");
    return caseAdmins;
  } catch (e) {
    console.error('讀取成箱管理員資料失敗: ' + e.toString());
    throw new Error('伺服器讀取成箱管理員資料時發生錯誤。');
  }
}

/**
 * 新增成箱管理員資料，並強制設定員工編號格式以保留前導零
 * @param {Object} adminData - { name, employeeId, email, isApproved, remarks }
 * @returns {Object} { success: boolean, error?: string }
 */
function addCaseAdminData(adminData) {
  try {
    const { name, employeeId, email, isApproved, remarks } = adminData; 
    if (!name || !String(employeeId).trim() || !email) {
      throw new Error("姓名、員工編號和 Email 為必填欄位。");
    }
    
    const sheet = getCaseAdminSheet_();
    const employeeIdStr = String(employeeId).trim();

    sheet.appendRow([
        name.trim(), 
        "'" + employeeIdStr, // 在員工編號前加上單引號
        email.trim().toLowerCase(), 
        isApproved === true, 
        remarks || '' 
    ]);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat("@"); // B欄: 員工編號

    SpreadsheetApp.flush();
    clearCaseAdminCache_();
    return { success: true };
  } catch (e) {
    console.error('新增成箱管理員失敗: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 更新指定列的成箱管理員資料，並強制設定員工編號格式以保留前導零
 * @param {Object} adminData - { rowNumber, name, employeeId, email, isApproved, remarks }
 * @returns {Object} { success: boolean, error?: string }
 */
function updateCaseAdminData(adminData) {
  try {
    const { rowNumber, name, employeeId, email, isApproved, remarks } = adminData; 
    if (!rowNumber || !name || !String(employeeId).trim() || !email) { 
      throw new Error("缺少必要更新資訊 (列號、姓名、員工編號、Email)。");
    }
    
    const sheet = getCaseAdminSheet_();
    const employeeIdStr = String(employeeId).trim();
    
    sheet.getRange(rowNumber, 1, 1, 5).setValues([[ 
        name.trim(), 
        "'" + employeeIdStr, // 在員工編號前加上單引號
        email.trim().toLowerCase(), 
        isApproved === true,
        remarks || '' 
    ]]);

    sheet.getRange(rowNumber, 2).setNumberFormat("@"); // B欄: 員工編號

    SpreadsheetApp.flush();
    clearCaseAdminCache_();
    return { success: true };
  } catch (e) {
    console.error('更新成箱管理員失敗 (列號: ' + adminData.rowNumber + '): ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 刪除指定列的成箱管理員資料
 * @param {Object} deleteInfo - { rowNumber }
 * @returns {Object} { success: boolean, error?: string }
 */
function deleteCaseAdminData(deleteInfo) {
  try {
    const { rowNumber } = deleteInfo;
    if (!rowNumber || typeof rowNumber !== 'number' || rowNumber < 2 ) {
      throw new Error("提供的列號無效或格式不正確。");
    }
    const sheet = getCaseAdminSheet_();
    if (rowNumber > sheet.getMaxRows() || rowNumber > sheet.getLastRow()) { 
        throw new Error("無效的列號，超出試算表範圍。");
    }
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    clearCaseAdminCache_();
    return { success: true };
  } catch (e) {
    console.error('刪除成箱管理員失敗 (列號: ' + deleteInfo.rowNumber + '): ' + e.toString());
    return { success: false, error: e.message };
  }
}
