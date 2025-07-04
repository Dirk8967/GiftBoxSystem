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

    // 【核心修改】讀取範圍擴大到 G 欄 (共 7 欄)
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    
    const caseAdmins = values.map(function(row, index) {
      // 【核心修改】回傳的物件新增兩個欄位
      return {
        rowNumber: index + 2,
        name: String(row[0] || ''),
        employeeId: String(row[1] || ''),
        email: String(row[2] || ''),
        affiliatedUnit: String(row[3] || ''),     // D欄: 隸屬單位名稱
        isGasStationStaff: row[4] === true,     // E欄: 是否為加油站人員
        isApproved: row[5] === true,              // F欄: 審核
        remarks: String(row[6] || '')             // G欄: 備註
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
 * (C)reate - [更新版] 新增成箱管理員資料
 */
function addCaseAdminData(adminData) {
  try {
    const { name, employeeId, email, affiliatedUnit, isGasStationStaff, isApproved, remarks } = adminData; 
    if (!name || !String(employeeId).trim() || !email) {
      throw new Error("姓名、員工編號和 Email 為必填欄位。");
    }
    
    const sheet = getCaseAdminSheet_();
    const employeeIdStr = String(employeeId).trim();

    // 【核心修改】寫入的陣列現在包含 7 個元素
    sheet.appendRow([
        name.trim(), 
        "'" + employeeIdStr, 
        email.trim().toLowerCase(), 
      affiliatedUnit || '',
      isGasStationStaff === true,
        isApproved === true, 
        remarks || '' 
    ]);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat("@"); 

    SpreadsheetApp.flush();
    clearCaseAdminCache_();
    return { success: true };
  } catch (e) {
    console.error('新增成箱管理員失敗: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * (U)pdate - [更新版] 更新指定列的成箱管理員資料
 */
function updateCaseAdminData(adminData) {
  try {
    const { rowNumber, name, employeeId, email, affiliatedUnit, isGasStationStaff, isApproved, remarks } = adminData; 
    if (!rowNumber || !name || !String(employeeId).trim() || !email) { 
      throw new Error("缺少必要更新資訊。");
    }
    
    const sheet = getCaseAdminSheet_();
    const employeeIdStr = String(employeeId).trim();
    
    // 【核心修改】更新的範圍和內容擴大到 7 欄
    sheet.getRange(rowNumber, 1, 1, 7).setValues([[ 
        name.trim(), 
        "'" + employeeIdStr, 
        email.trim().toLowerCase(), 
        affiliatedUnit || '',
        isGasStationStaff === true,
        isApproved === true,
        remarks || '' 
    ]]);

    sheet.getRange(rowNumber, 2).setNumberFormat("@");

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



/**
 * 【新增】獲取所有成箱訂單以供管理員查看
 * @returns {Object} 包含 { success: boolean, data?: Array<Array<any>>, error?: string } 的物件
 */
function getCaseOrdersForAdmin() {
  try {
    Logger.log("getCaseOrdersForAdmin: 函式開始執行。");
    
    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    if (!orderSheet || orderSheet.getLastRow() < 2) {
      Logger.log("getCaseOrdersForAdmin: '使用者訂單(箱)' 工作表為空或不存在。");
      return { success: true, data: [] };
    }

    const lastRow = orderSheet.getLastRow();
    Logger.log("getCaseOrdersForAdmin: 訂單工作表的最後一列是: " + lastRow);

    const fullRange = orderSheet.getRange("A2:Q" + lastRow).getValues();
    Logger.log("getCaseOrdersForAdmin: 成功讀取到 " + fullRange.length + " 列原始訂單資料。");
    
    const displayData = fullRange.map(function(row, index) {
      const rowNumber = index + 2;
      return [
        row[0],  // A: 訂購人姓名
        row[1],  // B: 訂購人電話
        row[2],  // C: 商品名稱
        row[3],  // D: 盒數
        row[4],  // E: 總計金額
        row[5],  // F: 寄送地點
        Date(row[6]),  // G: 寄送日期
        row[7],  // H: 訂單隸屬站點
        row[8],  // I: (FALSE)
        Date(row[10]), // K: 訂購時間
        row[11], // L: 操作人員姓名
        row[13], // N: 訂單編號
        rowNumber 
      ];
    });
    Logger.log("getCaseOrdersForAdmin: 資料轉換完成，準備回傳 " + displayData.length + " 筆訂單資料。");
    Logger.log(displayData);
    return { success: true, data: displayData };
  } catch (e) {
    Logger.log("getCaseOrdersForAdmin 發生嚴重錯誤: " + e.toString() + "\n" + e.stack);
    console.error("getCaseOrdersForAdmin 錯誤: " + e.toString());
    return { success: false, error: "讀取成箱訂單時發生伺服器錯誤: " + e.message };
  }
}

/**
 * 【新增】更新指定列的訂單編號
 * @param {Object} updateInfo - { rowNumber, newOrderNumber }
 * @returns {Object} { success: boolean, error?: string }
 */
function updateOrderNumber(updateInfo) {
  try {
    const { rowNumber, newOrderNumber } = updateInfo;
    if (!rowNumber || newOrderNumber === undefined || newOrderNumber === null) {
      throw new Error("缺少必要的更新資訊（列號和新的訂單編號）。");
    }

    const operatorEmail = Session.getActiveUser().getEmail();
    const operatorName = getOperatorNameByEmail_(operatorEmail);
    const operationTimestamp = new Date();

    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    
    orderSheet.getRange(rowNumber, 14, 1, 4).setValues([[
      newOrderNumber,     // N: 訂單編號
      operatorName,       // O: 成箱管理員姓名
      operatorEmail,      // P: 成箱管理員信箱
      operationTimestamp  // Q: 成箱管理員操作時間
    ]]);

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    console.error("updateOrderNumber 錯誤: " + e.toString());
    return { success: false, error: "更新訂單編號時發生伺服器錯誤: " + e.message };
  }
}
