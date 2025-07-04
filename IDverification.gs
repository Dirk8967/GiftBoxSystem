// ================================================================= //
//                 管理後台 - 使用者權限管理 CRUD 相關函式
// ================================================================= //

// 注意：此檔案中的函式會依賴定義在 Login.gs 中的輔助函式與常數。

/**
 * 清除使用者授權列表的快取
 */
function clearApplicantsCache_() {
  try {
    CacheService.getScriptCache().remove(APPLICANTS_CACHE_KEY); // APPLICANTS_CACHE_KEY 在 Login.gs
    Logger.log('使用者授權資料快取已清除。');
  } catch (e) {
    console.error('清除使用者授權快取失敗: ' + e.toString());
  }
}


/**
 * 新增使用者資料到試算表，並強制設定員工編號格式以保留前導零
 */
function addUserData(userData) {
  try {
    const { name, employeeId, email, affiliatedUnit, isGasStationStaff, isApproved, remarks } = userData; 
    if (!name || !String(employeeId).trim() || !email) { 
      throw new Error("姓名、員工編號和 Email 為必填欄位。");
    }

    const sheet = getAuthSheet_(); 
    const employeeIdStr = String(employeeId).trim();

    // 【核心修改】寫入的陣列現在包含 7 個元素
    sheet.appendRow([
      name.trim(), 
      "'" + employeeIdStr, 
      email.trim().toLowerCase(), 
      affiliatedUnit || '',         // 【新增】
      isGasStationStaff === true, // 【新增】
      isApproved === true, 
      remarks || '' 
    ]);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat("@"); 

    SpreadsheetApp.flush();
    clearApplicantsCache_();
    return { success: true };
  } catch (e) {
    console.error('新增使用者失敗: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * (U)pdate - [更新版] 更新指定列的使用者資料
 */
function updateUserData(userData) {
  try {
    const { rowNumber, name, employeeId, email, affiliatedUnit, isGasStationStaff, isApproved, remarks } = userData; 
    if (!rowNumber || !name || !String(employeeId).trim() || !email) { 
      throw new Error("缺少必要更新資訊。");
    }
    
    const sheet = getAuthSheet_();
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
    clearApplicantsCache_();
    return { success: true };
  } catch (e) {
    console.error('更新使用者失敗 (列號: ' + userData.rowNumber + '): ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 刪除指定列的使用者資料
 */
function deleteUserData(deleteInfo) {
  try {
    const { rowNumber } = deleteInfo;
    if (!rowNumber || typeof rowNumber !== 'number' || rowNumber < 2 ) {
      throw new Error("提供的列號無效或格式不正確。");
    }
    const sheet = getAuthSheet_();
    if (rowNumber > sheet.getMaxRows() || rowNumber > sheet.getLastRow()) { 
        throw new Error("無效的列號，超出試算表範圍。");
    }
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    clearApplicantsCache_();
    return { success: true };
  } catch (e) {
    console.error('刪除使用者失敗 (列號: ' + deleteInfo.rowNumber + '): ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 獲取所有申請者的資料以供後台顯示
 */
function getApplicantsData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(APPLICANTS_CACHE_KEY); 
  if (cached != null) {
    Logger.log('getApplicantsData: 從快取讀取使用者資料。');
    try { return JSON.parse(cached); } catch (e) { Logger.log('解析使用者快取失敗，重新讀取。 Error: ' + e.message);}
  }

  Logger.log('getApplicantsData: 快取未命中，從試算表讀取使用者資料。');
  try {
    const sheet = getAuthSheet_(); 
    if (sheet.getLastRow() < 2) { return []; } 

    // 【核心修改】讀取範圍擴大到 G 欄 (共 7 欄)
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); 
    
    const applicants = values.map(function(row, index) {
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
    
    cache.put(APPLICANTS_CACHE_KEY, JSON.stringify(applicants), CACHE_DURATION_SECONDS);
    Logger.log("使用者資料已讀取並快取，共 " + applicants.length + " 筆。");
    return applicants;
  } catch (e) {
    console.error('讀取申請者資料失敗 (getApplicantsData): ' + e.toString());
    throw new Error('伺服器讀取使用者資料時發生錯誤。');
  }
}

/**
 * [更新版] 更新試算表中的審核狀態
 */
function updateApprovalStatus(approvalData) {
  if (!Array.isArray(approvalData)) {
    return { success: false, error: "提供的資料格式錯誤。" };
  }
  try {
    const sheet = getAuthSheet_(); 
    let changesMade = 0;
    approvalData.forEach(function(item) {
      if (typeof item.rowNumber === 'number' && typeof item.isApproved === 'boolean') {
        // 【核心修改】將寫入的欄位從第 4 欄 (D欄) 改為第 6 欄 (F欄)
        sheet.getRange(item.rowNumber, 6).setValue(item.isApproved);
        changesMade++;
      } else {
        console.warn('updateApprovalStatus: 收到無效的項目資料', item);
      }
    });

    if (changesMade > 0) {
        SpreadsheetApp.flush();
        clearApplicantsCache_(); 
    }
    return { success: true, message: changesMade > 0 ? '審核狀態已更新。' : '沒有變更被儲存。' };
  } catch (e) {
    console.error('更新審核狀態失敗: ' + e.toString());
    return { success: false, error: e.message };
  }
}
