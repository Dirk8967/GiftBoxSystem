// ================================================================= //
//                       資料管理 - 清除功能
// ================================================================= //

/**
 * 通用的輔助函式，用來清除指定工作表的資料（保留標頭）
 * @param {string} sheetName - 要清除資料的工作表名稱
 * @returns {Object} 包含操作結果的物件
 * @private
 */
function clearSheetData_(sheetName) {
  try {
    if (!sheetName) throw new Error("未提供工作表名稱。");
    
    // ORDER_INVENTORY_SHEET_ID 應為您在其他地方定義的常數
    const spreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`找不到名為 "${sheetName}" 的工作表。`);
    }

    const lastRow = sheet.getLastRow();
    // 如果只有標頭或沒有資料，則無需執行
    if (lastRow > 1) {
      // 從第 2 列開始，刪除到最後一列
      sheet.deleteRows(2, lastRow - 1);
    }
    
    return { success: true, message: `工作表 [${sheetName}] 的資料已成功清除。` };
  } catch (e) {
    console.error(`清除工作表 [${sheetName}] 資料時發生錯誤: ` + e.toString());
    return { success: false, error: e.message };
  }
}


// --- 以下是供前端呼叫的四個公開函式 ---

function clearCaseOrders() {
  return clearSheetData_("使用者訂單(箱)");
}

function clearLooseOrders() {
  return clearSheetData_("使用者訂單(盒)");
}

function clearShipmentRecords() {
  return clearSheetData_("管理員出貨");
}

function clearPurchaseRecords() {
  return clearSheetData_("管理員進貨");
}
