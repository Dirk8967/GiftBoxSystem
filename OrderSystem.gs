// 【新增】訂單與庫存試算表的 ID
const ORDER_INVENTORY_SHEET_ID = "1tsj99qTZO0XunzYKNeJQw3GuCeTLneipOeiS6dnUjtg";


// ================================================================= //
//                       訂購系統 - 後端邏輯
// ================================================================= //

/**
 * 【核心】從伺服器獲取指定 HTML 子頁面的內容。
 * 這是實現頁面切換的核心函式。
 * @param {string} fileName - HTML 子頁面的檔案名稱 (不含 .html 後綴)
 * @returns {string} HTML 內容字串
 */
function getSubPageHtml(fileName) {
  try {
    // 建立一個允許的子頁面清單，防止傳入惡意或不存在的檔案名稱
    const allowedPages = [
        'Page_CaseOrder', 
        'Page_LooseOrder', 
        'Page_CaseSummary', 
        'Page_LooseSummary',
        'Page_CaseOrderAdmin',
        'Page_MyProfile',
        'Page_OrderHistory',
        'Page_DeliveryHistory',
        'Page_IOManagement',         
        'Page_InventoryManagement'
    ];
    
    if (!fileName || !allowedPages.includes(fileName)) {
      Logger.log("getSubPageHtml: 請求了不被允許或無效的檔案名稱: " + fileName);
      throw new Error("無效的頁面請求。");
    }

    Logger.log("getSubPageHtml: 正在請求檔案 " + fileName);
    return HtmlService.createHtmlOutputFromFile(fileName).getContent();
  } catch (e) {
    console.error("getSubPageHtml 錯誤 (請求檔案: " + fileName + "): " + e.toString());
    // 回傳一個錯誤提示的 HTML
    return '<div class="container"><p style="color:red;">載入頁面內容失敗：' + e.message + '</p></div>';
  }
}


/**
 * 【佔位函式範例】獲取當前登入使用者的 Email。
 * 這個函式會被前端呼叫以顯示在右上角。
 * 我們可以重複使用 Login.gs 中的 getCurrentUserEmail_()，但為了模組獨立，
 * 也可以在這裡建立一個公開版本。
 * @returns {string} 當前登入者的 Email
 */
function getCurrentUserEmailForDisplay() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '無法獲取使用者';
  }
}

/**
 * 【新增】輔助函式：根據 Email 從「使用者授權清單」中獲取姓名
 * @param {string} email - 要查詢的 Email
 * @returns {string} 找到的使用者姓名，或 Email 本身（如果找不到）
 */
function getOperatorNameByEmail_(email) {
  if (!email) return "未知操作員";
  try {
    // SPREADSHEET_ID 是定義在 Login.gs 中的常數，用於存取使用者授權清單
    const userAuthSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    if (userAuthSheet.getLastRow() < 2) return email; // 如果清單為空，直接回傳 email

    const data = userAuthSheet.getRange(2, 1, userAuthSheet.getLastRow() - 1, 3).getValues(); // 讀取 A:姓名, B:員工編號, C:Email
    const normalizedEmail = email.toLowerCase().trim();

    for (const row of data) {
      if (String(row[2] || '').toLowerCase().trim() === normalizedEmail) {
        return String(row[0] || email); // 回傳姓名，如果姓名為空則回傳 email
      }
    }
    return email; // 如果循環結束都沒找到，回傳 email
  } catch (e) {
    console.error("getOperatorNameByEmail_ 錯誤: " + e.toString());
    return email; // 出錯時也回傳 email
  }
}


/**
 * 【已更新】接收前端訂單資料，依照新的欄位順序寫入「使用者訂單(箱)」工作表
 * @param {Object} orderData - 從前端傳來的訂單物件
 * @returns {Object} { success: boolean, error?: string }
 */
function submitCaseOrderToServer(orderData) {
  try {
    const operatorEmail = Session.getActiveUser().getEmail();
    const operatorName = getOperatorNameByEmail_(operatorEmail); // 假設 getOperatorNameByEmail_ 已存在
    const orderTimestamp = new Date();

    // 驗證傳入的資料
    if (!orderData || !orderData.name || !orderData.productName || !orderData.quantity) {
      throw new Error("傳入的訂單資料不完整。");
    }

    const targetSpreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID); // ORDER_INVENTORY_SHEET_ID 應已定義
    const targetSheet = targetSpreadsheet.getSheetByName("使用者訂單(箱)");
    if (!targetSheet) {
      throw new Error("在目標試算表中找不到名為 '使用者訂單(箱)' 的工作表。");
    }

    const phoneStr = String(orderData.phone || '').trim();
    const phoneValueToWrite = phoneStr ? "'" + phoneStr : ""; 

    // 【修正】依照新的欄位順序準備要寫入的資料列 (共 13 欄，對應 A 到 M)
    const rowData = [
      orderData.name,             // A欄: 訂購人姓名
      phoneValueToWrite,          // B欄: 訂購人電話
      orderData.productName,      // C欄: 商品名稱
      orderData.quantity,         // D欄: 盒數
      orderData.totalPrice,       // E欄: 總計金額
      orderData.location,         // F欄: 寄送地點
      orderData.date,             // G欄: 寄送日期
      false,                      // H欄: (原 I 欄) 固定填入 FALSE
      '',                         // I欄: (原 J 欄) 空白
      orderTimestamp,             // J欄: (原 K 欄) 訂購時間
      operatorName,               // K欄: (原 L 欄) 操作人員姓名
      operatorEmail,              // L欄: (原 M 欄) 操作人員信箱
      ''                          // M欄: (原 H 欄) 空白
    ];

    targetSheet.appendRow(rowData);
    SpreadsheetApp.flush(); 

    const lastRow = targetSheet.getLastRow();
    // 強制設定電話欄位的格式 (B欄，索引為2，位置不變)
    if (phoneStr) {
        targetSheet.getRange(lastRow, 2).setNumberFormat("@"); 
    }
    
    return { success: true };
  } catch (e) {
    console.error("submitCaseOrderToServer 錯誤: " + e.toString() + " Stack: " + e.stack);
    return { success: false, error: "伺服器處理訂單時發生錯誤: " + e.message };
  }
}

/**
 * 【新增】接收前端的零星訂單資料，處理後寫入「使用者訂單(盒)」工作表
 * @param {Object} orderData - 從前端傳來的訂單物件
 * @returns {Object} { success: boolean, error?: string }
 */
function submitLooseOrderToServer(orderData) {
  try {
    const operatorEmail = Session.getActiveUser().getEmail();
    const operatorName = getOperatorNameByEmail_(operatorEmail);
    const orderTimestamp = new Date();

    if (!orderData || !orderData.name || !orderData.productName || !orderData.quantity) {
      throw new Error("傳入的訂單資料不完整。");
    }

    const targetSpreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID);
    const targetSheet = targetSpreadsheet.getSheetByName("使用者訂單(盒)"); // 【注意】寫入到不同的工作表
    if (!targetSheet) {
      throw new Error("在目標試算表中找不到名為 '使用者訂單(盒)' 的工作表。");
    }

    const phoneStr = String(orderData.phone || '').trim();
    const phoneValueToWrite = phoneStr ? "'" + phoneStr : ""; 

    // 【修正】依照零星訂單的欄位格式準備要寫入的資料列 (A 到 M)
    const rowData = [
      orderData.name,             // A欄: 訂購人姓名
      phoneValueToWrite,          // B欄: 訂購人電話
      orderData.productName,      // C欄: 商品名稱
      orderData.quantity,         // D欄: 盒數
      orderData.totalPrice,       // E欄: 總計金額
      orderData.location,         // F欄: 寄送地點
      false,                      // G欄: 固定填入 FALSE
      '',                         // H欄: 空白
      false,                      // I欄: 固定填入 FALSE
      '',                         // J欄: 空白
      orderTimestamp,             // K欄: 訂購時間
      operatorName,               // L欄: 操作人員姓名
      operatorEmail               // M欄: 操作人員信箱
    ];

    targetSheet.appendRow(rowData);
    SpreadsheetApp.flush(); 

    const lastRow = targetSheet.getLastRow();
    if (phoneStr) {
        targetSheet.getRange(lastRow, 2).setNumberFormat("@"); // 強制 B 欄為文字
    }
    
    return { success: true };
  } catch (e) {
    console.error("submitLooseOrderToServer 錯誤: " + e.toString());
    return { success: false, error: "伺服器處理零星訂單時發生錯誤: " + e.message };
  }
}


// --- 未來將在此處加入更多後端邏輯 ---
// 例如:
// function getOrderableProducts() { /* ... */ }
// function submitCaseOrder(orderData) { /* ... */ }
// function getUserOrderHistory() { /* ... */ }
