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
    const operatorName = getOperatorNameByEmail_(operatorEmail);
    const orderTimestamp = new Date();

    // 驗證傳入的資料，確保新欄位也一併檢查（如果需要）
    if (!orderData || !orderData.name || !orderData.productName || !orderData.quantity) {
      throw new Error("傳入的訂單資料不完整。");
    }

    const targetSpreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID);
    const targetSheet = targetSpreadsheet.getSheetByName("使用者訂單(箱)");
    if (!targetSheet) {
      throw new Error("在目標試算表中找不到名為 '使用者訂單(箱)' 的工作表。");
    }

    const phoneStr = String(orderData.phone || '').trim();
    const phoneValueToWrite = phoneStr ? "'" + phoneStr : ""; 

    // 【修正】依照新的 14 欄位格式準備要寫入的資料列 (A 到 N)
    const rowData = [
      orderData.name,             // A欄: 訂購人姓名
      phoneValueToWrite,          // B欄: 訂購人電話
      orderData.productName,      // C欄: 商品名稱
      orderData.quantity,         // D欄: 盒數
      orderData.totalPrice,       // E欄: 總計金額
      orderData.location,         // F欄: 寄送地點
      orderData.date,             // G欄: 寄送日期
      orderData.affiliatedSite,   // H欄: 【新增】訂單隸屬站點
      false,                      // I欄: (原 H 欄) 固定填入 FALSE
      '',                         // J欄: (原 I 欄) 空白
      orderTimestamp,             // K欄: (原 J 欄) 訂購時間
      operatorName,               // L欄: (原 K 欄) 操作人員姓名
      operatorEmail,              // M欄: (原 L 欄) 操作人員信箱
      ''                          // N欄: (原 M 欄) 空白
    ];

    targetSheet.appendRow(rowData);
    SpreadsheetApp.flush(); 

    const lastRow = targetSheet.getLastRow();
    // 強制設定電話欄位的格式 (B欄，索引為2，位置不變)
    if (phoneStr) {
        targetSheet.getRange(lastRow, 2).setNumberFormat("@"); 
    }
    // 未來如果隸屬站點也需要強制文字，可在此處加入
    // targetSheet.getRange(lastRow, 8).setNumberFormat("@"); 
    
    return { success: true };
  } catch (e) {
    console.error("submitCaseOrderToServer 錯誤: " + e.toString());
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


/**
 * 【已更新】獲取成箱訂單的彙總資料，並加入詳細的偵錯日誌
 */
function getCaseOrderSummaryData() {
  try {
    Logger.log("getCaseOrderSummaryData: 函式開始執行。");

    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    Logger.log("getCaseOrderSummaryData: 已獲取商品工作表。");
    
    const siteSheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID).getSheets()[0];
    Logger.log("getCaseOrderSummaryData: 已獲取站點工作表。");

    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    if (!orderSheet) {
      throw new Error("在目標試算表中找不到名為 '使用者訂單(箱)' 的工作表。");
    }
    Logger.log("getCaseOrderSummaryData: 已獲取成箱訂單工作表。");

    const lastProductRow = productSheet.getLastRow();
    if (lastProductRow < 2) { return { success: true, headers: ["站名"], summary: {}, sites: [] }; }
    const productValues = productSheet.getRange("B2:B" + lastProductRow).getValues();
    const productNames = [...new Set(productValues.map(row => row[0]).filter(Boolean))];
    Logger.log("getCaseOrderSummaryData: 獲取到 " + productNames.length + " 個不重複的商品名稱。");

    const lastSiteRow = siteSheet.getLastRow();
    if (lastSiteRow < 2) { return { success: true, headers: ["站名"].concat(productNames), summary: {}, sites: [] }; }
    const siteValues = siteSheet.getRange("B2:B" + lastSiteRow).getValues();
    const siteNames = [...new Set(siteValues.map(row => row[0]).filter(Boolean))];
    Logger.log("getCaseOrderSummaryData: 獲取到 " + siteNames.length + " 個不重複的站名。");

    const summaryData = {};
    siteNames.forEach(siteName => {
      summaryData[siteName] = {};
      productNames.forEach(productName => {
        summaryData[siteName][productName] = 0;
      });
    });
    Logger.log("getCaseOrderSummaryData: 已初始化彙總資料結構。");

    if (orderSheet.getLastRow() >= 2) {
      const orderRange = orderSheet.getRange("C2:H" + orderSheet.getLastRow()); 
      const orders = orderRange.getValues();
      Logger.log("getCaseOrderSummaryData: 讀取到 " + orders.length + " 筆訂單進行處理。");

      orders.forEach(function(order, index) {
        const productName = order[0]; 
        const quantity = order[1];    
        let affiliatedSiteRaw = order[5]; 
        let affiliatedSite = "";

        if (affiliatedSiteRaw && typeof affiliatedSiteRaw === 'string') {
          affiliatedSite = affiliatedSiteRaw.split('(')[0].trim();
        }

        if (summaryData[affiliatedSite] && summaryData[affiliatedSite].hasOwnProperty(productName)) {
          if (typeof quantity === 'number' && !isNaN(quantity)) {
            summaryData[affiliatedSite][productName] += quantity;
          }
        }
      });
    }
    Logger.log("getCaseOrderSummaryData: 訂單資料處理完畢。");

    const columnHeaders = ["站名"].concat(productNames);
    
    return {
      success: true,
      headers: columnHeaders,
      summary: summaryData,
      sites: siteNames
    };

  } catch (e) {
    Logger.log("getCaseOrderSummaryData 發生嚴重錯誤: " + e.toString() + "\n" + e.stack);
    return { success: false, error: "產生彙總表時發生伺服器錯誤: " + e.message };
  }
}

/**
 * 【已更新】獲取零星訂單的彙總資料，並加入詳細的偵錯日誌
 */
function getLooseOrderSummaryData() {
  try {
    Logger.log("getLooseOrderSummaryData: 函式開始執行。");

    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const siteSheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID).getSheets()[0];
    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");
    if (!orderSheet) {
      throw new Error("在目標試算表中找不到名為 '使用者訂單(盒)' 的工作表。");
    }
    Logger.log("getLooseOrderSummaryData: 已成功獲取所有需要的試算表。");

    const productValues = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues();
    const productNames = [...new Set(productValues.map(row => row[0]).filter(Boolean))];

    const siteValues = siteSheet.getRange("B2:B" + siteSheet.getLastRow()).getValues();
    const siteNames = [...new Set(siteValues.map(row => row[0]).filter(Boolean))];

    const summaryData = {};
    siteNames.forEach(siteName => {
      summaryData[siteName] = {};
      productNames.forEach(productName => {
        summaryData[siteName][productName] = 0;
      });
    });

    if (orderSheet.getLastRow() >= 2) {
      const orderRange = orderSheet.getRange("C2:F" + orderSheet.getLastRow());
      const orders = orderRange.getValues();
      Logger.log("getLooseOrderSummaryData: 讀取到 " + orders.length + " 筆訂單進行處理。");

      orders.forEach(function(order) {
        const productName = order[0]; // C欄
        const quantity = order[1];    // D欄
        let deliveryLocation = order[3]; // F欄

        if (deliveryLocation && typeof deliveryLocation === 'string') {
          deliveryLocation = deliveryLocation.split('(')[0].trim();
        }

        if (summaryData[deliveryLocation] && summaryData[deliveryLocation].hasOwnProperty(productName)) {
          if (typeof quantity === 'number' && !isNaN(quantity)) {
            summaryData[deliveryLocation][productName] += quantity;
          }
        }
      });
    }
    Logger.log("getLooseOrderSummaryData: 訂單資料處理完畢。");

    const columnHeaders = ["站名"].concat(productNames);
    
    return {
      success: true,
      headers: columnHeaders,
      summary: summaryData,
      sites: siteNames
    };

  } catch (e) {
    console.error("getLooseOrderSummaryData 錯誤: " + e.toString() + " Stack: " + e.stack);
    return { success: false, error: "產生零星彙總表時發生伺服器錯誤: " + e.message };
  }
}

// --- 未來將在此處加入更多後端邏輯 ---
// 例如:
// function getOrderableProducts() { /* ... */ }
// function submitCaseOrder(orderData) { /* ... */ }
// function getUserOrderHistory() { /* ... */ }
