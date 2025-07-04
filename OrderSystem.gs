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
    // 產生一個新的 UUID
    const newUuid = Utilities.getUuid();
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

    // 【修正】依照新的 15 欄位格式準備要寫入的資料列 (A 到 O)
    const rowData = [
      newUuid,                    // A欄: UUID
      orderData.name,             // B欄: 訂購人姓名
      phoneValueToWrite,          // C欄: 訂購人電話
      orderData.productName,      // D欄: 商品名稱
      orderData.quantity,         // E欄: 盒數
      orderData.totalPrice,       // F欄: 總計金額
      orderData.location,         // G欄: 寄送地點
      orderData.date,             // H欄: 寄送日期
      orderData.affiliatedSite,   // I欄: 訂單隸屬站點
      false,                      // J欄: 固定填入 FALSE
      '',                         // K欄: 空白
      orderTimestamp,             // L欄: 訂購時間
      operatorName,               // M欄: 操作人員姓名
      operatorEmail,              // N欄: 操作人員信箱
      ''                          // O欄: 空白
    ];

    targetSheet.appendRow(rowData);
    SpreadsheetApp.flush(); 

    const lastRow = targetSheet.getLastRow();
    // 強制設定電話欄位的格式 (C欄，索引為3，位置不變)
    if (phoneStr) {
        targetSheet.getRange(lastRow, 3).setNumberFormat("@"); 
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
    // 1. 產生一個新的 UUID
    const newUuid = Utilities.getUuid();

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
      newUuid,                    // A欄: UUID
      orderData.name,             // B欄: 訂購人姓名
      phoneValueToWrite,          // C欄: 訂購人電話
      orderData.productName,      // D欄: 商品名稱
      orderData.quantity,         // E欄: 盒數
      orderData.totalPrice,       // F欄: 總計金額
      orderData.location,         // G欄: 寄送地點
      false,                      // H欄: 固定填入 FALSE
      '',                         // I欄: 空白
      false,                      // J欄: 固定填入 FALSE
      '',                         // K欄: 空白
      orderTimestamp,             // L欄: 訂購時間
      operatorName,               // M欄: 操作人員姓名
      operatorEmail,              // N欄: 操作人員信箱
      '',                         // O欄: 後台管理員姓名
      '',                         // P欄: 後台管理員信箱
      ''                          // Q欄: 後台管理員操作時間
    ];

    targetSheet.appendRow(rowData);
    SpreadsheetApp.flush(); 

    const lastRow = targetSheet.getLastRow();
    if (phoneStr) {
        targetSheet.getRange(lastRow, 3).setNumberFormat("@"); // 強制 C 欄為文字
    }
    
    return { success: true };
  } catch (e) {
    console.error("submitLooseOrderToServer 錯誤: " + e.toString());
    return { success: false, error: "伺服器處理零星訂單時發生錯誤: " + e.message };
  }
}


/**
 * [含總金額版] 獲取成箱訂單的彙總資料
 */
function getCaseOrderSummaryData() {
  try {
    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const siteSheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID).getSheets()[0];
    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    if (!productSheet || !siteSheet || !orderSheet) throw new Error("無法開啟必要的試算表。");

    const productNames = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);
    const siteNames = siteSheet.getRange("B2:B" + siteSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);

    const summaryData = {};
    const totalAmountBySite = {}; // 【新增】用來儲存各站點總金額的物件
    siteNames.forEach(siteName => {
      summaryData[siteName] = {};
      totalAmountBySite[siteName] = 0; // 【新增】初始化每個站點的金額為 0
      productNames.forEach(productName => {
        summaryData[siteName][productName] = 0;
      });
    });

    if (orderSheet.getLastRow() >= 2) {
      // 讀取範圍擴大到 F 欄，以包含總計金額
      const orders = orderSheet.getRange("D2:I" + orderSheet.getLastRow()).getValues(); // D:商品, E:盒數, F:總金額, I:站點
      orders.forEach(order => {
        const productName = order[0]; // D
        const quantity = order[1];    // E
        const totalAmount = order[2]; // F: 總計金額
        let affiliatedSite = order[5]; // I: 訂單隸屬站點

        if (affiliatedSite && typeof affiliatedSite === 'string') {
          affiliatedSite = affiliatedSite.split('(')[0].trim();
        }

        if (summaryData[affiliatedSite] && summaryData[affiliatedSite][productName] !== undefined) {
          if (typeof quantity === 'number') {
            summaryData[affiliatedSite][productName] += quantity;
          }
          // 【新增】累加總金額
          if (typeof totalAmount === 'number') {
            totalAmountBySite[affiliatedSite] += totalAmount;
          }
        }
      });
    }

    const columnHeaders = ["站名"].concat(productNames);
    
    return {
      success: true,
      headers: columnHeaders,
      summary: summaryData,
      sites: siteNames,
      totalAmounts: totalAmountBySite // 【新增】將計算好的總金額物件一起回傳
    };
  } catch (e) { /* ... */ }
}

/**
 * [修正欄位對應版] 獲取零星訂單的彙總資料，包含訂購、已到貨、配送中三種數量
 * @returns {Object} 包含三種數量彙總結果的物件
 */
function getLooseOrderSummaryData() {
  try {
    // 1. 讀取所有需要的資料來源 (這部分不變)
    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const siteSheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID).getSheets()[0];
    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");
    const shipmentSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");

    if (!productSheet || !siteSheet || !orderSheet || !shipmentSheet) {
      throw new Error("無法開啟必要的試算表。");
    }

    // 2. 獲取欄/列標頭 (這部分不變)
    const productNames = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);
    const siteNames = siteSheet.getRange("B2:B" + siteSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);

    // 【修改點】在初始化資料結構時，一併初始化 totalAmountBySite
    const orderedSummary = {}, arrivedSummary = {}, inTransitSummary = {}, totalAmountBySite = {};
    siteNames.forEach(siteName => {
      orderedSummary[siteName] = {}; arrivedSummary[siteName] = {}; inTransitSummary[siteName] = {};
      totalAmountBySite[siteName] = 0; // 初始化每個站點的金額為 0
      productNames.forEach(productName => {
        orderedSummary[siteName][productName] = 0;
        arrivedSummary[siteName][productName] = 0;
        inTransitSummary[siteName][productName] = 0;
      });
    });

    // 計算「訂購」總數與「總金額」
    if (orderSheet.getLastRow() >= 2) {
      // 【修改點】讀取範圍需要包含 F 欄的「總計金額」
      const orders = orderSheet.getRange("A2:G" + orderSheet.getLastRow()).getValues();
      orders.forEach(order => {
        const productName = order[3]; // D欄: 商品名稱
        const quantity = order[4];    // E欄: 盒數
        const totalAmount = order[5]; // F欄: 總計金額
        let siteName = order[6];      // G欄: 寄送地點
        
        if (siteName && typeof siteName === 'string') siteName = siteName.split('(')[0].trim();
        
        if (orderedSummary[siteName] && orderedSummary[siteName][productName] !== undefined) {
          if (typeof quantity === 'number' && !isNaN(quantity)) {
            orderedSummary[siteName][productName] += quantity;
          }
          // 【新增】累加總金額到對應的站點
          if (typeof totalAmount === 'number' && !isNaN(totalAmount)) {
            totalAmountBySite[siteName] += totalAmount;
          }
        }
      });
    }

    // 5. 計算「已到貨」與「配送中」總數 (這部分不變，因為它讀取的是「管理員出貨」工作表)
    if (shipmentSheet.getLastRow() >= 2) {
      const shipments = shipmentSheet.getRange("A2:H" + shipmentSheet.getLastRow()).getValues();
      shipments.forEach(shipment => {
        const productName = shipment[1]; const quantity = shipment[2]; 
        let siteName = shipment[3];      const isCompleted = shipment[7];
        if (siteName && typeof siteName === 'string') siteName = siteName.split('(')[0].trim();
        if (typeof quantity === 'number' && arrivedSummary[siteName] && arrivedSummary[siteName][productName] !== undefined) {
          if (isCompleted === true) {
            arrivedSummary[siteName][productName] += quantity;
          } else {
            inTransitSummary[siteName][productName] += quantity;
          }
        }
      });
    }

    // 6. 準備最終要回傳給前端的物件 (這部分不變)
    const columnHeaders = ["站名"].concat(productNames);
    return {
      success: true,
      headers: columnHeaders,
      sites: siteNames,
      ordered: orderedSummary,
      arrived: arrivedSummary,
      inTransit: inTransitSummary,
      totalAmounts: totalAmountBySite // 【新增】將計算好的總金額物件一起回傳
    };

  } catch (e) {
    console.error("getLooseOrderSummaryData 發生錯誤: " + e.toString());
    return { success: false, error: "產生彙總表時發生伺服器錯誤: " + e.message };
  }
}

/**
 * 【新增】獲取「總彙總表」資料，合併計算成箱與零星訂單。
 * @returns {Object} 包含所有彙總結果的物件。
 */
function getGrandSummaryData() {
  try {
    // 1. 獲取所有需要的資料來源
    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const siteSheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID).getSheets()[0];
    const caseOrderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    const looseOrderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");

    if (!productSheet || !siteSheet || !caseOrderSheet || !looseOrderSheet) {
      throw new Error("無法開啟必要的資料來源工作表。");
    }

    // 2. 獲取欄/列標頭 (商品與站點)
    const productNames = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);
    const siteNames = siteSheet.getRange("B2:B" + siteSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);

    // 3. 建立空的資料結構來存放「數量」與「金額」的總和
    const summaryData = {};       // 存放[站點][商品]的數量
    const totalAmountBySite = {}; // 存放[站點]的總金額
    siteNames.forEach(siteName => {
      summaryData[siteName] = {};
      totalAmountBySite[siteName] = 0;
      productNames.forEach(productName => {
        summaryData[siteName][productName] = 0;
      });
    });

    // --- 4. 開始累加資料 ---

    // 4.1 處理「成箱訂單」
    if (caseOrderSheet.getLastRow() >= 2) {
      const caseOrders = caseOrderSheet.getRange("A2:I" + caseOrderSheet.getLastRow()).getValues();
      caseOrders.forEach(order => {
        const productName = order[3]; // D欄
        const quantity = order[4];    // E欄
        const totalAmount = order[5]; // F欄
        let affiliatedSite = order[8]; // I欄

        if (affiliatedSite && typeof affiliatedSite === 'string') {
          affiliatedSite = affiliatedSite.split('(')[0].trim();
        }
        if (summaryData[affiliatedSite] && summaryData[affiliatedSite][productName] !== undefined) {
          if (typeof quantity === 'number') summaryData[affiliatedSite][productName] += quantity;
          if (typeof totalAmount === 'number') totalAmountBySite[affiliatedSite] += totalAmount;
        }
      });
    }

    // 4.2 處理「零星訂單」
    if (looseOrderSheet.getLastRow() >= 2) {
      // 根據您的 schema，零星訂單的寄送地點在 G 欄
      const looseOrders = looseOrderSheet.getRange("A2:G" + looseOrderSheet.getLastRow()).getValues();
      looseOrders.forEach(order => {
        const productName = order[3]; // D欄
        const quantity = order[4];    // E欄
        const totalAmount = order[5]; // F欄
        let affiliatedSite = order[6]; // G欄

        if (affiliatedSite && typeof affiliatedSite === 'string') {
          affiliatedSite = affiliatedSite.split('(')[0].trim();
        }
        if (summaryData[affiliatedSite] && summaryData[affiliatedSite][productName] !== undefined) {
          if (typeof quantity === 'number') summaryData[affiliatedSite][productName] += quantity;
          if (typeof totalAmount === 'number') totalAmountBySite[affiliatedSite] += totalAmount;
        }
      });
    }

    // 5. 準備最終要回傳給前端的物件
    const columnHeaders = ["站名"].concat(productNames);
    return {
      success: true,
      headers: columnHeaders,
      sites: siteNames,
      summary: summaryData,
      totalAmounts: totalAmountBySite
    };

  } catch (e) {
    console.error("getGrandSummaryData 發生錯誤: " + e.toString());
    return { success: false, error: "產生總彙總表時發生伺服器錯誤: " + e.message };
  }
}

/**
 * 【新增】讀取所有成箱訂單資料，用於後台管理
 * @returns {Array<Object>} 訂單資料陣列
 */
function getAdminCaseOrders() {
  try {
    // 【偵錯點 1】檢查我們正在使用的 ID
    Logger.log("Attempting to open Spreadsheet with ID: " + ORDER_INVENTORY_SHEET_ID);

    const spreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('使用者訂單(箱)');
    
    // 【偵錯點 2】檢查工作表物件是否存在
    Logger.log("Sheet object found: " + sheet); 

    if (!sheet) {
      throw new Error("工作表 '使用者訂單(箱)' 未找到，請檢查名稱是否完全一致。");
    }

    const lastRow = sheet.getLastRow();
    // 【偵錯點 3】檢查指令碼認定的最後一列是第幾列
    Logger.log("Reported last row number: " + lastRow);

    if (lastRow < 2) {
      Logger.log("判定為沒有資料 (lastRow < 2)，返回空陣列。");
      return [];
    }

    const range = sheet.getRange(2, 1, lastRow - 1, 14);
    const values = range.getValues();
    Logger.log("成功讀取 " + values.length + " 筆資料。");
    
    // ... 後續處理 ...
    const orders = values.map(row => {
      return {
        uuid: row[0],       // A 欄: UUID
        orderTime: row[11], // L 欄: 訂購時間
        name: row[1],       // B 欄: 姓名
        phone: row[2],      // C 欄: 電話
        product: row[3],    // D 欄: 商品
        quantity: row[4],   // E 欄: 盒數
        totalAmount: row[5],// F 欄: 總金額
        location: row[6],   // G 欄: 寄送地點
        deliveryDate: row[7], // H 欄: 寄送日期
        site: row[8],       // I 欄: 隸屬站點
        isPaid: row[9],     // J 欄: 已付款
        paidtime: row[10],  // K 欄: 繳費時間(新增)
        operator: row[12],  // M 欄: 操作人員
        operatoremail: row[13], // N 欄: 操作人員信箱(新增)
        orderId: row[14]    // O 欄: 訂單編號
      };
    });
    return JSON.stringify(orders);

  } catch (e) {
    Logger.log("!!! 嚴重錯誤發生在 getAdminCaseOrders: " + e.toString());
    return { error: true, message: "伺服器內部錯誤: " + e.message };
  }
}


/**
 * 【新增】根據訂單編號刪除一筆訂單
 * @param {string} orderId 要刪除的訂單編號 (N欄)
 * @returns {Object} 包含 success: true 或 false 的物件
 */
function deleteCaseOrder(orderId) {
  try {
    if (!orderId) {
      throw new Error("未提供訂單編號，無法刪除。");
    }
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(箱)');
    if (!sheet) throw new Error("找不到名為 '使用者訂單(箱)' 的工作表");

    const orderIdColumn = 14; // N欄是第14欄
    const allOrderIds = sheet.getRange(1, orderIdColumn, sheet.getLastRow(), 1).getValues();

    let rowToDelete = -1;
    for (let i = 0; i < allOrderIds.length; i++) {
      if (String(allOrderIds[i][0]).trim() === String(orderId).trim()) {
        // 找到匹配的訂單編號，因為資料是從第1列開始讀，所以實際列數是 i + 1
        rowToDelete = i + 1;
        break;
      }
    }

    if (rowToDelete > 1) { // 確保不是標頭列
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "訂單 " + orderId + " 已成功刪除。" };
    } else {
      throw new Error("在試算表中找不到訂單編號: " + orderId);
    }

  } catch (e) {
    console.error("deleteCaseOrder 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}


/**
 * [UUID版] 根據 UUID 刪除一筆訂單紀錄
 * @param {string} uuid 要刪除的紀錄的 UUID (A欄)
 * @returns {Object} 包含 success: true 或 false 的物件
 */
function deleteOrderByUuid(uuid) {
  try {
    if (!uuid) {
      throw new Error("未提供 UUID，無法刪除。");
    }
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(箱)');
    if (!sheet) throw new Error("找不到名為 '使用者訂單(箱)' 的工作表");

    // 【核心修改】現在從第 1 欄 (A欄) 尋找 UUID
    const uuidColumn = 1; 
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();

    let rowToDelete = -1;
    for (let i = 1; i < allUuids.length; i++) {
      if (String(allUuids[i][0]).trim() === String(uuid).trim()) {
        rowToDelete = i + 1;
        break;
      }
    }

    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "紀錄 " + uuid + " 已成功刪除。" };
    } else {
      throw new Error("在試算表中找不到 UUID: " + uuid);
    }

  } catch (e) {
    console.error("deleteOrderByUuid 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * [UUID版] 接收訂單物件並更新試算表中對應的整列資料
 * @param {Object} orderData - 包含所有欄位資訊的訂單物件
 */
function updateOrder(orderData) {
  try {
    const uuid = orderData.uuid; // 【核心修改】改用 uuid 作為 key
    if (!uuid) throw new Error("缺少 UUID，無法更新。");

    // 【核心修改 1】在函式執行時，獲取當前的操作者資訊與時間
    const currentUserEmail = Session.getActiveUser().getEmail();
    const currentUserName = getUserNameByEmail_(currentUserEmail); // 呼叫剛剛建立的輔助函式
    const modificationTime = new Date(); // 當下時間
    const formattedTime = Utilities.formatDate(modificationTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); // 將其格式化

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(箱)');
    if (!sheet) throw new Error("找不到 '使用者訂單(箱)' 工作表");

    // 【核心修改】現在從第 1 欄 (A欄) 尋找 UUID
    const uuidColumn = 1;
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < allUuids.length; i++) {
        if (String(allUuids[i][0]).trim() === String(uuid).trim()) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex === -1) {
        throw new Error("在試算表中找不到訂單編號: " + orderId);
    }
    
    // 準備要寫入的整列資料，現在是 18 欄
    const newRowData = [
        orderData.uuid,       // A: UUID (從傳入的物件中獲取)
        orderData.name,       // B: 姓名
        orderData.phone,      // C: 電話
        orderData.product,    // D: 商品
        orderData.quantity,   // E: 盒數
        orderData.totalAmount,// F: 總金額
        orderData.location,   // G: 寄送地點
        orderData.deliveryDate, // H: 寄送日期
        orderData.site,       // I: 隸屬站點
        orderData.isPaid,     // J: 已付款
        orderData.paidtime,   // K: 付款時間
        orderData.orderTime,  // L: 訂購時間
        orderData.operator,   // M: 操作人員
        orderData.operatoremail, // N: 操作人員信箱
        orderData.orderId,    // O: 訂單編號
        currentUserName,      // P: 【新增】修改者姓名
        currentUserEmail,     // Q: 【新增】修改者信箱
        formattedTime         // R: 【新增】修改時間
    ];

    sheet.getRange(rowIndex, 1, 1, 18).setValues([newRowData]);
    
    return { success: true };
  } catch (e) {
    console.error("updateOrder 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 輔助函式：根據 Email 從使用者授權清單中查詢姓名
 * @param {string} email - 要查詢的電子郵件地址
 * @returns {string} - 找到的姓名，或是在找不到時回傳 Email 本身
 * @private
 */
function getUserNameByEmail_(email) {
  try {
    // 假設您的使用者授權清單在 SPREADSHEET_ID 所指向的試算表的第一個分頁中
    // 且姓名在 A 欄, Email 在 C 欄
    const authSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    const data = authSheet.getRange("A2:C" + authSheet.getLastRow()).getValues();

    for (let i = 0; i < data.length; i++) {
      // row[2] 是 C 欄 (Email), row[0] 是 A 欄 (姓名)
      if (data[i][2].toLowerCase() === email.toLowerCase()) {
        return data[i][0]; // 找到相符的 Email，回傳對應的姓名
      }
    }
    return email; // 如果迴圈跑完都找不到，直接回傳 Email 作為備用
  } catch (e) {
    console.error("getUserNameByEmail_ 發生錯誤: " + e.toString());
    return "查詢姓名失敗"; // 發生錯誤時回傳
  }
}

/**
 * 1. 讀取所有「零星訂單」資料，用於後台管理
 * @returns {string} 包含訂單資料陣列的 JSON 字串
 */
function getAdminLooseOrders() {
  try {
    // 【修改點】目標工作表改為「使用者訂單(盒)」
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(盒)');
    if (!sheet) throw new Error("找不到名為 '使用者訂單(盒)' 的工作表");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);

    const range = sheet.getRange(2, 1, lastRow - 1, 17); // 讀取 A 到 Q 欄
    const values = range.getValues();
    
    // 根據您最新的欄位結構建立物件
    const orders = values.map(row => {
      return {
        uuid:         row[0],  // A
        name:         row[1],  // B
        phone:        row[2],  // C
        product:      row[3],  // D
        quantity:     row[4],  // E
        totalAmount:  row[5],  // F
        location:     row[6],  // G
        isPaid:       row[7],  // H
        paymentTime:  row[8],  // I
        isPickedUp:   row[9],  // J
        pickupTime:   row[10], // K
        orderTime:    row[11], // L
        operatorName: row[12], // M
        operatorEmail:row[13], // N
        adminName:    row[14], // O
        adminEmail:   row[15], // P
        adminTime:    row[16]  // Q
      };
    });

    return JSON.stringify(orders);
  } catch (e) {
    console.error("getAdminLooseOrders 發生錯誤: " + e.toString());
    return JSON.stringify({ error: true, message: e.message });
  }
}

/**
 * 2. 根據 UUID 刪除一筆「零星訂單」紀錄
 * @param {string} uuid 要刪除的紀錄的 UUID (A欄)
 */
function deleteLooseOrderByUuid(uuid) {
  try {
    if (!uuid) throw new Error("未提供 UUID，無法刪除。");
    // 【修改點】目標工作表改為「使用者訂單(盒)」
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(盒)');
    // ... 後續刪除邏輯與 deleteOrderByUuid 完全相同 ...
    const uuidColumn = 1; 
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowToDelete = -1;
    for (let i = 1; i < allUuids.length; i++) {
      if (String(allUuids[i][0]).trim() === String(uuid).trim()) {
        rowToDelete = i + 1;
        break;
      }
    }
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "紀錄 " + uuid + " 已成功刪除。" };
    } else {
      throw new Error("在試算表中找不到 UUID: " + uuid);
    }
  } catch (e) {
    console.error("deleteLooseOrderByUuid 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 3. 更新一筆「零星訂單」的整列資料
 * @param {Object} orderData - 包含所有欄位資訊的訂單物件
 */
function updateLooseOrder(orderData) {
  try {
    const uuid = orderData.uuid;
    if (!uuid) throw new Error("缺少 UUID，無法更新。");
    // 【修改點】目標工作表改為「使用者訂單(盒)」
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName('使用者訂單(盒)');
    // ... 找到 rowIndex 的邏輯與 updateOrder 完全相同 ...
    // 管理員資訊寫入
    const currentUserEmail = Session.getActiveUser().getEmail();
    const currentUserName = getUserNameByEmail_(currentUserEmail); // 呼叫剛剛建立的輔助函式
    const modificationTime = new Date(); // 當下時間
    const formattedTime = Utilities.formatDate(modificationTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); // 將其格式化
    orderData.adminName = currentUserName;
    orderData.adminEmail = currentUserEmail;
    orderData.adminTime = modificationTime;


    const uuidColumn = 1;
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < allUuids.length; i++) {
        if (String(allUuids[i][0]).trim() === String(uuid).trim()) {
            rowIndex = i + 1;
            break;
        }
    }
    if (rowIndex === -1) throw new Error("在試算表中找不到 UUID: " + uuid);
    
    // 【修改點】寫入的資料陣列結構更新為 17 欄
    const newRowData = [
        orderData.uuid, orderData.name, orderData.phone, orderData.product,
        orderData.quantity, orderData.totalAmount, orderData.location,
        orderData.isPaid, orderData.paymentTime, orderData.isPickedUp,
        orderData.pickupTime, orderData.orderTime, orderData.operatorName,
        orderData.operatorEmail, orderData.adminName, orderData.adminEmail,
        orderData.adminTime
    ];

    sheet.getRange(rowIndex, 1, 1, 17).setValues([newRowData]);
    
    return { success: true };
  } catch (e) {
    console.error("updateLooseOrder 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * [修正欄位結構版] 根據商品名稱和站點名稱，查詢「管理員出貨」工作表的詳細資料。
 * @param {string} productName - 要查詢的商品名稱。
 * @param {string} siteName - 要查詢的站點名稱。
 * @returns {Object} 包含已出貨總數、配送中總數和詳細訂單物件列表的物件。
 */
function getShipmentDetails(productName, siteName) {
  try {
    const shipmentSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    if (!shipmentSheet) throw new Error("找不到名為 '管理員出貨' 的工作表");

    const lastRow = shipmentSheet.getLastRow();
    if (lastRow < 2) {
        return { success: true, shipped: 0, inTransit: 0, details: [] };
    }
    // 【核心修改】讀取範圍改為 A 到 K (共 11 欄)
    const allData = shipmentSheet.getRange(2, 1, lastRow - 1, 11).getValues(); 

    let shippedTotal = 0;
    let inTransitTotal = 0;
    const detailedRows = [];

    allData.forEach(row => {
      const rowProductName = row[1]; // B欄: 商品名稱
      const rowSiteName = row[3];    // D欄: 寄送地點
      
      if (rowProductName === productName && rowSiteName && rowSiteName.includes(siteName)) {
        // 建立物件時，對應到正確的欄位
        const detailObject = {
            uuid:          row[0], // A
            productName:   row[1], // B
            quantity:      row[2], // C
            location:      row[3], // D
            courierName:   row[4], // E
            courierEmail:  row[5], // F
            shippingTime:  row[6] ? new Date(row[6]).toISOString() : null, // G
            isCompleted:   row[7]  // H
            // I, J, K 欄的管理員資訊不需要在詳細列表中顯示，所以此處不加入
        };
        detailedRows.push(detailObject);
        
        const quantity = detailObject.quantity;
        const isCompleted = detailObject.isCompleted;

        if (isCompleted === true || String(isCompleted).toUpperCase() === 'V' || isCompleted === '是') {
          shippedTotal += quantity;
        } else {
          inTransitTotal += quantity;
        }
      }
    });

    return { success: true, shipped: shippedTotal, inTransit: inTransitTotal, details: detailedRows };
  } catch (e) {
    console.error("getShipmentDetails 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 【新增】根據 UUID 刪除一筆「管理員出貨」紀錄
 * @param {string} uuid 要刪除的紀錄的 UUID (A欄)
 */
function deleteShipmentByUuid(uuid) {
  try {
    if (!uuid) throw new Error("未提供 UUID，無法刪除。");
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    if (!sheet) throw new Error("找不到名為 '管理員出貨' 的工作表");

    const uuidColumn = 1; // A欄
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowToDelete = -1;
    for (let i = 1; i < allUuids.length; i++) {
      if (String(allUuids[i][0]) === String(uuid)) {
        rowToDelete = i + 1;
        break;
      }
    }
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true };
    } else {
      throw new Error("在 '管理員出貨' 表中找不到 UUID: " + uuid);
    }
  } catch (e) {
    console.error("deleteShipmentByUuid 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * [修正欄位結構版] 更新一筆「管理員出貨」紀錄，並寫入操作者資訊
 * @param {Object} shipmentData - 從前端傳來的出貨紀錄物件
 */
function updateShipment(shipmentData) {
  try {
    const uuid = shipmentData.uuid;
    if (!uuid) throw new Error("缺少 UUID，無法更新。");

    const adminEmail = Session.getActiveUser().getEmail();
    const adminName = getUserNameByEmail_(adminEmail);
    const adminTime = new Date();

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    if (!sheet) throw new Error("找不到名為 '管理員出貨' 的工作表");

    // ... 找到 rowIndex 的邏輯保持不變 ...
    const uuidColumn = 1;
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < allUuids.length; i++) {
        if (String(allUuids[i][0]) === String(uuid)) {
            rowIndex = i + 1;
            break;
        }
    }
    if (rowIndex === -1) throw new Error("找不到 UUID: " + uuid);
    
    // 【核心修改】準備要寫入的整列資料，現在是 11 欄，對應 A-K
    const newRowData = [
      shipmentData.uuid,          // A: UUID
      shipmentData.productName,   // B: 商品名稱
      shipmentData.quantity,      // C: 盒數
      shipmentData.location,      // D: 寄送地點
      shipmentData.courierName,   // E: 派送人員姓名
      shipmentData.courierEmail,  // F: 派送人員信箱
      shipmentData.shippingTime ? new Date(shipmentData.shippingTime) : null, // G: 派送時間
      shipmentData.isCompleted,   // H: 已派送完成
      adminName,                  // I: 後台管理員姓名
      adminEmail,                 // J: 後台管理員信箱
      adminTime                   // K: 後台管理員操作時間
    ];

    // 【核心修改】寫入範圍改為 11 欄
    sheet.getRange(rowIndex, 1, 1, 11).setValues([newRowData]);
    
    return { success: true };
  } catch (e) {
    console.error("updateShipment 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 1. 【新增】獲取所有可用的派送人員名單
 * 從「使用者授權清單」工作表讀取。
 * @returns {Array<Object>} 包含 {name: string, email: string} 的物件陣列。
 */
function getCourierList() {
  try {
    // 假設人員名單在主試算表的第一個分頁「訂購網頁權限」中
    const authSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("訂購網頁權限");
    if (!authSheet) throw new Error("找不到 '訂購網頁權限' 工作表");

    const lastRow = authSheet.getLastRow();
    if (lastRow < 2) return [];

    const data = authSheet.getRange("A2:C" + lastRow).getValues(); // 讀取 A:姓名, C:Email
    
    const couriers = data.map(row => ({
      name: row[0],  // A欄
      email: row[2]  // C欄
    })).filter(c => c.name && c.email); // 過濾掉沒有姓名或Email的資料

    return couriers;
  } catch(e) {
    console.error("getCourierList 發生錯誤: " + e.toString());
    return []; // 發生錯誤時回傳空陣列
  }
}


/**
 * 2. 【新增】建立一筆新的「管理員出貨」紀錄
 * @param {Object} shipmentData - 從前端傳來的新出貨資訊
 */
function addNewShipment(shipmentData) {
  try {
    if (!shipmentData || !shipmentData.productName || !shipmentData.courierName) {
      throw new Error("缺少必要的出貨資訊。");
    }

    // A. 自動產生欄位
    const uuid = Utilities.getUuid();
    const adminEmail = Session.getActiveUser().getEmail();
    const adminName = getUserNameByEmail_(adminEmail);
    const adminTime = new Date();
    
    // B. 根據派送員姓名，查詢其 Email
    const courierList = getCourierList();
    const selectedCourier = courierList.find(c => c.name === shipmentData.courierName);
    const courierEmail = selectedCourier ? selectedCourier.email : '';

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    if (!sheet) throw new Error("找不到 '管理員出貨' 工作表");
    
    // C. 準備要寫入的整列資料，共 11 欄
    const newRowData = [
      uuid,                     // A: UUID
      shipmentData.productName,   // B: 商品名稱
      shipmentData.quantity,      // C: 盒數
      shipmentData.location,      // D: 寄送地點
      shipmentData.courierName,   // E: 派送人員姓名
      courierEmail,               // F: 派送人員信箱
      null,                       // G: 派送時間 (空白)
      false,                      // H: 已派送完成 (否)
      adminName,                  // I: 後台管理員姓名
      adminEmail,                 // J: 後台管理員信箱
      adminTime                   // K: 後台管理員操作時間
    ];
    
    sheet.appendRow(newRowData);
    return { success: true };

  } catch (e) {
    console.error("addNewShipment 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 【新增】獲取庫存管理頁面所需的所有彙總資料。
 * @returns {Object} 包含所有產品庫存狀態的物件。
 */
function getInventoryManagementData() {
  try {
    // 1. 獲取所有需要的資料來源
    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const orderSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");
    const purchaseSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員進貨");
    const shipmentSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");

    if (!productSheet || !orderSheet || !purchaseSheet || !shipmentSheet) {
      throw new Error("找不到必要的資料工作表。");
    }

    // 2. 建立一個以「商品名稱」為 key 的資料結構，用來存放計算結果
    const inventoryData = {};
    const productNames = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues()
      .map(row => row[0]).filter(Boolean);
    
    productNames.forEach(name => {
      inventoryData[name] = {
        totalOrdered: 0,
        notArrived: 0,
        arrived: 0,
        shipped: 0,
        actualInventory: 0
      };
    });

    // 3. 計算「各站訂購總數」
    if (orderSheet.getLastRow() > 1) {
      const orders = orderSheet.getRange("D2:E" + orderSheet.getLastRow()).getValues(); // D:商品, E:盒數
      orders.forEach(order => {
        const productName = order[0];
        const quantity = order[1];
        if (inventoryData[productName] && typeof quantity === 'number') {
          inventoryData[productName].totalOrdered += quantity;
        }
      });
    }

    // 4. 計算「已向廠商訂購」的 (未到貨 / 已到貨)
    if (purchaseSheet.getLastRow() > 1) {
      const purchases = purchaseSheet.getRange("B2:D" + purchaseSheet.getLastRow()).getValues(); // B:商品, C:盒數, D:已到貨
      purchases.forEach(purchase => {
        const productName = purchase[0];
        const quantity = purchase[1];
        const isArrived = purchase[2];
        if (inventoryData[productName] && typeof quantity === 'number') {
          if (isArrived === true) {
            inventoryData[productName].arrived += quantity;
          } else {
            inventoryData[productName].notArrived += quantity;
          }
        }
      });
    }
    
    // 5. 計算「已出貨」總數
    if (shipmentSheet.getLastRow() > 1) {
        const shipments = shipmentSheet.getRange("B2:C" + shipmentSheet.getLastRow()).getValues(); // B:商品, C:盒數
        shipments.forEach(shipment => {
            const productName = shipment[0];
            const quantity = shipment[1];
            if(inventoryData[productName] && typeof quantity === 'number') {
                inventoryData[productName].shipped += quantity;
            }
        });
    }

    // 6. 計算「實際庫存」並將物件轉為陣列，方便前端渲染
    const result = productNames.map(name => {
      const data = inventoryData[name];
      data.actualInventory = data.arrived - data.shipped;
      return { productName: name, ...data };
    });

    return { success: true, data: result };

  } catch (e) {
    console.error("getInventoryManagementData 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}


/**
 * 1. 【新增】根據商品名稱，獲取所有相關的「管理員進貨」紀錄
 * @param {string} productName - 要查詢的商品名稱
 * @returns {Object} 包含進貨紀錄陣列的物件
 */
function getPurchaseRecords(productName) {
  try {
    if (!productName) throw new Error("未提供商品名稱。");

    const purchaseSheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員進貨");
    if (!purchaseSheet) throw new Error("找不到 '管理員進貨' 工作表");

    const lastRow = purchaseSheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const allData = purchaseSheet.getRange(2, 1, lastRow - 1, 11).getValues(); // A-K
    const filteredData = [];

    allData.forEach(row => {
      // B欄是商品名稱
      if (row[1] === productName) {
        filteredData.push({
          uuid:         row[0], // A
          productName:  row[1], // B
          quantity:     row[2], // C
          isArrived:    row[3], // D
          arrivedTime:  row[4] ? new Date(row[4]).toISOString() : null, // E
          purchaseTime: row[5] ? new Date(row[5]).toISOString() : null, // F
          operatorName: row[6], // G
          operatorEmail:row[7]  // H
        });
      }
    });

    return { success: true, data: filteredData };
  } catch (e) {
    console.error("getPurchaseRecords 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}


/**
 * 2. 【新增】根據 UUID 刪除一筆「管理員進貨」紀錄
 * @param {string} uuid - 要刪除的紀錄的 UUID (A欄)
 */
function deletePurchaseRecord(uuid) {
  try {
    if (!uuid) throw new Error("未提供 UUID，無法刪除。");
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員進貨");
    if (!sheet) throw new Error("找不到 '管理員進貨' 工作表");

    const uuidColumn = 1; // A欄
    const allUuids = sheet.getRange(1, uuidColumn, sheet.getLastRow(), 1).getValues();
    let rowToDelete = -1;
    for (let i = 1; i < allUuids.length; i++) {
      if (String(allUuids[i][0]) === String(uuid)) {
        rowToDelete = i + 1;
        break;
      }
    }

    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true };
    } else {
      throw new Error("在 '管理員進貨' 表中找不到 UUID: " + uuid);
    }
  } catch (e) {
    console.error("deletePurchaseRecord 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 【可編輯更多欄位版】批次更新多筆「管理員進貨」紀錄
 * @param {Array<Object>} records - 從前端傳來，包含多筆已修改紀錄的陣列
 */
function updatePurchaseRecords(records) {
  try {
    if (!records || records.length === 0) {
      return { success: true, message: "沒有需要更新的資料。" };
    }

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員進貨");
    if (!sheet) throw new Error("找不到 '管理員進貨' 工作表");

    const range = sheet.getDataRange();
    const allData = range.getValues();
    const headers = allData.shift();

    const adminEmail = Session.getActiveUser().getEmail();
    const adminName = getUserNameByEmail_(adminEmail);
    const adminTime = new Date();

    let updatesMade = 0;
    records.forEach(recordToUpdate => {
      const rowIndex = allData.findIndex(row => row[0] === recordToUpdate.uuid);

      if (rowIndex !== -1) {
        // 【核心修改】除了原本的欄位，現在也更新盒數和訂購時間
        allData[rowIndex][2] = recordToUpdate.quantity;    // C欄: 盒數
        allData[rowIndex][3] = recordToUpdate.isArrived;   // D欄: 已到貨
        allData[rowIndex][4] = recordToUpdate.arrivedTime ? new Date(recordToUpdate.arrivedTime) : null; // E欄: 到貨時間
        allData[rowIndex][5] = recordToUpdate.purchaseTime ? new Date(recordToUpdate.purchaseTime) : null; // F欄: 訂購時間

        // 更新管理員操作紀錄
        allData[rowIndex][8] = adminName;
        allData[rowIndex][9] = adminEmail;
        allData[rowIndex][10] = adminTime;
        updatesMade++;
      }
    });

    if (updatesMade > 0) {
      range.offset(1, 0, allData.length, allData[0].length).setValues(allData);
    }
    
    return { success: true };

  } catch (e) {
    console.error("updatePurchaseRecords 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}


/**
 * 2. 【確認此版本】建立一筆新的「管理員進貨」紀錄
 */
function addNewPurchaseRecord(recordData) {
  try {
    if (!recordData || !recordData.productName || !recordData.purchaseTime) {
      throw new Error("缺少必要的進貨資訊。");
    }

    const uuid = Utilities.getUuid();
    const operatorEmail = Session.getActiveUser().getEmail(); // 訂購的操作人員即為當前使用者
    const operatorName = getUserNameByEmail_(operatorEmail);

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員進貨");
    if (!sheet) throw new Error("找不到 '管理員進貨' 工作表");
    
    // 根據您最新的 schema (UUID, 商品, 盒數, 已到貨, 到貨時間, 訂購時間, 操作員, 操作員信箱)
    const newRowData = [
      uuid,
      recordData.productName,
      recordData.quantity,
      recordData.isArrived,
      recordData.arrivedTime ? new Date(recordData.arrivedTime) : null,
      new Date(recordData.purchaseTime),
      operatorName,
      operatorEmail,
      '', // 後台管理員姓名 (留空)
      '', // 後台管理員信箱 (留空)
      ''  // 後台管理員操作時間 (留空)
    ];
    
    sheet.appendRow(newRowData);
    return { success: true };
  } catch(e) {
    console.error("addNewPurchaseRecord 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * 【新增】獲取「總彙總表」所需的全部資料，在後端完成所有計算。
 * @returns {Object} 包含表頭和最終彙總資料的物件。
 */
function getGrandSummaryReportData() {
  try {
    // 1. 獲取所有需要的資料來源
    const userAuthSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const orderInventorySpreadsheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID);
    
    const productSheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID).getSheets()[0];
    const unitSheet = userAuthSpreadsheet.getSheetByName("營業處單位資訊");
    const userAuthSheet = userAuthSpreadsheet.getSheetByName("訂購網頁權限");
    const caseOrderSheet = orderInventorySpreadsheet.getSheetByName("使用者訂單(箱)");
    const looseOrderSheet = orderInventorySpreadsheet.getSheetByName("使用者訂單(盒)");

    // 2. 準備基礎資料和對應表
    const productNames = productSheet.getRange("B2:B" + productSheet.getLastRow()).getValues().map(r => r[0]).filter(Boolean);
    const unitDataRows = unitSheet.getRange("B2:D" + unitSheet.getLastRow()).getValues(); // B:單位名稱, D:目標業績
    
    const userAuthRows = userAuthSheet.getRange("C2:D" + userAuthSheet.getLastRow()).getValues(); // C:Email, D:隸屬單位

    const userToUnitMap = {}; // { email: unitName }
    userAuthRows.forEach(r => { userToUnitMap[r[0]] = r[1]; });
    
    const unitTargets = {}; // { unitName: target }
    const unitNames = [];   // [unitName1, unitName2, ...]
    unitDataRows.forEach(r => {
        const unitName = r[0];
        unitNames.push(unitName);
        unitTargets[unitName] = parseFloat(r[2]) || 0; // D欄是目標業績
    });

    // 3. 初始化資料結構
    const summaryData = {};
    const totalAmountByUnit = {};
    unitNames.forEach(unitName => {
      summaryData[unitName] = {};
      totalAmountByUnit[unitName] = 0;
      productNames.forEach(productName => {
        summaryData[unitName][productName] = 0;
      });
    });

    // 4. 彙總計算 (與之前本地檔案的邏輯相同)
    if (caseOrderSheet.getLastRow() > 1) {
      const caseOrders = caseOrderSheet.getRange("A2:O" + caseOrderSheet.getLastRow()).getValues(); // 讀取到訂單編號
      caseOrders.forEach(order => {
        const operatorEmail = order[13]; // N欄
        const unitName = userToUnitMap[operatorEmail];
        if (unitName) {
          const productName = order[3]; // D欄
          const quantity = order[4];    // E欄
          const totalAmount = order[5]; // F欄
          if (summaryData[unitName] && summaryData[unitName][productName] !== undefined) {
            if (typeof quantity === 'number') summaryData[unitName][productName] += quantity;
            if (typeof totalAmount === 'number') totalAmountByUnit[unitName] += totalAmount;
          }
        }
      });
    }

    if (looseOrderSheet.getLastRow() > 1) {
      const looseOrders = looseOrderSheet.getRange("A2:N" + looseOrderSheet.getLastRow()).getValues(); // 讀取到操作人員信箱
      looseOrders.forEach(order => {
        const operatorEmail = order[13]; // N欄
        const unitName = userToUnitMap[operatorEmail];
        if (unitName) {
          const productName = order[3]; // D欄
          const quantity = order[4];    // E欄
          const totalAmount = order[5]; // F欄
          if (summaryData[unitName] && summaryData[unitName][productName] !== undefined) {
            if (typeof quantity === 'number') summaryData[unitName][productName] += quantity;
            if (typeof totalAmount === 'number') totalAmountByUnit[unitName] += totalAmount;
          }
        }
      });
    }
    
    // 5. 準備最終要回傳給前端的物件
    return { 
      success: true, 
      productHeaders: productNames, 
      unitData: unitNames.map(name => ({
          unitName: name,
          products: summaryData[name],
          totalAmount: totalAmountByUnit[name],
          target: unitTargets[name]
      }))
    };

  } catch (e) {
    console.error("getGrandSummaryReportData 發生錯誤: " + e.toString());
    return { success: false, error: "產生報表時發生伺服器錯誤。" };
  }
}
