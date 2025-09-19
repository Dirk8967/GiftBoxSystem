// 檔案: OrderHistory.gs

/**
 * 讀取「成箱訂單」紀錄 (僅限當前使用者，依操作人員信箱比對)
 */
function getCaseOrderHistoryForUser() {
  try {
    // 【核心修改 1】直接獲取當前使用者的 Email
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) throw new Error("無法獲取使用者信箱。");

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    
    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues(); // 讀取 A 到 P 欄

    // 【核心修改 2】過濾條件改為比對 N 欄 (索引為13) 的操作人員信箱
    const operatorEmailIndex = 14; // O 欄
    const userOrders = allData.filter(row => 
        row[operatorEmailIndex] && row[operatorEmailIndex].toLowerCase() === currentUserEmail.toLowerCase()
    ).map(row => ({
      // map 的內容保持不變
      uuid: row[0], name: row[1], phone: row[2], productName: row[3],
      quantity: row[4], totalAmount: row[5], location: row[6],
      deliveryDate: row[7] ? new Date(row[7]).toISOString() : null,
      taxidentificationnumber: row[8],
      affiliatedSite: row[9], isPaid: row[10],
      paymentTime: row[11] ? new Date(row[11]).toISOString() : null,
      orderTime: row[12] ? new Date(row[12]).toISOString() : null,
      orderId: row[15]
    }));
    return { success: true, data: userOrders };
  } catch(e) {
    console.error("getCaseOrderHistoryForUser 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 讀取「零星訂單」紀錄 (僅限當前使用者，依操作人員信箱比對)
 */
function getLooseOrderHistoryForUser() {
  try {
    // 【核心修改 1】直接獲取當前使用者的 Email
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) throw new Error("無法獲取使用者信箱。");

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues(); // 讀取 A 到 Q 欄
    
    // 【核心修改 2】過濾條件改為比對 N 欄 (索引為13) 的操作人員信箱
    const operatorEmailIndex = 13; // N 欄
    const userOrders = allData.filter(row => 
        row[operatorEmailIndex] && row[operatorEmailIndex].toLowerCase() === currentUserEmail.toLowerCase()
    ).map(row => ({
      // map 的內容保持不變
      uuid: row[0], name: row[1], phone: row[2], productName: row[3],
      quantity: row[4], totalAmount: row[5], location: row[6],
      isPaid: row[7], paymentTime: row[8] ? new Date(row[8]).toISOString() : null,
      isPickedUp: row[9], pickupTime: row[10] ? new Date(row[10]).toISOString() : null,
      orderTime: row[11] ? new Date(row[11]).toISOString() : null
      // 根據您最新的欄位結構，零星訂單似乎沒有獨立的訂單編號，故此處不回傳
    }));
    return { success: true, data: userOrders };
  } catch(e) {
    console.error("getLooseOrderHistoryForUser 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 3. 更新「成箱訂單」的狀態
 */
function updateCaseOrderStatus(updates) {
  try {
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(箱)");
    const range = sheet.getDataRange();
    const allData = range.getValues();
    headers = allData.shift();

    const adminEmail = Session.getActiveUser().getEmail();
    const adminName = getUserNameByEmail_(adminEmail);
    const adminTime = new Date();

    updates.forEach(update => {
      const rowIndex = allData.findIndex(row => row[0] === update.uuid); // A欄是UUID
      if (rowIndex !== -1) {
        allData[rowIndex][9] = update.isPaid; // J欄: 已繳費
        allData[rowIndex][10] = update.paymentTime ? new Date(update.paymentTime) : null; // K欄: 繳費時間
        // allData[rowIndex][14] = adminName;  // O欄
        // allData[rowIndex][15] = adminEmail; // P欄
        // allData[rowIndex][16] = adminTime;  // Q欄
      }
    });

    range.offset(1, 0, allData.length, allData[0].length).setValues(allData);
    return { success: true };
  } catch(e) { /* ... */ }
}

/**
 * 4. 更新「零星訂單」的狀態
 */
function updateLooseOrderStatus(updates) {
  try {
    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("使用者訂單(盒)");
    // ...與 updateCaseOrderStatus 類似的邏輯 ...
    const range = sheet.getDataRange();
    const allData = range.getValues();
    headers = allData.shift();

    const adminEmail = Session.getActiveUser().getEmail();
    const adminName = getUserNameByEmail_(adminEmail);
    const adminTime = new Date();

    updates.forEach(update => {
      const rowIndex = allData.findIndex(row => row[0] === update.uuid);
      if (rowIndex !== -1) {
        allData[rowIndex][7] = update.isPaid;     // H欄
        allData[rowIndex][8] = update.paymentTime ? new Date(update.paymentTime) : null; // I欄
        allData[rowIndex][9] = update.isPickedUp; // J欄
        allData[rowIndex][10] = update.pickupTime ? new Date(update.pickupTime) : null; // K欄
        // allData[rowIndex][14] = adminName;  // O欄
        // allData[rowIndex][15] = adminEmail; // P欄
        // allData[rowIndex][16] = adminTime;  // Q欄
      }
    });

    range.offset(1, 0, allData.length, allData[0].length).setValues(allData);
    return { success: true };
  } catch(e) { /* ... */ }
}


// 派送紀錄相關程式碼↓↓↓

/**
 * 1. 【新增】根據當前登入者信箱，獲取其被指派的派送紀錄
 * @returns {Object} 包含 { success: boolean, data?: Array<Object>, error?: string } 的物件
 */
function getDeliveryHistoryForCourier() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) {
      throw new Error("無法獲取使用者信箱，請先登入。");
    }

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData.shift(); // 移除標頭
    const emailColumnIndex = 5; // F欄: 派送人員信箱 (索引為5)

    const courierDeliveries = [];
    allData.forEach(row => {
      // 比對 F 欄的派送人員信箱
      if (row[emailColumnIndex] && row[emailColumnIndex].toLowerCase() === currentUserEmail.toLowerCase()) {
        courierDeliveries.push({
          uuid:         row[0], // A
          productName:  row[1], // B
          quantity:     row[2], // C
          location:     row[3], // D
          courierName:  row[4], // E
          courierEmail: row[5], // F
          shippingTime: row[6] ? new Date(row[6]).toISOString() : null, // G
          isCompleted:  row[7] === true // H
        });
      }
    });

    return { success: true, data: courierDeliveries };
  } catch(e) {
    console.error("getDeliveryHistoryForCourier 發生錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 2. 【新增】更新「管理員出貨」工作表中的派送狀態
 * @param {Array<Object>} updates - 從前端傳來，包含多筆已修改紀錄的陣列 [{uuid, shippingTime, isCompleted}]
 */
function updateDeliveryStatus(updates) {
  try {
    if (!updates || updates.length === 0) {
      return { success: true, message: "沒有需要更新的資料。" };
    }

    const sheet = SpreadsheetApp.openById(ORDER_INVENTORY_SHEET_ID).getSheetByName("管理員出貨");
    const range = sheet.getDataRange();
    const allData = range.getValues();
    const headers = allData.shift();

    let updatesMade = 0;
    updates.forEach(update => {
      const rowIndex = allData.findIndex(row => row[0] === update.uuid); // A欄是UUID

      if (rowIndex !== -1) {
        // 更新 G欄(派送時間) 和 H欄(已派送完成)
        allData[rowIndex][6] = update.shippingTime ? new Date(update.shippingTime) : null;
        allData[rowIndex][7] = update.isCompleted;
        updatesMade++;
      }
    });

    if (updatesMade > 0) {
      range.offset(1, 0, allData.length, allData[0].length).setValues(allData);
    }
    
    return { success: true };
  } catch (e) {
    console.error("updateDeliveryStatus 發生錯誤: " + e.toString());
    return { success: false, message: e.message };
  }
}
