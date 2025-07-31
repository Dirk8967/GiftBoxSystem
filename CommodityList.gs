// ================================================================= //
//                       商品管理系統 - 後端邏輯
// ================================================================= //

const PRODUCT_SHEET_ID = "12lnNMF0Nu8enQjB1LTEKbhvI4faClKKg-iC1fp9aeE0";
const PRODUCT_IMAGE_FOLDER_NAME = "商品圖片"; 
const PRODUCT_CACHE_KEY = 'allProductsData_v5'; // 更新快取鍵
const PRODUCT_CACHE_DURATION_SECONDS = 300; 

/**
 * 輔助函式：獲取商品資訊工作表物件
 */
function getProductSheet_() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PRODUCT_SHEET_ID);
    if (!spreadsheet) {
        throw new Error("無法透過ID開啟商品資訊試算表。");
    }
    const sheet = spreadsheet.getSheets()[0]; 
    if (!sheet) {
        throw new Error("商品資訊試算表中找不到第一個工作表。");
    }
    Logger.log("getProductSheet_: 成功獲取工作表 '" + sheet.getName() + "'");
    return sheet;
  } catch (e) {
    console.error("getProductSheet_ 捕捉到嚴重錯誤: " + e.toString());
    throw new Error("讀取商品試算表設定時發生內部錯誤: " + e.message);
  }
}


/**
 * 輔助函式：【已修改】直接透過 ID 獲取已分享的商品圖片儲存資料夾
 * @returns {GoogleAppsScript.Drive.Folder} 資料夾物件
 */
function getProductImageFolder_() {
  // --- 請將 'YOUR_SHARED_FOLDER_ID_HERE' 替換為您實際的「商品圖片」資料夾 ID ---
  const SHARED_IMAGE_FOLDER_ID = "1Wo-L_92LHxmZo64OTtHvZbq3pCLd65ET"; 

  try {
    const imageFolder = DriveApp.getFolderById(SHARED_IMAGE_FOLDER_ID);
    // 為了確保萬無一失，可以檢查一下資料夾是否存在
    if (!imageFolder) {
       throw new Error("無法透過指定的 ID 找到「商品圖片」資料夾。請檢查 ID 是否正確，且執行者是否有權限存取。");
    }
    return imageFolder;
  } catch (e) {
    console.error("getProductImageFolder_ 錯誤: " + e.toString());
    // 拋出一個對使用者更友善的錯誤訊息
    throw new Error("伺服器錯誤：無法存取商品圖片儲存位置。請聯繫管理員確認資料夾設定。");
  }
}
/**
 * 輔助函式：獲取或建立商品圖片儲存資料夾
 */
// 原先以試算表ID為定位存取在同一個資料夾中的子資料夾
// function getProductImageFolder_() {
//   try {
//     const productSheetFile = DriveApp.getFileById(PRODUCT_SHEET_ID);
//     const parentFolders = productSheetFile.getParents();
//     if (!parentFolders.hasNext()) {
//         throw new Error("商品資訊試算表沒有父資料夾，無法定位圖片儲存位置。");
//     }
//     const parentFolder = parentFolders.next(); 
    
//     let imageFolder;
//     const folders = parentFolder.getFoldersByName(PRODUCT_IMAGE_FOLDER_NAME);
//     if (folders.hasNext()) {
//       imageFolder = folders.next();
//     } else {
//       imageFolder = parentFolder.createFolder(PRODUCT_IMAGE_FOLDER_NAME);
//       Logger.log("已建立資料夾: '" + PRODUCT_IMAGE_FOLDER_NAME + "' 於資料夾ID '" + parentFolder.getId() + "'");
//     }
//     return imageFolder;
//   } catch (e) {
//     console.error("getProductImageFolder_ 錯誤: " + e.toString());
//     throw new Error("伺服器錯誤：處理商品圖片資料夾失敗: " + e.message);
//   }
// }



/**
 * 輔助函式：清除商品資料快取
 */
function clearProductsCache_() {
  try {
    CacheService.getScriptCache().remove(PRODUCT_CACHE_KEY);
    Logger.log("商品資料快取已清除 (Key: " + PRODUCT_CACHE_KEY + ")。");
  } catch (e) {
    console.error('清除商品快取失敗: ' + e.toString());
  }
}

/**
 * 輔助函式：處理圖片上傳並回傳 File ID
 */
function uploadImageToDrive_(base64Data, fileName, mimeType) {
  try {
    const imageFolder = getProductImageFolder_();
    const pureBase64Data = base64Data.substring(base64Data.indexOf(',') + 1);
    const decodedData = Utilities.base64Decode(pureBase64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    const uniqueFileName = new Date().getTime() + "_" + fileName.replace(/[^a-zA-Z0-9._-]/g, '');
    const file = imageFolder.createFile(blob.setName(uniqueFileName));
    Logger.log("圖片已上傳: " + file.getName() + ", ID: " + file.getId());
    return file.getId();
  } catch (e) {
    console.error("uploadImageToDrive_ 錯誤 (檔案: " + fileName + "): " + e.toString());
    throw new Error("圖片 '" + fileName + "' 上傳至雲端硬碟失敗。");
  }
}

/**
 * 輔助函式：刪除 Drive 中的檔案 (移至垃圾桶)
 */
function deleteDriveFile_(fileId) {
  if (!fileId || typeof fileId !== 'string' || fileId.trim() === '') {
      Logger.log("deleteDriveFile_: 無效或空的 fileId: '" + fileId + "'");
      return;
  }
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true); 
    Logger.log("檔案 ID '" + fileId + "' 已移至垃圾桶。");
  } catch (e) {
    console.warn("刪除 Drive 檔案 ID '" + fileId + "' 失敗: " + e.toString() + ". 可能檔案已不存在或權限問題。");
  }
}

/**
 * 獲取所有商品資料 (包含詳細日誌)
 * @returns {Array<Object>} 商品資料陣列
 */
function getProductListData() {
  // 快取邏輯暫時停用以方便偵錯
  // const cache = CacheService.getScriptCache();
  // const cached = cache.get(PRODUCT_CACHE_KEY);
  // if (cached != null) { /* ... */ }

  Logger.log('getProductListData: 函式開始執行。');
  try {
    const sheet = getProductSheet_();
    
    const lastRow = sheet.getLastRow();
    Logger.log("getProductListData: 工作表 '" + sheet.getName() + "' 的最後一列是: " + lastRow);

    if (lastRow < 2) { 
      Logger.log("getProductListData: 工作表資料不足，將回傳空陣列。");
      return []; 
    }

    const numDataRows = lastRow - 1;
    const rangeToRead = "A2:H" + lastRow; // 讀取 A 到 H 欄
    Logger.log("getProductListData: 準備讀取範圍: " + rangeToRead);
    
    const values = sheet.getRange(rangeToRead).getValues(); 
    Logger.log("getProductListData: 成功讀取到 " + values.length + " 列原始資料。");
    
    const products = values.map(function(row, index) {
      // 轉換每一列資料
      return {
        rowNumber: index + 2,
        companyName: String(row[0] || ''),      // A欄
        productName: String(row[1] || ''),      // B欄
        sellingPricePerBox: parseFloat(row[2]) || 0, // C欄
        casePricePerBox: parseFloat(row[3]) || 0,    // D欄
        totalCasePrice: parseFloat(row[4]) || 0,     // E欄
        boxesPerCase: parseInt(row[5]) || 0,         // F欄
        canShipAboveCase: row[6] === true,           // G欄
        imageId: String(row[7] || '')                // H欄
      };
    });
    
    Logger.log("getProductListData: 資料轉換完成，準備回傳 " + products.length + " 筆商品資料。");
    // cache.put(PRODUCT_CACHE_KEY, JSON.stringify(products), PRODUCT_CACHE_DURATION_SECONDS);
    return products;
  } catch (e) {
    Logger.log("getProductListData 發生嚴重錯誤: " + e.toString() + "\n" + e.stack);
    console.error('讀取商品資料失敗 (getProductListData): ' + e.toString());
    // 拋出錯誤讓前端的 withFailureHandler 捕捉
    throw new Error('伺服器在讀取商品資料時發生錯誤：' + e.message);
  }
}

/**
 * 【已更新】新增商品資料 (包含圖片上傳) - 8 欄
 * @param {Object} productData - { companyName, productName, sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase, canShipAboveCase }
 * @param {Object} imageInfo - (可選) { base64Data, fileName, mimeType }
 * @returns {Object} { success: boolean, error?: string }
 */
function addProductData(productData, imageInfo) {
  try {
    const { companyName, productName, sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase, canShipAboveCase } = productData;
    if (!companyName || !productName || String(sellingPricePerBox).trim() === '' || String(boxesPerCase).trim() === '') { // 成箱盒數也設為必填
      throw new Error("公司名稱、商品名稱、售價/盒和成箱盒數為必填。");
    }
    const numFields = {sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase}; // boxesPerCase 也加入數字驗證
    for (const key in numFields) {
        if (String(numFields[key]).trim() !== '' && isNaN(parseFloat(numFields[key]))) {
            throw new Error("價格與數量相關欄位必須是有效的數字或空白（成箱價/總價可為空）。");
        }
    }

    let imageFileId = '';
    if (imageInfo && imageInfo.base64Data && imageInfo.fileName && imageInfo.mimeType) {
      imageFileId = uploadImageToDrive_(imageInfo.base64Data, imageInfo.fileName, imageInfo.mimeType);
    }

    const sheet = getProductSheet_();
    sheet.appendRow([
      companyName.trim(),
      productName.trim(),
      parseFloat(sellingPricePerBox), // 確保是數字
      String(casePricePerBox).trim() === '' ? null : parseFloat(casePricePerBox),
      String(totalCasePrice).trim() === '' ? null : parseFloat(totalCasePrice),
      parseInt(boxesPerCase), // 確保是數字
      canShipAboveCase === true,
      imageFileId
    ]);
    SpreadsheetApp.flush();
    clearProductsCache_();
    return { success: true };
  } catch (e) {
    console.error("新增商品失敗 (商品: " + productData.productName + "): " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 【已更新】更新商品資料 (包含圖片更新) - 8 欄
 * @param {Object} productData - { rowNumber, companyName, productName, sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase, canShipAboveCase, existingImageId, deleteCurrentImage }
 * @param {Object} imageInfo - (可選) { base64Data, fileName, mimeType } 表示新圖片
 * @returns {Object} { success: boolean, error?: string }
 */
function updateProductData(productData, imageInfo) {
  try {
    const { rowNumber, companyName, productName, sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase, canShipAboveCase, existingImageId, deleteCurrentImage } = productData;
    if (!rowNumber || !companyName || !productName || String(sellingPricePerBox).trim() === '' || String(boxesPerCase).trim() === '') {
      throw new Error("缺少必要更新資訊 (列號、公司名稱、商品名稱、售價/盒、成箱盒數)。");
    }
    const numFields = {sellingPricePerBox, casePricePerBox, totalCasePrice, boxesPerCase};
    for (const key in numFields) {
        if (String(numFields[key]).trim() !== '' && isNaN(parseFloat(numFields[key]))) {
            throw new Error("價格與數量相關欄位必須是有效的數字或空白（成箱價/總價可為空）。");
        }
    }

    let newImageFileId = existingImageId || '';
    let oldImageToDelete = null;

    if (imageInfo && imageInfo.base64Data && imageInfo.fileName && imageInfo.mimeType) {
      newImageFileId = uploadImageToDrive_(imageInfo.base64Data, imageInfo.fileName, imageInfo.mimeType);
      if (existingImageId && existingImageId !== newImageFileId) {
        oldImageToDelete = existingImageId; 
      }
    } else if (deleteCurrentImage === true && existingImageId) {
        newImageFileId = ''; 
        oldImageToDelete = existingImageId;
    }

    const sheet = getProductSheet_();
    sheet.getRange(rowNumber, 1, 1, 8).setValues([[ // 更新 A 到 H 共 8 欄
      companyName.trim(),
      productName.trim(),
      parseFloat(sellingPricePerBox),
      String(casePricePerBox).trim() === '' ? null : parseFloat(casePricePerBox),
      String(totalCasePrice).trim() === '' ? null : parseFloat(totalCasePrice),
      parseInt(boxesPerCase),
      canShipAboveCase === true, 
      newImageFileId
    ]]);
    SpreadsheetApp.flush();
    
    if (oldImageToDelete) {
        deleteDriveFile_(oldImageToDelete); 
    }

    clearProductsCache_();
    return { success: true };
  } catch (e) {
    console.error("更新商品失敗 (列號: " + productData.rowNumber + ", 商品: " + productData.productName + "): " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 刪除商品資料 (包含圖片)
 */
function deleteProductData(productInfo) {
  try {
    const { rowNumber, imageId } = productInfo; 
    if (!rowNumber || typeof rowNumber !== 'number' || rowNumber < 2) {
      throw new Error("提供的列號無效。");
    }

    const sheet = getProductSheet_();
    if (rowNumber > sheet.getMaxRows() || rowNumber > sheet.getLastRow()) {
      throw new Error("無效的列號，超出試算表資料範圍。");
    }
    
    sheet.deleteRow(rowNumber); 
    SpreadsheetApp.flush();

    if (imageId) { 
      deleteDriveFile_(imageId);
    }
    
    clearProductsCache_();
    return { success: true };
  } catch (e) {
    console.error("刪除商品失敗 (列號: " + productInfo.rowNumber + "): " + e.toString());
    return { success: false, error: e.message };
  }
}
