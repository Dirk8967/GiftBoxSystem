// ================================================================= //
//                       站點資訊管理 - 後端邏輯
// ================================================================= //

const SITE_INFO_SHEET_ID = "1kipkFbx_-ryYPSyrDkJAbDJWxK7Xoo5wEZbDiayqfDs";
const SITE_INFO_CACHE_KEY = 'allSiteInfoData_v2'; // 使用與之前不同的快取鍵或更新版本
const SITE_INFO_CACHE_DURATION_SECONDS = 300; // 快取 5 分鐘

/**
 * 輔助函式：獲取站點資訊工作表物件
 */
function getSiteInfoSheet_() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SITE_INFO_SHEET_ID);
    if (!spreadsheet) {
        Logger.log("getSiteInfoSheet_ 錯誤: 無法透過ID開啟試算表: " + SITE_INFO_SHEET_ID);
        throw new Error("伺服器錯誤：無法開啟站點資訊試算表。");
    }
    const sheet = spreadsheet.getSheets()[0]; 
    if (!sheet) {
        Logger.log("getSiteInfoSheet_ 錯誤: 站點資訊試算表中找不到工作表。ID: " + SITE_INFO_SHEET_ID);
        throw new Error("伺服器錯誤：站點資訊試算表中找不到工作表。");
    }
    return sheet;
  } catch (e) {
    console.error("getSiteInfoSheet_ 捕捉到嚴重錯誤: " + e.toString() + " Stack: " + e.stack);
    throw new Error("讀取站點資訊試算表設定時發生內部錯誤: " + e.message);
  }
}

/**
 * 輔助函式：清除站點資訊資料快取
 */
function clearSiteInfoCache_() {
  try {
    CacheService.getScriptCache().remove(SITE_INFO_CACHE_KEY);
    Logger.log("站點資訊資料快取已清除 (Key: " + SITE_INFO_CACHE_KEY + ")。");
  } catch (e) {
    console.error('清除站點資訊快取失敗: ' + e.toString());
  }
}

/**
 * 獲取所有站點資訊資料
 */
function getSiteInfoListData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SITE_INFO_CACHE_KEY);
  if (cached != null) {
    Logger.log('getSiteInfoListData: 從快取讀取站點資訊資料。');
    try { return JSON.parse(cached); } catch (e) { Logger.log('解析站點資訊快取失敗，重新讀取。 Error: ' + e.message);}
  }

  Logger.log('getSiteInfoListData: 從試算表讀取站點資訊資料。');
  try {
    const sheet = getSiteInfoSheet_();
    if (sheet.getLastRow() < 1) { Logger.log("站點資訊表空白"); cache.put(SITE_INFO_CACHE_KEY, JSON.stringify([]), SITE_INFO_CACHE_DURATION_SECONDS); return []; }
    if (sheet.getLastRow() === 1 && sheet.getRange("A1").getValue() === "") { 
        Logger.log("站點資訊表可能完全空白（A1為空）"); cache.put(SITE_INFO_CACHE_KEY, JSON.stringify([]), SITE_INFO_CACHE_DURATION_SECONDS); return []; 
    }
    if (sheet.getLastRow() === 1) { 
        Logger.log("站點資訊表只有標頭"); cache.put(SITE_INFO_CACHE_KEY, JSON.stringify([]), SITE_INFO_CACHE_DURATION_SECONDS); return []; 
    }

    const numDataRows = sheet.getLastRow() - 1;
    if (numDataRows <= 0) { cache.put(SITE_INFO_CACHE_KEY, JSON.stringify([]), SITE_INFO_CACHE_DURATION_SECONDS); return []; }
    
    const values = sheet.getRange(2, 1, numDataRows, 4).getValues(); // A到D欄 (站代號, 站名, 地址, 電話)
    
    const sites = values.map(function(row, index) {
      return {
        rowNumber: index + 2,
        siteCode: (row[0] !== undefined && row[0] !== null) ? String(row[0]) : '',  
        siteName: (row[1] !== undefined && row[1] !== null) ? String(row[1]) : '',  
        address: (row[2] !== undefined && row[2] !== null) ? String(row[2]) : '',   
        phone: (row[3] !== undefined && row[3] !== null) ? String(row[3]) : ''     
      };
    });
    
    cache.put(SITE_INFO_CACHE_KEY, JSON.stringify(sites), SITE_INFO_CACHE_DURATION_SECONDS);
    Logger.log("站點資訊資料已讀取並快取，共 " + sites.length + " 筆。");
    return sites;
  } catch (e) {
    console.error('讀取站點資訊失敗 (getSiteInfoListData): ' + e.toString() + ' Stack: ' + e.stack);
    throw new Error('伺服器讀取站點資訊時發生錯誤。');
  }
}

/**
 * 新增站點資訊資料，並強制設定站代號和電話格式以保留前導零
 */
function addSiteInfoData(siteData) {
  try {
    const { siteCode, siteName, address, phone } = siteData;
    if (!siteCode || !siteName) { 
      throw new Error("站代號和站名為必填欄位。");
    }
    
    const sheet = getSiteInfoSheet_(); 
    const siteCodeStr = String(siteCode).trim();
    const phoneStr = String(phone || '').trim(); 

    // 可選：檢查站代號是否已存在
    if (sheet.getLastRow() >= 2) {
        const siteCodesInSheet = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
        if (siteCodesInSheet.map(function(sc){ return String(sc).trim(); }).includes(siteCodeStr)) {
            throw new Error("站代號 '" + siteCodeStr + "' 已存在。");
        }
    }

    sheet.appendRow([
      "'" + siteCodeStr,       // 在站代號前加上單引號
      String(siteName).trim(),
      String(address || '').trim(),
      phoneStr ? ("'" + phoneStr) : ""   // 如果電話有值，則在電話前加上單引號
    ]);

    const lastRow = sheet.getLastRow();
    // 強制設定格式
    sheet.getRange(lastRow, 1).setNumberFormat("@"); // A欄: 站代號
    sheet.getRange(lastRow, 4).setNumberFormat("@"); // D欄: 電話
    
    SpreadsheetApp.flush();
    clearSiteInfoCache_();
    return { success: true };
  } catch (e) {
    console.error("新增站點資訊失敗 (站名: " + siteData.siteName + "): " + e.toString() + ' Stack: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * 更新站點資訊資料，並強制設定站代號和電話格式以保留前導零
 */
function updateSiteInfoData(siteData) {
  try {
    const { rowNumber, siteCode, siteName, address, phone } = siteData;
    if (!rowNumber || !siteCode || !siteName ) {
      throw new Error("缺少必要更新資訊 (列號、站代號、站名)。");
    }
    
    const sheet = getSiteInfoSheet_();
    const siteCodeStr = String(siteCode).trim();
    const phoneStr = String(phone || '').trim();

    // 可選：檢查更新後的站代號是否與其他列重複
    if (sheet.getLastRow() >=2) {
        const allSiteDataCodes = sheet.getRange(2, 1, sheet.getLastRow() -1, 1).getValues(); 
        for(let i=0; i < allSiteDataCodes.length; i++) {
            const currentRowInSheet = i + 2; 
            if (currentRowInSheet !== rowNumber && String(allSiteDataCodes[i][0]).trim() === siteCodeStr) {
                throw new Error("更新後的站代號 '" + siteCodeStr + "' 已存在於第 " + currentRowInSheet + " 列中。");
            }
        }
    }
    
    sheet.getRange(rowNumber, 1, 1, 4).setValues([[ 
      "'" + siteCodeStr,       // 在站代號前加上單引號
      String(siteName).trim(),
      String(address || '').trim(),
      phoneStr ? ("'" + phoneStr) : ""   // 如果電話有值，則在電話前加上單引號
    ]]);

    // 強制設定格式
    sheet.getRange(rowNumber, 1).setNumberFormat("@"); // A欄: 站代號
    sheet.getRange(rowNumber, 4).setNumberFormat("@"); // D欄: 電話

    SpreadsheetApp.flush();
    clearSiteInfoCache_();
    return { success: true };
  } catch (e) {
    console.error("更新站點資訊失敗 (列號: " + siteData.rowNumber + ", 站名: " + siteData.siteName + "): " + e.toString() + ' Stack: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * 刪除站點資訊資料
 */
function deleteSiteInfoData(siteInfo) {
  try {
    const { rowNumber } = siteInfo;
    if (!rowNumber || typeof rowNumber !== 'number' || rowNumber < 2) {
      throw new Error("提供的列號無效。");
    }

    const sheet = getSiteInfoSheet_();
    if (rowNumber > sheet.getMaxRows() || rowNumber > sheet.getLastRow()) {
      throw new Error("無效的列號，超出試算表資料範圍。");
    }
    
    sheet.deleteRow(rowNumber); 
    SpreadsheetApp.flush();
    clearSiteInfoCache_();
    return { success: true };
  } catch (e) {
    console.error("刪除站點資訊失敗 (列號: " + siteInfo.rowNumber + "): " + e.toString() + ' Stack: ' + e.stack);
    return { success: false, error: e.message };
  }
}
