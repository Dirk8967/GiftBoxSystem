// --- 1. 請在這裡完成您的設定 ---

// 您要備份的來源資料夾 ID
const SOURCE_FOLDER_ID = '1StxlP5447n0iAcPCnVIj40PESelm9jVS';

// 您要存放備份檔案的目的地資料夾 ID
const BACKUP_FOLDER_ID = '16yDB3y7Kr1yOZAMwOf_N5bs2Xm94zDQE';

// 【特殊檔案清單】將所有含有綁定腳本、需要被特殊處理的試算表檔案 ID 放入這個清單中
// 請用單引號 ' ' 包住每個 ID，並用逗號 , 隔開
const SPECIAL_SHEET_IDS_TO_COPY_DATA_ONLY = [
  SPREADSHEET_ID // 使用者授權清單
  // 'ID_A_在此貼上第一個特殊試算表的ID',
  // 'ID_B_在此貼上第二個特殊試算表的ID',
  // 'ID_C_在此貼上第三個特殊試算表的ID'  // 如果還有更多，請繼續往下加
];

// --- 設定結束 ---


/**
 * 每日備份主函式 (已支援多個特殊檔案處理)
 */
function createDailyBackup() {
  try {
    // --- 步驟 1: 執行舊備份刪除 ---
    Logger.log("--- 開始執行舊備份清理任務 ---");
    deleteOldBackups();
    Logger.log("--- 舊備份清理任務完成 ---");
    
    // --- 步驟 2: 建立今日新備份 ---
    Logger.log("--- 開始執行今日備份任務 ---");
    const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const backupRootFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);

    const today = new Date();
    const timezone = Session.getScriptTimeZone();
    const folderName = "備份 " + Utilities.formatDate(today, timezone, "yyyy-MM-dd");
    const todayBackupFolder = backupRootFolder.createFolder(folderName);

    Logger.log(`成功建立今日備份資料夾: ${folderName}`);

    const files = sourceFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const fileId = file.getId();

      // 【邏輯變更處】檢查當前檔案的 ID 是否包含在我們的特殊清單中
      if (SPECIAL_SHEET_IDS_TO_COPY_DATA_ONLY.includes(fileId)) {
        // --- 執行特殊備份：只複製內容，不複製腳本 ---
        copySheetDataOnly(file, todayBackupFolder);
        Logger.log(`已特殊備份 (無腳本): ${fileName}`);
      } else {
        // --- 執行一般備份：完整複製 ---
        file.makeCopy(fileName, todayBackupFolder);
        Logger.log(`已一般備份: ${fileName}`);
      }
    }
    Logger.log("--- 今日備份任務完成！ ---");

  } catch (e) {
    Logger.log(`過程中發生錯誤: ${e.toString()}`);
    // MailApp.sendEmail('your-email@example.com', 'GAS 備份失敗通知', e.toString());
  }
}

/**
 * 【功能不變】刪除超過一年的舊備份資料夾
 */
function deleteOldBackups() {
  const backupRootFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  Logger.log(`將刪除建立於 ${Utilities.formatDate(oneYearAgo, Session.getScriptTimeZone(), "yyyy-MM-dd")} 之前的備份。`);
  const folders = backupRootFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const creationDate = folder.getDateCreated();
    if (creationDate < oneYearAgo) {
      try {
        Logger.log(`準備刪除舊備份資料夾: ${folder.getName()} (建立於 ${creationDate})`);
        folder.setTrashed(true);
        Logger.log(` -> 已成功移至垃圾桶。`);
      } catch (err) {
        Logger.log(` -> 刪除失敗: ${err.toString()}`);
      }
    }
  }
}

/**
 * 【功能不變】特殊函式：只複製一個試算表的「所有分頁內容」，但不包含綁定的腳本
 */
function copySheetDataOnly(sourceFile, destinationFolder) {
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceFile.getId());
  const sourceSheets = sourceSpreadsheet.getSheets();
  const newSpreadsheetName = sourceFile.getName();
  const newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);
  const newSpreadsheetFile = DriveApp.getFileById(newSpreadsheet.getId());
  newSpreadsheetFile.moveTo(destinationFolder);

  sourceSheets.forEach((sheet, index) => {
    const copiedSheet = sheet.copyTo(newSpreadsheet);
    copiedSheet.setName(sheet.getName());
    const tabColor = sheet.getTabColor();
    if (tabColor) {
      copiedSheet.setTabColor(tabColor);
    }
  });

  const defaultSheet = newSpreadsheet.getSheetByName('工作表1');
  if (defaultSheet && newSpreadsheet.getSheets().length > 1) {
    newSpreadsheet.deleteSheet(defaultSheet);
  }
}
