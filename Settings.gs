// ================================================================= //
//                       設定管理 - 後端邏輯
// ================================================================= //

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

/**
 * 【已更新】儲存所有檔期設定
 * @param {Object} settings - 包含 { activityName, activityStartTime, activityEndTime, looseOrderStartTime, looseOrderEndTime } 的物件
 * @returns {Object} { success: boolean, error?: string }
 */
function saveSettings(settings) {
  try {
    // 全站檔期的時間是必填的
    if (!settings || !settings.activityStartTime || !settings.activityEndTime) {
      throw new Error("全站檔期的開始和結束時間為必填欄位。");
    }
    SCRIPT_PROPERTIES.setProperties({
      'activityName': settings.activityName || '',
      'activityStartTime': settings.activityStartTime,
      'activityEndTime': settings.activityEndTime,
      // 零星訂購的時間可以為空，表示不限制
      'looseOrderStartTime': settings.looseOrderStartTime || '',
      'looseOrderEndTime': settings.looseOrderEndTime || ''
    });
    return { success: true };
  } catch (e) {
    console.error("saveSettings 錯誤: " + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * 【已更新】獲取目前儲存的所有設定
 * @returns {Object} 包含所有設定的物件
 */
function getSettings() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    return {
      activityName: props.activityName || '',
      activityStartTime: props.activityStartTime || '',
      activityEndTime: props.activityEndTime || '',
      looseOrderStartTime: props.looseOrderStartTime || '',
      looseOrderEndTime: props.looseOrderEndTime || ''
    };
  } catch (e) {
    console.error("getSettings 錯誤: " + e.toString());
    return { error: e.message };
  }
}

/**
 * 供使用者端（index.html）獲取公開的活動資訊
 * @returns {Object} { name }
 */
function getPublicInfo() {
    try {
        const name = SCRIPT_PROPERTIES.getProperty('activityName');
        return { name: name || '預設活動' };
    } catch(e) {
        return { name: '活動' };
    }
}


/**
 * 【新增】檢查目前時間是否在「零星訂購」的開放區間內
 * @returns {boolean} true 表示在區間內或未設定，false 表示不在區間內
 */
function isLooseOrderPeriodActive() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    const startTimeStr = props.looseOrderStartTime;
    const endTimeStr = props.looseOrderEndTime;

    // 如果沒有設定零星訂購的開始或結束時間，預設為永遠開放
    if (!startTimeStr || !endTimeStr) {
      Logger.log("零星訂購時間未設定，預設為開啟。");
      return true;
    }

    const start = new Date(startTimeStr);
    const end = new Date(endTimeStr);
    const now = new Date();

    Logger.log("零星訂購檔期檢查: 開始=" + start + ", 結束=" + end + ", 現在=" + now);
    return now >= start && now <= end;

  } catch(e) {
    console.error("isLooseOrderPeriodActive 錯誤: " + e.toString());
    return false; // 如果設定有誤，為安全起見，預設為關閉
  }
}


/**
 * 檢查目前時間是否在「全站」活動檔期內 (此函式名稱和邏輯保持不變，供 doGet 使用)
 * @returns {boolean} true 表示在檔期內，false 表示不在檔期內
 */
function checkActivityPeriod_() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    const startTimeStr = props.activityStartTime;
    const endTimeStr = props.activityEndTime;

    if (!startTimeStr || !endTimeStr) {
      Logger.log("全站檔期未設定，預設為開啟。");
      return true;
    }

    const start = new Date(startTimeStr);
    const end = new Date(endTimeStr);
    const now = new Date();

    Logger.log("全站檔期檢查: 開始=" + start + ", 結束=" + end + ", 現在=" + now);
    return now >= start && now <= end;

  } catch(e) {
    console.error("checkActivityPeriod_ 錯誤: " + e.toString());
    return false;
  }
}
