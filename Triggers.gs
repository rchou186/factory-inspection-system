/**
 * ============================================================================
 * 工廠環境衛生點檢管理系統 - 觸發器模組
 * Factory Environmental Hygiene Inspection System - Triggers Module
 * ============================================================================
 * 
 * @file        Triggers.gs
 * @version     v2.2.5
 * @date        2026-01-26
 * @author      System Developer
 * @description 管理系統定時觸發器，包含每日統計、自動報表產生等排程任務
 * 
 * @functions
 *   - setupTriggers()                 設定定時觸發器
 *   - dailyStatistics()               每日統計（23:00）
 *   - autoGenerateReports()           自動產生報表
 *   - cleanupOldData()                清理舊資料
 * 
 * @schedule
 *   - 每日 23:00：執行統計
 *   - 每週日：產生週報表
 *   - 每月底：產生月報表
 * 
 * @dependencies
 *   - Code.gs (SHEETS 常數)
 *   - Statistics.gs (統計功能)
 * 
 * @changelog
 *   v2.2.5 (2026-01-26) - ISO 22000 整合
 *   v2.2.0 (2026-01-22) - 觸發器管理功能
 * 
 * ============================================================================
 */

// ==================== 觸發器設定 ====================

/**
 * 設定所有觸發器
 */
function setupTriggers() {
  try {
    // 刪除現有觸發器
    deleteAllTriggers();
    
    // 每日 23:00 執行統計
    ScriptApp.newTrigger('dailyStatistics')
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();
    
    logSystem('每日統計觸發器已設定 (23:00)');
    
    // 每日 08:00 發送統計摘要
    ScriptApp.newTrigger('sendDailySummary')
      .timeBased()
      .atHour(8)
      .everyDays(1)
      .create();
    
    logSystem('每日摘要觸發器已設定 (08:00)');
    
    // 每 4 小時檢查待追蹤項目
    ScriptApp.newTrigger('checkPendingItems')
      .timeBased()
      .everyHours(4)
      .create();
    
    logSystem('待追蹤檢查觸發器已設定 (每4小時)');
    
    return {
      success: true,
      message: '觸發器設定完成'
    };
    
  } catch (error) {
    logSystem('設定觸發器失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 刪除所有觸發器
 */
function deleteAllTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    
    logSystem(`已刪除 ${triggers.length} 個觸發器`);
    
  } catch (error) {
    logSystem('刪除觸發器失敗：' + error.message, 'ERROR');
  }
}

// ==================== 定時執行功能 ====================

/**
 * 每日統計 - 在每天 23:00 執行
 */
function dailyStatistics() {
  try {
    logSystem('開始執行每日統計...', 'INFO');
    
    const today = new Date();
    const dateStr = formatDate(today);
    
    // 產生當日報表
    generateCustomPeriodReport(dateStr, dateStr, `${dateStr} 當日報表`);
    
    // 檢查當週是否為週日，若是則產生週報表
    if (today.getDay() === 0) {
      const weekStart = new Date(today);
      weekStart.setDate(weekStart.getDate() - 6);
      
      generateCustomPeriodReport(
        formatDate(weekStart),
        dateStr,
        `${formatDate(weekStart)} ~ ${dateStr} 週報表`
      );
      
      logSystem('週報表已產生', 'INFO');
    }
    
    // 檢查是否為月底，若是則產生月報表
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    if (tomorrow.getMonth() !== today.getMonth()) {
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      
      generateCustomPeriodReport(
        formatDate(monthStart),
        dateStr,
        `${today.getFullYear()}年${today.getMonth() + 1}月 月報表`
      );
      
      logSystem('月報表已產生', 'INFO');
    }
    
    logSystem('每日統計執行完成', 'INFO');
    
  } catch (error) {
    logSystem('每日統計執行失敗：' + error.message, 'ERROR');
  }
}

/**
 * 自動產生常用期間報表
 */
function autoGenerateReports() {
  try {
    const today = new Date();
    
    // 本週報表
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() - today.getDay());
    generateCustomPeriodReport(
      formatDate(weekStart),
      formatDate(today),
      '本週報表'
    );
    
    // 本月報表
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    generateCustomPeriodReport(
      formatDate(monthStart),
      formatDate(today),
      '本月報表'
    );
    
    // 上週報表
    const lastWeekEnd = new Date(weekStart);
    lastWeekEnd.setDate(lastWeekEnd.getDate() - 1);
    const lastWeekStart = new Date(lastWeekEnd);
    lastWeekStart.setDate(lastWeekEnd.getDate() - 6);
    generateCustomPeriodReport(
      formatDate(lastWeekStart),
      formatDate(lastWeekEnd),
      '上週報表'
    );
    
    // 上月報表
    const lastMonthEnd = new Date(monthStart);
    lastMonthEnd.setDate(0);
    const lastMonthStart = new Date(lastMonthEnd.getFullYear(), lastMonthEnd.getMonth(), 1);
    generateCustomPeriodReport(
      formatDate(lastMonthStart),
      formatDate(lastMonthEnd),
      '上月報表'
    );
    
    logSystem('常用期間報表已自動產生', 'INFO');
    
  } catch (error) {
    logSystem('自動產生報表失敗：' + error.message, 'ERROR');
  }
}

// ==================== 手動觸發功能 ====================

/**
 * 手動執行每日統計
 */
function manualDailyStatistics() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '執行每日統計',
    '確定要執行每日統計嗎？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    dailyStatistics();
    ui.alert('✅ 完成', '每日統計已執行完成。', ui.ButtonSet.OK);
  }
}

/**
 * 手動執行報表產生
 */
function manualGenerateReports() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '產生常用報表',
    '確定要產生本週、本月、上週、上月的報表嗎？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    autoGenerateReports();
    ui.alert('✅ 完成', '常用報表已產生完成。', ui.ButtonSet.OK);
  }
}

/**
 * 查看觸發器狀態
 */
function viewTriggerStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    ui.alert('📋 觸發器狀態', '目前沒有設定任何觸發器。\n\n請執行「初始化系統」來設定觸發器。', ui.ButtonSet.OK);
    return;
  }
  
  let message = `目前已設定 ${triggers.length} 個觸發器：\n\n`;
  
  triggers.forEach((trigger, index) => {
    const handlerFunction = trigger.getHandlerFunction();
    const eventType = trigger.getEventType();
    
    message += `${index + 1}. ${handlerFunction}\n`;
    message += `   類型：${eventType}\n\n`;
  });
  
  ui.alert('📋 觸發器狀態', message, ui.ButtonSet.OK);
}

// ==================== 資料清理 ====================

/**
 * 清理過期資料（保留3年）
 */
function cleanupOldData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = [SHEETS.RECORDS, SHEETS.TEMP_HUMIDITY];
    
    const threeYearsAgo = new Date();
    threeYearsAgo.setFullYear(threeYearsAgo.getFullYear() - 3);
    
    let totalDeleted = 0;
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const rowsToDelete = [];
      
      // 找出要刪除的列（從下往上，以便刪除時不影響索引）
      for (let i = data.length - 1; i > 0; i--) {
        const recordDate = new Date(data[i][1]); // 假設第2欄是日期
        
        if (recordDate < threeYearsAgo) {
          rowsToDelete.push(i + 1); // +1 因為陣列索引從0開始，但工作表列從1開始
        }
      }
      
      // 刪除過期資料
      rowsToDelete.forEach(rowNum => {
        sheet.deleteRow(rowNum);
        totalDeleted++;
      });
      
      if (rowsToDelete.length > 0) {
        logSystem(`${sheetName}: 已刪除 ${rowsToDelete.length} 筆過期資料`);
      }
    });
    
    if (totalDeleted > 0) {
      logSystem(`資料清理完成：共刪除 ${totalDeleted} 筆過期資料`, 'INFO');
    } else {
      logSystem('無過期資料需要清理', 'INFO');
    }
    
  } catch (error) {
    logSystem('清理過期資料失敗：' + error.message, 'ERROR');
  }
}

/**
 * 手動清理資料
 */
function manualCleanupData() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '清理過期資料',
    '這將刪除3年前的點檢記錄和溫濕度記錄。\n\n確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    cleanupOldData();
    ui.alert('✅ 完成', '資料清理已執行完成。', ui.ButtonSet.OK);
  }
}

// ==================== 備份功能 ====================

/**
 * 備份資料到新的試算表
 */
function backupData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss');
    const backupName = `工廠點檢系統備份_${timestamp}`;
    
    // 複製試算表
    const backup = ss.copy(backupName);
    
    logSystem(`資料備份完成：${backupName}`, 'INFO');
    
    return {
      success: true,
      backupName: backupName,
      backupId: backup.getId(),
      backupUrl: backup.getUrl()
    };
    
  } catch (error) {
    logSystem('資料備份失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 手動備份
 */
function manualBackup() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '備份資料',
    '確定要建立系統備份嗎？\n\n這將複製整個試算表。',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const backupResult = backupData();
    
    if (backupResult.success) {
      ui.alert(
        '✅ 備份完成',
        `備份名稱：${backupResult.backupName}\n\n備份已儲存在您的 Google Drive 中。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ 備份失敗',
        `錯誤訊息：${backupResult.error}`,
        ui.ButtonSet.OK
      );
    }
  }
}
