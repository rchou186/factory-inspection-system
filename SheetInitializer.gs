/**
 * 工作表初始化模組
 * Sheet Initializer Module
 */

// ==================== Sheet 1: 點檢記錄表 ====================

function createSheet_InspectionRecords(ss) {
  const headers = [
    '記錄ID', '日期', '時段', '區域', '項目編號', '項目名稱',
    '狀態', '前次狀態', '備註', '點檢人員', '時間戳記'
  ];
  
  const sheet = getOrCreateSheet(ss, SHEETS.RECORDS, headers);
  
  // 設定欄寬
  sheet.setColumnWidth(1, 150);  // 記錄ID
  sheet.setColumnWidth(2, 100);  // 日期
  sheet.setColumnWidth(3, 60);   // 時段
  sheet.setColumnWidth(4, 120);  // 區域
  sheet.setColumnWidth(5, 80);   // 項目編號
  sheet.setColumnWidth(6, 300);  // 項目名稱
  sheet.setColumnWidth(7, 80);   // 狀態
  sheet.setColumnWidth(8, 80);   // 前次狀態
  sheet.setColumnWidth(9, 200);  // 備註
  sheet.setColumnWidth(10, 100); // 點檢人員
  sheet.setColumnWidth(11, 150); // 時間戳記
  
  return sheet;
}

// ==================== Sheet 2: 項目主檔 ====================

function createSheet_ItemMaster(ss) {
  const headers = [
    '項目ID', '區域', '項目編號', '項目名稱(中)', '項目名稱(英)', '是否啟用'
  ];
  
  const sheet = getOrCreateSheet(ss, SHEETS.ITEM_MASTER, headers);
  
  sheet.setColumnWidth(1, 120);  // 項目ID
  sheet.setColumnWidth(2, 120);  // 區域
  sheet.setColumnWidth(3, 80);   // 項目編號
  sheet.setColumnWidth(4, 350);  // 項目名稱(中)
  sheet.setColumnWidth(5, 400);  // 項目名稱(英)
  sheet.setColumnWidth(6, 80);   // 是否啟用
  
  return sheet;
}

// ==================== Sheet 3: 溫濕度記錄 ====================

function createSheet_TempHumidity(ss) {
  const headers = [
    '記錄ID', '日期', '時段', '區域', '測點名稱', '數值', '單位',
    '判定結果', '超標說明', '備註', '點檢人員', '告警狀態'
  ];
  
  const sheet = getOrCreateSheet(ss, SHEETS.TEMP_HUMIDITY, headers);
  
  sheet.setColumnWidth(1, 150);  // 記錄ID
  sheet.setColumnWidth(2, 100);  // 日期
  sheet.setColumnWidth(3, 60);   // 時段
  sheet.setColumnWidth(4, 120);  // 區域
  sheet.setColumnWidth(5, 150);  // 測點名稱
  sheet.setColumnWidth(6, 80);   // 數值
  sheet.setColumnWidth(7, 60);   // 單位
  sheet.setColumnWidth(8, 100);  // 判定結果
  sheet.setColumnWidth(9, 150);  // 超標說明
  sheet.setColumnWidth(10, 200); // 備註
  sheet.setColumnWidth(11, 100); // 點檢人員
  sheet.setColumnWidth(12, 100); // 告警狀態
  
  return sheet;
}

// ==================== Sheet 4: 自訂期間統計 ====================

function createSheet_CustomStats(ss) {
  const headers = [
    '統計ID', '起始日期', '結束日期', '期間描述', '區域', '項目名稱',
    'OK次數', 'NOK次數', '待追蹤次數', '總次數', 'NOK比率', '排名'
  ];
  
  const sheet = getOrCreateSheet(ss, SHEETS.CUSTOM_STATS, headers);
  
  sheet.setColumnWidth(1, 150);  // 統計ID
  sheet.setColumnWidth(2, 100);  // 起始日期
  sheet.setColumnWidth(3, 100);  // 結束日期
  sheet.setColumnWidth(4, 150);  // 期間描述
  sheet.setColumnWidth(5, 120);  // 區域
  sheet.setColumnWidth(6, 300);  // 項目名稱
  sheet.setColumnWidth(7, 80);   // OK次數
  sheet.setColumnWidth(8, 80);   // NOK次數
  sheet.setColumnWidth(9, 100);  // 待追蹤次數
  sheet.setColumnWidth(10, 80);  // 總次數
  sheet.setColumnWidth(11, 80);  // NOK比率
  sheet.setColumnWidth(12, 60);  // 排名
  
  return sheet;
}

// ==================== Sheet 5: 系統設定 ====================

function createSheet_SystemConfig(ss) {
  const headers = ['設定項目', '設定值'];
  const sheet = getOrCreateSheet(ss, SHEETS.SYSTEM_CONFIG, headers);
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 400);
  
  return sheet;
}

// ==================== Sheet 6: 溫濕度標準設定 ====================

function createSheet_TempCriteria(ss) {
  const headers = [
    '測點名稱', '所屬區域', '測量類型', '單位', '下限值', '上限值',
    '是否啟用告警', '告警方式', '告警收件人', '備註'
  ];
  
  const sheet = getOrCreateSheet(ss, SHEETS.TEMP_CRITERIA, headers);
  
  sheet.setColumnWidth(1, 150);  // 測點名稱
  sheet.setColumnWidth(2, 120);  // 所屬區域
  sheet.setColumnWidth(3, 100);  // 測量類型
  sheet.setColumnWidth(4, 60);   // 單位
  sheet.setColumnWidth(5, 80);   // 下限值
  sheet.setColumnWidth(6, 80);   // 上限值
  sheet.setColumnWidth(7, 120);  // 是否啟用告警
  sheet.setColumnWidth(8, 100);  // 告警方式
  sheet.setColumnWidth(9, 250);  // 告警收件人
  sheet.setColumnWidth(10, 200); // 備註
  
  return sheet;
}

// ==================== Sheet 7: 多語系設定 ====================

function createSheet_Language(ss) {
  const headers = ['語系代碼', '語系名稱', '欄位鍵值', '翻譯內容'];
  const sheet = getOrCreateSheet(ss, SHEETS.LANGUAGE, headers);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 400);
  
  return sheet;
}

// ==================== 初始化項目主檔資料 ====================

function initializeItemMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ITEM_MASTER);
  
  if (!sheet) {
    throw new Error('項目主檔工作表不存在');
  }
  
  // 檢查是否已有資料
  if (sheet.getLastRow() > 1) {
    Logger.log('項目主檔已有資料，跳過初始化');
    return;
  }
  
  const items = getAllInspectionItems();
  const data = [];
  
  items.forEach((area) => {
    area.items.forEach((item) => {
      data.push([
        `${area.code}_${item.no}`,  // 項目ID
        area.name_zh,                // 區域
        item.no,                     // 項目編號
        item.name_zh,                // 項目名稱(中)
        item.name_en,                // 項目名稱(英)
        true                         // 是否啟用
      ]);
    });
  });
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
  
  Logger.log(`項目主檔初始化完成：${data.length} 個項目`);
}

// ==================== 初始化溫濕度標準 ====================

function initializeTempCriteria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TEMP_CRITERIA);
  
  if (!sheet) {
    throw new Error('溫濕度標準設定工作表不存在');
  }
  
  // 檢查是否已有資料
  if (sheet.getLastRow() > 1) {
    Logger.log('溫濕度標準已有資料，跳過初始化');
    return;
  }
  
  const criteria = [
    ['秤料室濕度', 'B1區', '濕度', '%RH', 30, 60, true, '郵件', '', '依食品GMP規範'],
    ['物料室濕度', 'B1區', '濕度', '%RH', 30, 60, true, '郵件', '', '依食品GMP規範'],
    ['半成品暫存區溫度', '蒸煮區', '溫度', '℃', 0, 7, true, '郵件', '', '冷藏溫度控制'],
    ['半成品冷凍區溫度', '蒸煮區', '溫度', '℃', -18, -10, true, '郵件', '', '冷凍溫度控制'],
    ['原料庫溫度', '其他', '溫度', '℃', -20, -15, true, '郵件', '', '冷凍庫溫度控制'],
    ['成品庫溫度', '其他', '溫度', '℃', -20, -15, true, '郵件', '', '冷凍庫溫度控制']
  ];
  
  sheet.getRange(2, 1, criteria.length, 10).setValues(criteria);
  
  Logger.log(`溫濕度標準初始化完成：${criteria.length} 個測點`);
}

// ==================== 初始化系統設定 ====================

function initializeSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SYSTEM_CONFIG);
  
  if (!sheet) {
    throw new Error('系統設定工作表不存在');
  }
  
  // 檢查是否已有資料
  if (sheet.getLastRow() > 1) {
    Logger.log('系統設定已有資料，跳過初始化');
    return;
  }
  
  const config = [
    ['點檢時段', '早,晚,其它'],
    ['預設點檢人員', '張三,李四,王五,趙六,陳七'],
    ['自動統計時間', '每日23:00'],
    ['數據保留期限', '3年'],
    ['預設語系', '繁體中文'],
    ['ISO表單編號', 'QP-7.5-001'],
    ['系統版本', 'v2.2.4'],
    ['最後更新日期', Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd')]
  ];
  
  sheet.getRange(2, 1, config.length, 2).setValues(config);
  
  Logger.log('系統設定初始化完成');
}
