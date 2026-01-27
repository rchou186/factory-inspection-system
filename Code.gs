/**
 * 工廠環境衛生點檢管理系統
 * Factory Environmental Hygiene Inspection System
 * 
 * 文件編號: SYS-PLAN-2026-001
 * 版本: v2.2
 * ISO 9001 表單編號: QP-7.5-001
 */

// ==================== 全域變數設定 ====================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Sheet 名稱常數
const SHEETS = {
  RECORDS: 'Inspection_Records',           // Sheet 1: 點檢記錄表
  ITEM_MASTER: 'Item_Master',              // Sheet 2: 項目主檔
  TEMP_HUMIDITY: 'Temperature_Humidity',   // Sheet 3: 溫濕度記錄
  CUSTOM_STATS: 'Custom_Period_Statistics',// Sheet 4: 自訂期間統計
  SYSTEM_CONFIG: 'System_Config',          // Sheet 5: 系統設定
  TEMP_CRITERIA: 'TempHumidity_Criteria',  // Sheet 6: 溫濕度標準設定
  LANGUAGE: 'Language_Settings'            // Sheet 7: 多語系設定
};

// 系統預設值
const DEFAULTS = {
  SHIFTS: ['早', '晚', '其它'],
  LANGUAGE: 'zh-TW',
  ISO_FORM_PREFIX: 'QP-7.5-'
};

// ==================== 初始化函數 ====================

/**
 * 建立選單 - 當試算表開啟時自動執行
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏭 點檢系統')
    .addItem('📝 開啟點檢表單', 'openInspectionForm')
    .addItem('📊 查看統計報表', 'openStatistics')
    .addSeparator()
    .addItem('⚙️ 系統設定', 'openSettings')
    .addItem('🔄 初始化系統', 'initializeSystem')
    .addItem('📤 匯出資料', 'exportData')
    .addToUi();
}

/**
 * 初始化整個系統 - 建立所有必要的工作表和資料
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '初始化系統',
    '這將建立所有必要的工作表和資料。\n已存在的工作表將被保留。\n\n確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // 建立各個工作表
    createSheet_InspectionRecords(ss);
    createSheet_ItemMaster(ss);
    createSheet_TempHumidity(ss);
    createSheet_CustomStats(ss);
    createSheet_SystemConfig(ss);
    createSheet_TempCriteria(ss);
    createSheet_Language(ss);
    
    // 初始化資料
    initializeItemMaster();
    initializeTempCriteria();
    initializeLanguageSettings();
    initializeSystemConfig();
    
    // 設定觸發器
    setupTriggers();
    
    ui.alert('✅ 系統初始化完成！', '所有工作表和資料已建立完成。', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 錯誤', '初始化失敗：' + error.message, ui.ButtonSet.OK);
    Logger.log('初始化錯誤：' + error.stack);
  }
}

/**
 * 取得或建立工作表
 */
function getOrCreateSheet(ss, sheetName, headers = []) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  }
  
  return sheet;
}

// ==================== Web App 入口 ====================

/**
 * Web App 主入口 - doGet
 */
function doGet(e) {
  const page = e.parameter.page || 'inspection';
  const lang = e.parameter.lang || 'zh-TW';
  
  let template;
  
  switch(page) {
    case 'inspection':
      template = HtmlService.createTemplateFromFile('InspectionForm');
      break;
    case 'statistics':
      template = HtmlService.createTemplateFromFile('StatisticsView');
      break;
    default:
      template = HtmlService.createTemplateFromFile('InspectionForm');
  }
  
  template.lang = lang;
  
  return template.evaluate()
    .setTitle('工廠環境衛生點檢系統 - Factory Environmental Hygiene Inspection System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 開啟點檢表單
 */
function openInspectionForm() {
  const html = HtmlService.createTemplateFromFile('InspectionForm')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '工廠環境衛生點檢系統');
}

/**
 * 開啟統計報表
 */
function openStatistics() {
  const html = HtmlService.createTemplateFromFile('StatisticsView')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '統計分析報表');
}

/**
 * 開啟系統設定
 */
function openSettings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.SYSTEM_CONFIG);
  
  if (!configSheet) {
    ui.alert('錯誤', '找不到系統設定工作表，請先執行「初始化系統」', ui.ButtonSet.OK);
    return;
  }
  
  // 直接開啟系統設定工作表
  ss.setActiveSheet(configSheet);
  ui.alert(
    '系統設定',
    '您現在可以在「System_Config」工作表中修改系統設定。\n\n常用設定：\n• 預設點檢人員（逗號分隔）\n• 點檢時段\n• ISO表單編號',
    ui.ButtonSet.OK
  );
}

/**
 * 匯出資料
 */
function exportData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.alert(
    '匯出資料',
    '請選擇要匯出的資料類型：\n\n' +
    '「是」- 匯出點檢記錄\n' +
    '「否」- 匯出溫濕度記錄\n' +
    '「取消」- 取消操作',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  try {
    if (response === ui.Button.YES) {
      // 匯出點檢記錄
      const recordSheet = ss.getSheetByName(SHEETS.RECORDS);
      if (!recordSheet) {
        ui.alert('錯誤', '找不到點檢記錄表', ui.ButtonSet.OK);
        return;
      }
      
      const data = recordSheet.getDataRange().getValues();
      if (data.length <= 1) {
        ui.alert('提示', '目前沒有點檢記錄可匯出', ui.ButtonSet.OK);
        return;
      }
      
      // 建立新試算表
      const newSs = SpreadsheetApp.create('點檢記錄匯出_' + Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss'));
      const newSheet = newSs.getActiveSheet();
      newSheet.setName('Inspection_Records');
      
      // 複製資料
      newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      
      // 格式化
      newSheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
      newSheet.autoResizeColumns(1, data[0].length);
      
      ui.alert(
        '✅ 匯出成功',
        `點檢記錄已匯出 ${data.length - 1} 筆\n\n請至 Google Drive 查看：\n「點檢記錄匯出_${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss')}」`,
        ui.ButtonSet.OK
      );
      
      // 開啟新檔案
      const url = newSs.getUrl();
      const htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          window.open('${url}', '_blank');
          google.script.host.close();
        </script>
      `);
      ui.showModalDialog(htmlOutput, '正在開啟匯出檔案...');
      
    } else if (response === ui.Button.NO) {
      // 匯出溫濕度記錄
      const tempSheet = ss.getSheetByName(SHEETS.TEMP_HUMIDITY);
      if (!tempSheet) {
        ui.alert('錯誤', '找不到溫濕度記錄表', ui.ButtonSet.OK);
        return;
      }
      
      const data = tempSheet.getDataRange().getValues();
      if (data.length <= 1) {
        ui.alert('提示', '目前沒有溫濕度記錄可匯出', ui.ButtonSet.OK);
        return;
      }
      
      // 建立新試算表
      const newSs = SpreadsheetApp.create('溫濕度記錄匯出_' + Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss'));
      const newSheet = newSs.getActiveSheet();
      newSheet.setName('Temperature_Humidity');
      
      // 複製資料
      newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      
      // 格式化
      newSheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
      newSheet.autoResizeColumns(1, data[0].length);
      
      ui.alert(
        '✅ 匯出成功',
        `溫濕度記錄已匯出 ${data.length - 1} 筆\n\n請至 Google Drive 查看：\n「溫濕度記錄匯出_${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss')}」`,
        ui.ButtonSet.OK
      );
      
      // 開啟新檔案
      const url = newSs.getUrl();
      const htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          window.open('${url}', '_blank');
          google.script.host.close();
        </script>
      `);
      ui.showModalDialog(htmlOutput, '正在開啟匯出檔案...');
    }
    
  } catch (error) {
    ui.alert('❌ 錯誤', '匯出失敗：' + error.message, ui.ButtonSet.OK);
    Logger.log('匯出錯誤：' + error.stack);
  }
}

/**
 * Include 函數 - 用於載入 HTML 檔案中的 CSS/JS
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== 輔助函數 ====================

/**
 * 生成記錄ID
 */
function generateRecordId(prefix, date, shift) {
  const dateStr = Utilities.formatDate(date, 'GMT+8', 'yyyyMMdd');
  const shiftCode = shift === '早' ? 'AM' : (shift === '晚' ? 'PM' : 'OT');
  const timestamp = new Date().getTime().toString().slice(-3);
  return `${prefix}${dateStr}${shiftCode}${timestamp}`;
}

/**
 * 格式化日期
 */
function formatDate(date, format = 'yyyy-MM-dd') {
  return Utilities.formatDate(new Date(date), 'GMT+8', format);
}

/**
 * 取得當前時間戳記
 */
function getTimestamp() {
  return new Date();
}

/**
 * 記錄系統日誌
 */
function logSystem(message, level = 'INFO') {
  const timestamp = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss');
  Logger.log(`[${timestamp}] [${level}] ${message}`);
}

/**
 * 取得點檢人員清單
 */
function getInspectorList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SYSTEM_CONFIG);
    
    if (!sheet) return ['張三', '李四', '王五', '趙六', '陳七'];
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === '預設點檢人員' && data[i][1]) {
        const inspectors = data[i][1].toString().split(',').map(name => name.trim()).filter(name => name);
        if (inspectors.length > 0) {
          return inspectors;
        }
      }
    }
    
    return ['張三', '李四', '王五', '趙六', '陳七'];
    
  } catch (error) {
    logSystem('取得點檢人員清單失敗：' + error.message, 'ERROR');
    return ['張三', '李四', '王五', '趙六', '陳七'];
  }
}

