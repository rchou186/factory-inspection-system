/**
 * ============================================================================
 * 工廠環境衛生點檢管理系統 - 多語系模組
 * Factory Environmental Hygiene Inspection System - i18n Module
 * ============================================================================
 * 
 * @file        I18n.gs
 * @version     v2.2.5
 * @date        2026-01-26
 * @author      System Developer
 * @description 提供中英雙語系支援，管理所有 UI 介面的語系切換功能
 * 
 * @languages
 *   - zh-TW：繁體中文
 *   - en-US：English
 * 
 * @functions
 *   - loadLanguage()                  載入語系資料
 *   - translateUI()                   翻譯 UI 介面
 *   - getCurrentLanguage()            取得當前語系
 *   - setLanguage()                   設定使用語系
 * 
 * @dependencies
 *   - Code.gs (SHEETS 常數)
 * 
 * @changelog
 *   v2.2.5 (2026-01-26) - ISO 22000 整合
 *   v2.2.0 (2026-01-22) - 多語系切換功能
 * 
 * ============================================================================
 */

// ==================== 初始化語系資料 ====================

/**
 * 初始化多語系設定
 */
function initializeLanguageSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LANGUAGE);
  
  if (!sheet) {
    throw new Error('多語系設定工作表不存在');
  }
  
  // 檢查是否已有資料
  if (sheet.getLastRow() > 1) {
    Logger.log('多語系設定已有資料，跳過初始化');
    return;
  }
  
  const languageData = getDefaultLanguageData();
  
  if (languageData.length > 0) {
    sheet.getRange(2, 1, languageData.length, 4).setValues(languageData);
  }
  
  Logger.log(`多語系設定初始化完成：${languageData.length} 筆翻譯`);
}

/**
 * 取得預設語系資料
 */
function getDefaultLanguageData() {
  return [
    // 系統標題
    ['zh-TW', '繁體中文', 'system_title', '工廠環境衛生點檢系統'],
    ['en-US', 'English', 'system_title', 'Factory Environmental Hygiene Inspection System'],
    
    // 表單編號
    ['zh-TW', '繁體中文', 'form_number', '表單編號'],
    ['en-US', 'English', 'form_number', 'Form Number'],
    
    // 日期時段
    ['zh-TW', '繁體中文', 'date', '日期'],
    ['en-US', 'English', 'date', 'Date'],
    ['zh-TW', '繁體中文', 'shift', '時段'],
    ['en-US', 'English', 'shift', 'Shift'],
    ['zh-TW', '繁體中文', 'shift_early', '早'],
    ['en-US', 'English', 'shift_early', 'Morning'],
    ['zh-TW', '繁體中文', 'shift_late', '晚'],
    ['en-US', 'English', 'shift_late', 'Evening'],
    ['zh-TW', '繁體中文', 'shift_other', '其它'],
    ['en-US', 'English', 'shift_other', 'Other'],
    
    // 人員
    ['zh-TW', '繁體中文', 'inspector', '點檢人員'],
    ['en-US', 'English', 'inspector', 'Inspector'],
    
    // 區域選擇
    ['zh-TW', '繁體中文', 'area_selection', '區域選擇'],
    ['en-US', 'English', 'area_selection', 'Area Selection'],
    
    // 狀態
    ['zh-TW', '繁體中文', 'status', '狀態'],
    ['en-US', 'English', 'status', 'Status'],
    ['zh-TW', '繁體中文', 'status_ok', 'OK'],
    ['en-US', 'English', 'status_ok', 'OK'],
    ['zh-TW', '繁體中文', 'status_nok', 'NOK'],
    ['en-US', 'English', 'status_nok', 'NOK'],
    ['zh-TW', '繁體中文', 'status_pending', '待追蹤'],
    ['en-US', 'English', 'status_pending', 'Pending'],
    
    // 備註
    ['zh-TW', '繁體中文', 'note', '備註'],
    ['en-US', 'English', 'note', 'Note'],
    
    // 按鈕
    ['zh-TW', '繁體中文', 'submit', '提交'],
    ['en-US', 'English', 'submit', 'Submit'],
    ['zh-TW', '繁體中文', 'clear', '清除'],
    ['en-US', 'English', 'clear', 'Clear'],
    ['zh-TW', '繁體中文', 'back', '返回'],
    ['en-US', 'English', 'back', 'Back'],
    ['zh-TW', '繁體中文', 'export', '匯出'],
    ['en-US', 'English', 'export', 'Export'],
    ['zh-TW', '繁體中文', 'query', '查詢'],
    ['en-US', 'English', 'query', 'Query'],
    
    // 統計報表
    ['zh-TW', '繁體中文', 'statistics', '統計分析'],
    ['en-US', 'English', 'statistics', 'Statistics'],
    ['zh-TW', '繁體中文', 'custom_period', '自訂統計期間'],
    ['en-US', 'English', 'custom_period', 'Custom Period'],
    ['zh-TW', '繁體中文', 'start_date', '起始日期'],
    ['en-US', 'English', 'start_date', 'Start Date'],
    ['zh-TW', '繁體中文', 'end_date', '結束日期'],
    ['en-US', 'English', 'end_date', 'End Date'],
    
    // 快速選擇
    ['zh-TW', '繁體中文', 'this_week', '本週'],
    ['en-US', 'English', 'this_week', 'This Week'],
    ['zh-TW', '繁體中文', 'this_month', '本月'],
    ['en-US', 'English', 'this_month', 'This Month'],
    ['zh-TW', '繁體中文', 'this_quarter', '本季'],
    ['en-US', 'English', 'this_quarter', 'This Quarter'],
    ['zh-TW', '繁體中文', 'this_year', '本年'],
    ['en-US', 'English', 'this_year', 'This Year'],
    ['zh-TW', '繁體中文', 'last_week', '上週'],
    ['en-US', 'English', 'last_week', 'Last Week'],
    ['zh-TW', '繁體中文', 'last_month', '上月'],
    ['en-US', 'English', 'last_month', 'Last Month'],
    
    // NOK排名
    ['zh-TW', '繁體中文', 'top_nok', 'NOK項目排名'],
    ['en-US', 'English', 'top_nok', 'Top NOK Items'],
    ['zh-TW', '繁體中文', 'rank', '排名'],
    ['en-US', 'English', 'rank', 'Rank'],
    ['zh-TW', '繁體中文', 'nok_count', 'NOK次數'],
    ['en-US', 'English', 'nok_count', 'NOK Count'],
    ['zh-TW', '繁體中文', 'nok_ratio', 'NOK比率'],
    ['en-US', 'English', 'nok_ratio', 'NOK Ratio'],
    
    // 圖表
    ['zh-TW', '繁體中文', 'trend_chart', '趨勢圖'],
    ['en-US', 'English', 'trend_chart', 'Trend Chart'],
    ['zh-TW', '繁體中文', 'pie_chart', '圓餅圖'],
    ['en-US', 'English', 'pie_chart', 'Pie Chart'],
    
    // 溫濕度
    ['zh-TW', '繁體中文', 'temp_humidity', '溫濕度監控'],
    ['en-US', 'English', 'temp_humidity', 'Temperature/Humidity Monitor'],
    ['zh-TW', '繁體中文', 'value', '數值'],
    ['en-US', 'English', 'value', 'Value'],
    ['zh-TW', '繁體中文', 'unit', '單位'],
    ['en-US', 'English', 'unit', 'Unit'],
    ['zh-TW', '繁體中文', 'normal', '正常'],
    ['en-US', 'English', 'normal', 'Normal'],
    ['zh-TW', '繁體中文', 'abnormal', '異常'],
    ['en-US', 'English', 'abnormal', 'Abnormal'],
    ['zh-TW', '繁體中文', 'standard_range', '標準範圍'],
    ['en-US', 'English', 'standard_range', 'Standard Range'],
    
    // 訊息
    ['zh-TW', '繁體中文', 'msg_submit_success', '提交成功'],
    ['en-US', 'English', 'msg_submit_success', 'Submit Successfully'],
    ['zh-TW', '繁體中文', 'msg_submit_failed', '提交失敗'],
    ['en-US', 'English', 'msg_submit_failed', 'Submit Failed'],
    ['zh-TW', '繁體中文', 'msg_required_fields', '請填寫所有必填欄位'],
    ['en-US', 'English', 'msg_required_fields', 'Please fill in all required fields'],
    ['zh-TW', '繁體中文', 'msg_loading', '載入中...'],
    ['en-US', 'English', 'msg_loading', 'Loading...'],
    
    // 完成度標示
    ['zh-TW', '繁體中文', 'progress', '進度'],
    ['en-US', 'English', 'progress', 'Progress'],
    ['zh-TW', '繁體中文', 'completed', '已完成'],
    ['en-US', 'English', 'completed', 'Completed'],
    ['zh-TW', '繁體中文', 'not_started', '未開始'],
    ['en-US', 'English', 'not_started', 'Not Started'],
    ['zh-TW', '繁體中文', 'in_progress', '進行中'],
    ['en-US', 'English', 'in_progress', 'In Progress']
  ];
}

// ==================== 語系載入功能 ====================

/**
 * 載入指定語系的所有翻譯
 */
function loadLanguage(langCode = 'zh-TW') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.LANGUAGE);
    
    if (!sheet) {
      Logger.log('多語系設定工作表不存在，使用預設值');
      return getDefaultTranslations(langCode);
    }
    
    const data = sheet.getDataRange().getValues();
    const translations = {};
    
    for (let i = 1; i < data.length; i++) {
      const [code, , key, value] = data[i];
      
      if (code === langCode) {
        translations[key] = value;
      }
    }
    
    return translations;
    
  } catch (error) {
    logSystem('載入語系失敗：' + error.message, 'ERROR');
    return getDefaultTranslations(langCode);
  }
}

/**
 * 取得預設翻譯（當工作表不存在時使用）
 */
function getDefaultTranslations(langCode) {
  const defaultData = getDefaultLanguageData();
  const translations = {};
  
  defaultData.forEach(([code, , key, value]) => {
    if (code === langCode) {
      translations[key] = value;
    }
  });
  
  return translations;
}

/**
 * 翻譯單一鍵值
 */
function translate(key, langCode = 'zh-TW') {
  const translations = loadLanguage(langCode);
  return translations[key] || key;
}

/**
 * 取得當前語系
 */
function getCurrentLanguage() {
  try {
    const properties = PropertiesService.getUserProperties();
    return properties.getProperty('language') || 'zh-TW';
  } catch (error) {
    return 'zh-TW';
  }
}

/**
 * 設定當前語系
 */
function setLanguage(langCode) {
  try {
    const properties = PropertiesService.getUserProperties();
    properties.setProperty('language', langCode);
    return {success: true};
  } catch (error) {
    logSystem('設定語系失敗：' + error.message, 'ERROR');
    return {success: false, error: error.message};
  }
}

/**
 * 取得所有可用語系
 */
function getAvailableLanguages() {
  return [
    {code: 'zh-TW', name: '繁體中文'},
    {code: 'en-US', name: 'English'}
  ];
}

// ==================== Web App 使用的語系函數 ====================

/**
 * 為 Web App 提供翻譯資料（JSON 格式）
 */
function getTranslationsForWeb(langCode = 'zh-TW') {
  const translations = loadLanguage(langCode);
  return JSON.stringify(translations);
}

/**
 * 切換語系（供 Web App 呼叫）
 */
function switchLanguageWeb(langCode) {
  setLanguage(langCode);
  return {
    success: true,
    langCode: langCode,
    translations: loadLanguage(langCode)
  };
}
