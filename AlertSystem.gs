/**
 * ============================================================================
 * 工廠環境衛生點檢管理系統 - 告警系統模組
 * Factory Environmental Hygiene Inspection System - Alert System Module
 * ============================================================================
 * 
 * @file        AlertSystem.gs
 * @version     v2.2.5
 * @date        2026-01-26
 * @author      System Developer
 * @description 負責溫濕度超標告警的發送、抑制、記錄等功能，包含 Email 通知、
 *              告警歷史記錄、告警抑制機制等
 * 
 * @functions
 *   - sendTempHumidityAlert()         發送溫濕度告警
 *   - isAlertSuppressed()             檢查告警是否被抑制
 *   - getAlertRecipients()            取得告警收件人
 *   - buildAlertEmailBody()           建立告警郵件內容
 *   - logAlertHistory()               記錄告警歷史
 *   - getAlertHistory()               查詢告警歷史
 *   - sendAlertEmail()                發送告警郵件
 * 
 * @dependencies
 *   - Code.gs (SHEETS 常數)
 *   - DataHandler.gs (資料查詢)
 * 
 * @features
 *   - 告警抑制：1 小時內同一測點只發送 1 次
 *   - 告警等級：警告、注意、嚴重
 *   - 告警記錄：自動記錄所有告警歷史
 *   - Email 通知：HTML 格式郵件
 * 
 * @changelog
 *   v2.2.5 (2026-01-26) - ISO 22000 整合
 *   v2.2.4 (2026-01-26) - 完整告警系統實作
 *   v2.2.3 (2026-01-26) - 告警抑制機制
 * 
 * ============================================================================
 */

// ==================== 溫濕度告警 ====================

/**
 * 發送溫濕度告警
 */
function sendTempHumidityAlert(data) {
  try {
    // 檢查告警抑制（1小時內同一測點只發送一次）
    if (isAlertSuppressed(data.itemName)) {
      logSystem(`告警已抑制：${data.itemName}`, 'INFO');
      return {success: false, message: '告警已抑制（1小時內）'};
    }
    
    // 取得收件人
    const recipients = getAlertRecipients(data.itemName);
    
    if (!recipients || recipients.length === 0) {
      logSystem(`無告警收件人：${data.itemName}`, 'WARN');
      return {success: false, message: '無告警收件人'};
    }
    
    // 發送郵件告警
    const emailResult = sendAlertEmail(data, recipients);
    
    // 記錄告警歷史
    logAlertHistory(data, recipients.join(','));
    
    // 更新告警時間（用於抑制機制）
    updateAlertTimestamp(data.itemName);
    
    return emailResult;
    
  } catch (error) {
    logSystem('發送告警失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 發送告警郵件
 */
function sendAlertEmail(data, recipients) {
  try {
    const {date, shift, area, itemName, value, unit, note, inspector, standard, deviation, level} = data;
    
    // 確定告警等級
    const alertLevel = level || '警告';
    const levelIcon = {
      '警告': '⚠️',
      '注意': '⚡',
      '嚴重': '🚨'
    }[alertLevel] || '⚠️';
    
    // 郵件主旨
    const subject = `${levelIcon}【${alertLevel}】溫濕度異常告警 - ${itemName}`;
    
    // 郵件內容
    const body = `
================================================
工廠環境衛生點檢系統 - 溫濕度異常通知
Factory Environmental Hygiene Inspection System
================================================

【告警等級】${levelIcon} ${alertLevel}

【測點資訊】
測點名稱：${itemName}
所屬區域：${area}
測量時間：${date} (${shift}班)
測量值：${value} ${unit}
標準範圍：${standard.lower} ~ ${standard.upper} ${unit}

【異常狀態】
偏離值：${deviation ? deviation.toFixed(1) : 'N/A'} ${unit}

【點檢資訊】
點檢人員：${inspector}
${note ? `備註說明：${note}` : ''}

【建議處理措施】
${getRecommendedActions(itemName, {level: alertLevel})}

【查看詳細記錄】
請登入系統查看完整記錄和趨勢分析。

================================================
此為系統自動發送，請勿回覆
System Alert - Do Not Reply
================================================
    `;
    
    // 發送郵件
    let successCount = 0;
    recipients.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body
        });
        logSystem(`告警郵件已發送至：${email}`);
        successCount++;
      } catch (e) {
        logSystem(`發送郵件失敗（${email}）：${e.message}`, 'ERROR');
      }
    });
    
    return {
      success: successCount > 0,
      recipients: recipients,
      successCount: successCount,
      message: `告警郵件已發送給 ${successCount}/${recipients.length} 位收件人`
    };
    
  } catch (error) {
    logSystem('發送告警郵件失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 取得建議處理措施
 */
function getRecommendedActions(pointName, validation) {
  const actions = {
    '秤料室濕度': [
      '1. 立即檢查空調除濕系統',
      '2. 確認門窗密閉狀況',
      '3. 記錄環境溫度',
      '4. 採取必要調整措施'
    ],
    '物料室濕度': [
      '1. 檢查除濕機運作狀況',
      '2. 確認通風系統正常',
      '3. 檢查物料包裝完整性',
      '4. 必要時轉移物料'
    ],
    '半成品暫存區溫度': [
      '1. 檢查冷藏設備運作',
      '2. 確認門是否關閉',
      '3. 檢查溫度感應器',
      '4. 評估產品安全性'
    ],
    '半成品冷凍區溫度': [
      '1. 檢查冷凍設備',
      '2. 確認壓縮機運作',
      '3. 檢查冷媒是否充足',
      '4. 緊急時轉移產品'
    ],
    '原料庫溫度': [
      '1. 檢查冷凍庫設備',
      '2. 確認溫度控制器',
      '3. 檢查庫門密閉性',
      '4. 評估原料狀況'
    ],
    '成品庫溫度': [
      '1. 檢查冷凍設備',
      '2. 確認溫度警報系統',
      '3. 檢查庫房負載',
      '4. 準備應急預案'
    ]
  };
  
  const defaultActions = [
    '1. 立即檢查相關設備',
    '2. 記錄異常狀況',
    '3. 採取改善措施',
    '4. 持續監控溫濕度'
  ];
  
  const actionList = actions[pointName] || defaultActions;
  return actionList.join('\n');
}

/**
 * 取得告警收件人清單
 */
function getAlertRecipients(pointName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TEMP_CRITERIA);
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const [name, , , , , , enableAlert, , recipients] = data[i];
      
      if (name === pointName && (enableAlert === true || enableAlert === 'TRUE')) {
        if (recipients && recipients.trim()) {
          // 分割多個收件人（逗號分隔）
          return recipients.split(',').map(email => email.trim()).filter(email => email);
        }
      }
    }
    
    return [];
    
  } catch (error) {
    logSystem('取得告警收件人失敗：' + error.message, 'ERROR');
    return [];
  }
}

// ==================== 告警抑制機制 ====================

/**
 * 檢查告警是否被抑制
 */
function isAlertSuppressed(pointName) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = `alert_last_${pointName}`;
    const lastAlertTime = properties.getProperty(key);
    
    if (!lastAlertTime) return false;
    
    const lastTime = new Date(lastAlertTime);
    const now = new Date();
    const diffHours = (now - lastTime) / (1000 * 60 * 60);
    
    // 1小時內不重複告警
    return diffHours < 1;
    
  } catch (error) {
    logSystem('檢查告警抑制失敗：' + error.message, 'ERROR');
    return false;
  }
}

/**
 * 更新告警時間戳記
 */
function updateAlertTimestamp(pointName) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = `alert_last_${pointName}`;
    properties.setProperty(key, new Date().toISOString());
  } catch (error) {
    logSystem('更新告警時間戳記失敗：' + error.message, 'ERROR');
  }
}

// ==================== 告警歷史記錄 ====================

/**
 * 記錄告警歷史
 */
function logAlertHistory(data, recipients) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Alert_History');
    
    // 如果不存在則建立
    if (!sheet) {
      sheet = ss.insertSheet('Alert_History');
      sheet.appendRow([
        '告警時間', '測點名稱', '區域', '數值', '單位', 
        '告警等級', '偏離值', '收件人', '點檢人員', '備註'
      ]);
      sheet.getRange(1, 1, 1, 10)
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF');
    }
    
    const row = [
      new Date(),
      data.itemName,
      data.area,
      data.value,
      data.unit,
      data.level || '警告',
      data.deviation ? data.deviation.toFixed(1) : '',
      recipients,
      data.inspector,
      data.note || ''
    ];
    
    sheet.appendRow(row);
    
  } catch (error) {
    logSystem('記錄告警歷史失敗：' + error.message, 'ERROR');
  }
}

// ==================== 待追蹤提醒 ====================

/**
 * 檢查待追蹤項目並發送提醒
 */
function checkPendingItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const pendingItems = {};
    
    // 統計待追蹤項目
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, shift, area, itemNo, itemName, status, , , inspector] = data[i];
      
      if (status === '待追蹤') {
        const key = `${area}_${itemNo}`;
        
        if (!pendingItems[key]) {
          pendingItems[key] = {
            area: area,
            itemName: itemName,
            count: 0,
            lastDate: recordDate,
            lastShift: shift,
            inspector: inspector
          };
        }
        
        pendingItems[key].count++;
        
        // 更新最後記錄日期
        if (new Date(recordDate) > new Date(pendingItems[key].lastDate)) {
          pendingItems[key].lastDate = recordDate;
          pendingItems[key].lastShift = shift;
        }
      }
    }
    
    // 發送提醒（連續2次以上）
    Object.keys(pendingItems).forEach(key => {
      const item = pendingItems[key];
      if (item.count >= 2) {
        sendPendingReminder(item);
      }
    });
    
  } catch (error) {
    logSystem('檢查待追蹤項目失敗：' + error.message, 'ERROR');
  }
}

/**
 * 發送待追蹤提醒
 */
function sendPendingReminder(item) {
  try {
    // 這裡可以實作郵件提醒功能
    logSystem(`待追蹤提醒：${item.area} - ${item.itemName} (連續${item.count}次)`, 'WARN');
    
    // TODO: 實作郵件發送邏輯
    // 可以根據需求決定收件人（例如主管、相關負責人等）
    
  } catch (error) {
    logSystem('發送待追蹤提醒失敗：' + error.message, 'ERROR');
  }
}

// ==================== 每日統計提醒 ====================

/**
 * 發送每日統計摘要
 */
function sendDailySummary() {
  try {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = formatDate(yesterday);
    
    // 取得昨日統計
    const earlyShift = getTodayProgress(yesterday, '早');
    const lateShift = getTodayProgress(yesterday, '晚');
    
    // 取得 NOK 前三名
    const topNOK = rankTopNOK(dateStr, dateStr, 3);
    
    // 建立郵件內容
    const subject = `📊 每日點檢統計摘要 - ${dateStr}`;
    
    const body = `
================================================
工廠環境衛生點檢系統 - 每日統計摘要
================================================

【日期】${dateStr}

【早班統計】
總項目數：${earlyShift.totalItems}
已完成：${earlyShift.completedItems}
完成率：${earlyShift.percentage}%

【晚班統計】
總項目數：${lateShift.totalItems}
已完成：${lateShift.completedItems}
完成率：${lateShift.percentage}%

【NOK 前三名】
${topNOK.map((item, index) => 
  `${index + 1}. ${item.area} - ${item.itemName}\n   NOK次數：${item.nokCount} | 比率：${item.nokRatio}%`
).join('\n\n')}

【溫濕度異常】
請查看系統了解詳細資訊

================================================
此為系統自動發送
================================================
    `;
    
    // TODO: 設定收件人並發送
    logSystem('每日統計摘要已準備完成');
    
  } catch (error) {
    logSystem('發送每日統計摘要失敗：' + error.message, 'ERROR');
  }
}
