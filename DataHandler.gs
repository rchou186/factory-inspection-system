/**
 * 資料處理模組
 * Data Handler Module
 */

// ==================== 點檢資料提交 ====================

/**
 * 提交點檢資料
 */
function submitInspection(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) {
      throw new Error('點檢記錄表不存在');
    }
    
    const {date, shift, area, itemNo, itemName, status, note, inspector} = data;
    
    // 檢查必要欄位
    if (!date || !shift || !area || !itemNo || !status || !inspector) {
      throw new Error('缺少必要欄位');
    }
    
    // 取得前次狀態
    const previousStatus = getPreviousStatus(date, shift, area, itemNo);
    
    // 產生記錄ID
    const recordId = generateRecordId('IR', new Date(date), shift);
    
    // 準備寫入的資料
    const row = [
      recordId,              // 記錄ID
      date,                  // 日期
      shift,                 // 時段
      area,                  // 區域
      itemNo,                // 項目編號
      itemName,              // 項目名稱
      status,                // 狀態
      previousStatus || '',  // 前次狀態
      note || '',            // 備註
      inspector,             // 點檢人員
      new Date()             // 時間戳記
    ];
    
    // 寫入資料
    sheet.appendRow(row);
    
    // 如果前次狀態是「待追蹤」，更新該筆記錄
    if (previousStatus === '待追蹤') {
      updatePendingRecord(date, shift, area, itemNo, status);
    }
    
    logSystem(`點檢記錄已提交：${area} - ${itemName} (${status})`);
    
    return {
      success: true,
      recordId: recordId,
      message: '點檢記錄已成功提交'
    };
    
  } catch (error) {
    logSystem('提交點檢資料失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 批次提交點檢資料
 */
function submitInspectionBatch(dataArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName(SHEETS.RECORDS);
    const tempHumiditySheet = ss.getSheetByName(SHEETS.TEMP_HUMIDITY);
    
    if (!recordSheet) {
      throw new Error('點檢記錄表不存在');
    }
    
    if (!tempHumiditySheet) {
      throw new Error('溫濕度記錄表不存在');
    }
    
    const inspectionRows = [];
    const tempHumidityRows = [];
    const alertQueue = [];  // 告警佇列，稍後統一發送
    const results = [];
    
    // 效能優化：一次性讀取所有歷史記錄，建立快速查找表
    const previousStatusMap = buildPreviousStatusMap(recordSheet);
    
    dataArray.forEach(data => {
      const {date, shift, area, itemNo, itemName, status, value, type, note, inspector} = data;
      
      // 判斷是否為溫濕度資料
      if (status === 'TEMP' && value && type) {
        // 溫濕度資料 - 寫入 Temperature_Humidity
        const recordId = generateRecordId('TH', new Date(date), shift);
        
        // 驗證溫濕度是否超標
        const validation = validateTempHumidity(itemName, parseFloat(value), type);
        
        const tempRow = [
          recordId,                    // 記錄ID
          date,                        // 日期
          shift,                       // 時段
          area,                        // 區域
          itemName,                    // 測點名稱
          parseFloat(value),           // 數值
          type === 'temperature' ? '℃' : '%RH',  // 單位
          validation.result,           // 判定結果 (正常/異常)
          validation.message || '',    // 超標說明
          note || '',                  // 備註
          inspector,                   // 點檢人員
          validation.alertSent ? 'Y' : 'N'  // 告警狀態
        ];
        
        tempHumidityRows.push(tempRow);
        
        // 如果異常，加入告警佇列（稍後統一發送）
        if (!validation.isNormal && validation.alertSent) {
          alertQueue.push({
            area: area,
            itemName: itemName,
            value: parseFloat(value),
            unit: type === 'temperature' ? '℃' : '%RH',
            standard: validation.standard,
            deviation: validation.deviation,
            level: validation.level,
            inspector: inspector,
            date: date,
            shift: shift
          });
        }
        
        results.push({recordId, status: 'success', type: 'temperature_humidity'});
        
      } else if (status && status !== 'TEMP') {
        // 一般點檢資料 - 寫入 Inspection_Records
        // 從快速查找表取得前次狀態（不再每次讀取工作表）
        const previousStatus = previousStatusMap[`${area}_${itemNo}`] || '';
        const recordId = generateRecordId('IR', new Date(date), shift);
        
        const row = [
          recordId,
          date,
          shift,
          area,
          itemNo,
          itemName,
          status,
          previousStatus,
          note || '',
          inspector,
          new Date()
        ];
        
        inspectionRows.push(row);
        results.push({recordId, status: 'success', type: 'inspection'});
        
        // 更新待追蹤記錄（批次處理）
        if (previousStatus === '待追蹤') {
          updatePendingRecord(date, shift, area, itemNo, status);
        }
      }
    });
    
    // 批次寫入點檢記錄
    if (inspectionRows.length > 0) {
      recordSheet.getRange(recordSheet.getLastRow() + 1, 1, inspectionRows.length, 11).setValues(inspectionRows);
      logSystem(`點檢記錄已提交：${inspectionRows.length} 筆`);
    }
    
    // 批次寫入溫濕度記錄
    if (tempHumidityRows.length > 0) {
      tempHumiditySheet.getRange(tempHumiditySheet.getLastRow() + 1, 1, tempHumidityRows.length, 12).setValues(tempHumidityRows);
      logSystem(`溫濕度記錄已提交：${tempHumidityRows.length} 筆`);
    }
    
    // 統一發送告警（非同步處理，不阻塞回應）
    if (alertQueue.length > 0) {
      // 使用 setTimeout 非同步發送告警，不阻塞回應
      alertQueue.forEach(alertData => {
        try {
          sendTempHumidityAlert(alertData);
        } catch (error) {
          logSystem(`告警發送失敗：${error.message}`, 'WARNING');
        }
      });
    }
    
    logSystem(`批次提交完成：點檢 ${inspectionRows.length} 筆 + 溫濕度 ${tempHumidityRows.length} 筆`);
    
    return {
      success: true,
      count: inspectionRows.length + tempHumidityRows.length,
      inspectionCount: inspectionRows.length,
      tempHumidityCount: tempHumidityRows.length,
      alertCount: alertQueue.length,
      results: results
    };
    
  } catch (error) {
    logSystem('批次提交失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 建立前次狀態快速查找表（效能優化）
 */
function buildPreviousStatusMap(recordSheet) {
  const statusMap = {};
  
  try {
    const data = recordSheet.getDataRange().getValues();
    
    // 如果只有標題列，返回空映射
    if (data.length <= 1) {
      return statusMap;
    }
    
    // 從最新記錄往回建立映射（倒序處理，保留最新狀態）
    // 欄位順序：記錄ID(0), 日期(1), 時段(2), 區域(3), 項目編號(4), 項目名稱(5), 狀態(6)...
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const area = row[3];      // 區域
      const itemNo = row[4];    // 項目編號
      const status = row[6];    // 狀態
      
      // 檢查必要欄位是否存在
      if (!area || !itemNo || !status) {
        continue;
      }
      
      const key = `${area}_${itemNo}`;
      
      // 如果該項目還沒有記錄，則記錄其最新狀態
      if (!statusMap[key]) {
        statusMap[key] = status;
      }
    }
    
    logSystem(`建立狀態查找表完成：${Object.keys(statusMap).length} 個項目`, 'INFO');
    
  } catch (error) {
    logSystem('建立狀態查找表失敗：' + error.message, 'WARNING');
  }
  
  return statusMap;
}

// ==================== 溫濕度資料提交 ====================

/**
 * 提交溫濕度資料
 */
function submitTempHumidity(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TEMP_HUMIDITY);
    
    if (!sheet) {
      throw new Error('溫濕度記錄表不存在');
    }
    
    const {date, shift, area, pointName, value, unit, type, note, inspector} = data;
    
    // 檢查必要欄位
    if (!date || !shift || !area || !pointName || value === undefined || !unit || !inspector) {
      throw new Error('缺少必要欄位');
    }
    
    // 驗證溫濕度是否超標
    const validation = validateTempHumidity(pointName, value, type);
    
    // 產生記錄ID
    const recordId = generateRecordId('TH', new Date(date), shift);
    
    // 準備寫入的資料
    const row = [
      recordId,                        // 記錄ID
      date,                            // 日期
      shift,                           // 時段
      area,                            // 區域
      pointName,                       // 測點名稱
      value,                           // 數值
      unit,                            // 單位
      validation.result,               // 判定結果
      validation.message || '',        // 超標說明
      note || '',                      // 備註
      inspector,                       // 點檢人員
      validation.alertSent ? 'Y' : 'N' // 告警狀態
    ];
    
    // 寫入資料
    sheet.appendRow(row);
    
    // 如果需要告警，發送通知
    if (!validation.isNormal && validation.alertSent) {
      sendTempHumidityAlert({
        area: area,
        itemName: pointName,
        value: value,
        unit: unit,
        standard: validation.standard,
        deviation: validation.deviation,
        level: validation.level,
        inspector: inspector,
        date: date,
        shift: shift,
        note: note || ''
      });
    }
    
    logSystem(`溫濕度記錄已提交：${pointName} = ${value}${unit} (${validation.result})`);
    
    return {
      success: true,
      recordId: recordId,
      validation: validation,
      message: '溫濕度記錄已成功提交'
    };
    
  } catch (error) {
    logSystem('提交溫濕度資料失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== 溫濕度驗證 ====================

/**
 * 驗證溫濕度是否超標
 */
function validateTempHumidity(pointName, value, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TEMP_CRITERIA);
    
    if (!sheet) {
      return {
        result: '未設定',
        isNormal: true,
        alertSent: false,
        message: '',
        standard: {}
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 查找該測點的標準值
    for (let i = 1; i < data.length; i++) {
      const [name, , measureType, , lowerLimit, upperLimit, enableAlert] = data[i];
      
      if (name === pointName) {
        const lower = parseFloat(lowerLimit);
        const upper = parseFloat(upperLimit);
        const val = parseFloat(value);
        
        // 判定結果
        if (val >= lower && val <= upper) {
          return {
            result: '正常',
            isNormal: true,
            alertSent: false,
            message: `正常範圍 ${lower}-${upper}`,
            standard: {
              lower: lower,
              upper: upper,
              unit: type === 'temperature' ? '℃' : '%RH'
            }
          };
        } else {
          let message = '';
          let level = '警告';
          let deviation = 0;
          
          if (val < lower) {
            deviation = lower - val;
            message = `低於下限 ${deviation.toFixed(1)}`;
            if (deviation > (upper - lower) * 0.2) {
              level = '嚴重';
            }
          } else {
            deviation = val - upper;
            message = `超過上限 ${deviation.toFixed(1)}`;
            if (deviation > (upper - lower) * 0.2) {
              level = '嚴重';
            }
          }
          
          return {
            result: '異常',
            isNormal: false,
            alertSent: enableAlert === true || enableAlert === 'TRUE',
            message: message,
            level: level,
            deviation: deviation,
            standard: {
              lower: lower,
              upper: upper,
              unit: type === 'temperature' ? '℃' : '%RH'
            }
          };
        }
      }
    }
    
    return {
      result: '未設定標準',
      isNormal: true,
      alertSent: false,
      message: '此測點未設定標準值',
      standard: {}
    };
    
  } catch (error) {
    logSystem('溫濕度驗證失敗：' + error.message, 'ERROR');
    return {
      result: '驗證失敗',
      isNormal: true,
      alertSent: false,
      message: error.message,
      standard: {}
    };
  }
}

// ==================== 查詢功能 ====================

/**
 * 取得前次狀態
 */
function getPreviousStatus(currentDate, currentShift, area, itemNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    
    // 從最新記錄往回找（倒序）
    for (let i = data.length - 1; i > 0; i--) {
      const [, recordDate, recordShift, recordArea, recordItemNo, , recordStatus] = data[i];
      
      // 找到同一項目的前一次記錄
      if (recordArea === area && recordItemNo === itemNo) {
        // 確保不是同一次點檢
        const isSameInspection = (recordDate === currentDate && recordShift === currentShift);
        if (!isSameInspection) {
          return recordStatus;
        }
      }
    }
    
    return null;
    
  } catch (error) {
    logSystem('取得前次狀態失敗：' + error.message, 'ERROR');
    return null;
  }
}

/**
 * 更新待追蹤記錄
 */
function updatePendingRecord(currentDate, currentShift, area, itemNo, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    
    // 找到最近的「待追蹤」記錄並更新
    for (let i = data.length - 1; i > 0; i--) {
      const [, recordDate, recordShift, recordArea, recordItemNo, , recordStatus] = data[i];
      
      if (recordArea === area && recordItemNo === itemNo && recordStatus === '待追蹤') {
        // 更新該記錄的狀態
        sheet.getRange(i + 1, 7).setValue(newStatus);
        logSystem(`更新待追蹤記錄：第${i+1}行 → ${newStatus}`);
        break;
      }
    }
    
  } catch (error) {
    logSystem('更新待追蹤記錄失敗：' + error.message, 'ERROR');
  }
}

/**
 * 取得點檢歷史記錄
 */
function getInspectionHistory(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = data[i][index];
      });
      
      // 套用篩選條件
      let match = true;
      if (filters.startDate && new Date(record['日期']) < new Date(filters.startDate)) match = false;
      if (filters.endDate && new Date(record['日期']) > new Date(filters.endDate)) match = false;
      if (filters.area && record['區域'] !== filters.area) match = false;
      if (filters.shift && record['時段'] !== filters.shift) match = false;
      if (filters.status && record['狀態'] !== filters.status) match = false;
      
      if (match) {
        records.push(record);
      }
    }
    
    return records;
    
  } catch (error) {
    logSystem('取得點檢歷史失敗：' + error.message, 'ERROR');
    return [];
  }
}

/**
 * 取得今日點檢進度
 */
function getTodayProgress(date, shift) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName(SHEETS.RECORDS);
    const itemSheet = ss.getSheetByName(SHEETS.ITEM_MASTER);
    
    if (!recordSheet || !itemSheet) {
      throw new Error('所需工作表不存在');
    }
    
    const targetDate = formatDate(date);
    
    // 取得所有項目
    const itemData = itemSheet.getDataRange().getValues();
    const totalItems = itemData.length - 1; // 扣除標題列
    
    // 取得已完成的項目
    const recordData = recordSheet.getDataRange().getValues();
    const completedItems = new Set();
    
    for (let i = 1; i < recordData.length; i++) {
      const [, recordDate, recordShift, recordArea, recordItemNo] = recordData[i];
      
      if (formatDate(recordDate) === targetDate && recordShift === shift) {
        completedItems.add(`${recordArea}_${recordItemNo}`);
      }
    }
    
    // 計算各區域進度
    const areaProgress = {};
    const allAreas = getAllInspectionItems();
    
    allAreas.forEach(area => {
      const areaCode = area.code;
      const areaName = area.name_zh;
      const totalCount = area.items.length;
      let completedCount = 0;
      
      area.items.forEach(item => {
        const key = `${areaName}_${item.no}`;
        if (completedItems.has(key)) {
          completedCount++;
        }
      });
      
      areaProgress[areaCode] = {
        name: areaName,
        total: totalCount,
        completed: completedCount,
        percentage: totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0
      };
    });
    
    return {
      totalItems: totalItems,
      completedItems: completedItems.size,
      percentage: totalItems > 0 ? Math.round((completedItems.size / totalItems) * 100) : 0,
      areaProgress: areaProgress
    };
    
  } catch (error) {
    logSystem('取得今日進度失敗：' + error.message, 'ERROR');
    return {
      totalItems: 0,
      completedItems: 0,
      percentage: 0,
      areaProgress: {}
    };
  }
}
