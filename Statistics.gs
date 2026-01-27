/**
 * 統計分析模組
 * Statistics Module
 */

// ==================== 自訂期間統計 ====================

/**
 * 產生自訂期間報表
 */
function generateCustomPeriodReport(startDate, endDate, periodDesc = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName(SHEETS.RECORDS);
    const statsSheet = ss.getSheetByName(SHEETS.CUSTOM_STATS);
    
    if (!recordSheet || !statsSheet) {
      throw new Error('所需工作表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    // 產生期間描述
    if (!periodDesc) {
      periodDesc = `${formatDate(start)} ~ ${formatDate(end)}`;
    }
    
    // 取得記錄資料
    const data = recordSheet.getDataRange().getValues();
    
    // 統計資料結構：{area: {itemName: {ok, nok, pending}}}
    const stats = {};
    
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, , area, , itemName, status] = data[i];
      const date = new Date(recordDate);
      
      // 檢查是否在期間內
      if (date >= start && date <= end) {
        if (!stats[area]) stats[area] = {};
        if (!stats[area][itemName]) {
          stats[area][itemName] = {ok: 0, nok: 0, pending: 0, total: 0};
        }
        
        stats[area][itemName].total++;
        
        if (status === 'OK') {
          stats[area][itemName].ok++;
        } else if (status === 'NOK') {
          stats[area][itemName].nok++;
        } else if (status === '待追蹤') {
          stats[area][itemName].pending++;
        }
      }
    }
    
    // 準備寫入資料
    const rows = [];
    const statsId = `STAT${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMddHHmmss')}`;
    
    Object.keys(stats).forEach(area => {
      Object.keys(stats[area]).forEach(itemName => {
        const s = stats[area][itemName];
        const nokRatio = s.total > 0 ? (s.nok / s.total * 100).toFixed(2) : 0;
        
        rows.push([
          statsId,
          formatDate(start),
          formatDate(end),
          periodDesc,
          area,
          itemName,
          s.ok,
          s.nok,
          s.pending,
          s.total,
          parseFloat(nokRatio),
          0  // 排名稍後計算
        ]);
      });
    });
    
    // 按 NOK 次數排序
    rows.sort((a, b) => b[7] - a[7]);
    
    // 設定排名
    rows.forEach((row, index) => {
      row[11] = index + 1;
    });
    
    // 寫入統計表
    if (rows.length > 0) {
      statsSheet.getRange(statsSheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    }
    
    logSystem(`自訂期間統計完成：${periodDesc}，共 ${rows.length} 筆`);
    
    return {
      success: true,
      statsId: statsId,
      count: rows.length,
      periodDesc: periodDesc
    };
    
  } catch (error) {
    logSystem('產生自訂期間報表失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== NOK 排名 ====================

/**
 * 取得 NOK 前五名
 */
function rankTopNOK(startDate, endDate, limit = 5) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) {
      throw new Error('點檢記錄表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const data = sheet.getDataRange().getValues();
    
    // 統計 NOK 次數
    const nokCount = {};
    
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, , area, , itemName, status] = data[i];
      const date = new Date(recordDate);
      
      if (date >= start && date <= end && status === 'NOK') {
        const key = `${area} - ${itemName}`;
        
        if (!nokCount[key]) {
          nokCount[key] = {
            area: area,
            itemName: itemName,
            count: 0,
            total: 0
          };
        }
        
        nokCount[key].count++;
      }
      
      // 計算總次數（用於計算比率）
      if (date >= start && date <= end) {
        const key = `${area} - ${itemName}`;
        if (nokCount[key]) {
          nokCount[key].total++;
        }
      }
    }
    
    // 轉換為陣列並排序
    const ranking = Object.keys(nokCount).map(key => {
      const item = nokCount[key];
      const ratio = item.total > 0 ? (item.count / item.total * 100).toFixed(1) : 0;
      
      return {
        rank: 0,
        area: item.area,
        itemName: item.itemName,
        nokCount: item.count,
        totalCount: item.total,
        nokRatio: parseFloat(ratio)
      };
    });
    
    // 排序：主要依 NOK 次數，次要依比率
    ranking.sort((a, b) => {
      if (b.nokCount !== a.nokCount) {
        return b.nokCount - a.nokCount;
      }
      return b.nokRatio - a.nokRatio;
    });
    
    // 設定排名並取前N名
    const topN = ranking.slice(0, limit).map((item, index) => {
      item.rank = index + 1;
      return item;
    });
    
    return topN;
    
  } catch (error) {
    logSystem('NOK 排名失敗：' + error.message, 'ERROR');
    return [];
  }
}

// ==================== 趨勢分析 ====================

/**
 * 產生趨勢圖資料
 */
function generateTrendData(startDate, endDate, areaFilter = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) {
      throw new Error('點檢記錄表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const data = sheet.getDataRange().getValues();
    
    // 按日期統計
    const dailyStats = {};
    
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, , area, , , status] = data[i];
      const date = new Date(recordDate);
      
      if (date >= start && date <= end) {
        // 如果有區域篩選
        if (areaFilter && area !== areaFilter) continue;
        
        const dateKey = formatDate(date);
        
        if (!dailyStats[dateKey]) {
          dailyStats[dateKey] = {ok: 0, nok: 0, pending: 0, total: 0};
        }
        
        dailyStats[dateKey].total++;
        
        if (status === 'OK') {
          dailyStats[dateKey].ok++;
        } else if (status === 'NOK') {
          dailyStats[dateKey].nok++;
        } else if (status === '待追蹤') {
          dailyStats[dateKey].pending++;
        }
      }
    }
    
    // 轉換為陣列格式
    const trendData = Object.keys(dailyStats).sort().map(dateKey => {
      const stats = dailyStats[dateKey];
      const okRate = stats.total > 0 ? (stats.ok / stats.total * 100).toFixed(1) : 0;
      const nokRate = stats.total > 0 ? (stats.nok / stats.total * 100).toFixed(1) : 0;
      
      return {
        date: dateKey,
        ok: stats.ok,
        nok: stats.nok,
        pending: stats.pending,
        total: stats.total,
        okRate: parseFloat(okRate),
        nokRate: parseFloat(nokRate)
      };
    });
    
    return trendData;
    
  } catch (error) {
    logSystem('產生趨勢資料失敗：' + error.message, 'ERROR');
    return [];
  }
}

/**
 * 產生圓餅圖資料
 */
function generatePieData(startDate, endDate, areaFilter = null, shiftFilter = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.RECORDS);
    
    if (!sheet) {
      throw new Error('點檢記錄表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const data = sheet.getDataRange().getValues();
    
    const stats = {ok: 0, nok: 0, pending: 0, total: 0};
    
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, shift, area, , , status] = data[i];
      const date = new Date(recordDate);
      
      if (date >= start && date <= end) {
        // 套用篩選條件
        if (areaFilter && area !== areaFilter) continue;
        if (shiftFilter && shift !== shiftFilter) continue;
        
        stats.total++;
        
        if (status === 'OK') {
          stats.ok++;
        } else if (status === 'NOK') {
          stats.nok++;
        } else if (status === '待追蹤') {
          stats.pending++;
        }
      }
    }
    
    // 計算百分比
    const pieData = [
      {
        label: 'OK',
        value: stats.ok,
        percentage: stats.total > 0 ? (stats.ok / stats.total * 100).toFixed(1) : 0,
        color: '#4CAF50'
      },
      {
        label: 'NOK',
        value: stats.nok,
        percentage: stats.total > 0 ? (stats.nok / stats.total * 100).toFixed(1) : 0,
        color: '#F44336'
      },
      {
        label: '待追蹤',
        value: stats.pending,
        percentage: stats.total > 0 ? (stats.pending / stats.total * 100).toFixed(1) : 0,
        color: '#FFC107'
      }
    ];
    
    return {
      data: pieData,
      total: stats.total
    };
    
  } catch (error) {
    logSystem('產生圓餅圖資料失敗：' + error.message, 'ERROR');
    return {data: [], total: 0};
  }
}

// ==================== 溫濕度趨勢分析 ====================

/**
 * 分析溫濕度趨勢
 */
function analyzeTempHumidityTrend(pointName, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TEMP_HUMIDITY);
    
    if (!sheet) {
      throw new Error('溫濕度記錄表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const data = sheet.getDataRange().getValues();
    
    const trendData = [];
    let normalCount = 0;
    let abnormalCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const [, recordDate, shift, , name, value, unit, result] = data[i];
      const date = new Date(recordDate);
      
      if (name === pointName && date >= start && date <= end) {
        trendData.push({
          date: formatDate(date),
          shift: shift,
          value: parseFloat(value),
          unit: unit,
          result: result
        });
        
        if (result === '正常') {
          normalCount++;
        } else if (result === '異常') {
          abnormalCount++;
        }
      }
    }
    
    // 排序
    trendData.sort((a, b) => a.date.localeCompare(b.date));
    
    // 計算統計值
    const values = trendData.map(d => d.value);
    const avg = values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    const max = values.length > 0 ? Math.max(...values) : 0;
    const min = values.length > 0 ? Math.min(...values) : 0;
    
    return {
      pointName: pointName,
      data: trendData,
      statistics: {
        count: trendData.length,
        normalCount: normalCount,
        abnormalCount: abnormalCount,
        normalRate: trendData.length > 0 ? (normalCount / trendData.length * 100).toFixed(1) : 0,
        average: avg.toFixed(2),
        max: max.toFixed(2),
        min: min.toFixed(2)
      }
    };
    
  } catch (error) {
    logSystem('分析溫濕度趨勢失敗：' + error.message, 'ERROR');
    return null;
  }
}

// ==================== KPI 計算 ====================

/**
 * 計算點檢 KPI 指標
 */
function calculateKPI(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName(SHEETS.RECORDS);
    const itemSheet = ss.getSheetByName(SHEETS.ITEM_MASTER);
    
    if (!recordSheet || !itemSheet) {
      throw new Error('所需工作表不存在');
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    // 計算期間天數
    const days = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
    
    // 取得項目總數
    const itemData = itemSheet.getDataRange().getValues();
    const totalItems = itemData.length - 1;
    
    // 計算應點檢次數（每日早晚兩次）
    const expectedInspections = totalItems * days * 2;
    
    // 統計實際點檢資料
    const recordData = recordSheet.getDataRange().getValues();
    let actualInspections = 0;
    let okCount = 0;
    let nokCount = 0;
    let pendingCount = 0;
    
    for (let i = 1; i < recordData.length; i++) {
      const [, recordDate, , , , , status] = recordData[i];
      const date = new Date(recordDate);
      
      if (date >= start && date <= end) {
        actualInspections++;
        
        if (status === 'OK') okCount++;
        else if (status === 'NOK') nokCount++;
        else if (status === '待追蹤') pendingCount++;
      }
    }
    
    // 計算 KPI
    const completionRate = expectedInspections > 0 ? (actualInspections / expectedInspections * 100).toFixed(2) : 0;
    const qualifiedRate = actualInspections > 0 ? (okCount / actualInspections * 100).toFixed(2) : 0;
    const defectRate = actualInspections > 0 ? (nokCount / actualInspections * 100).toFixed(2) : 0;
    
    return {
      period: {
        startDate: formatDate(start),
        endDate: formatDate(end),
        days: days
      },
      inspections: {
        expected: expectedInspections,
        actual: actualInspections,
        completionRate: parseFloat(completionRate)
      },
      status: {
        ok: okCount,
        nok: nokCount,
        pending: pendingCount
      },
      rates: {
        qualified: parseFloat(qualifiedRate),
        defect: parseFloat(defectRate)
      }
    };
    
  } catch (error) {
    logSystem('計算 KPI 失敗：' + error.message, 'ERROR');
    return null;
  }
}

// ==================== 匯出統計資料 ====================

/**
 * 匯出統計資料為 CSV
 */
function exportStatistics(startDate, endDate) {
  try {
    const data = getInspectionHistory({
      startDate: startDate,
      endDate: endDate
    });
    
    if (data.length === 0) {
      return {success: false, message: '無資料可匯出'};
    }
    
    // 產生 CSV 內容
    const headers = Object.keys(data[0]);
    let csv = headers.join(',') + '\n';
    
    data.forEach(record => {
      const row = headers.map(header => {
        const value = record[header];
        // 處理包含逗號的內容
        return typeof value === 'string' && value.includes(',') ? `"${value}"` : value;
      });
      csv += row.join(',') + '\n';
    });
    
    // 建立檔案
    const fileName = `inspection_statistics_${formatDate(new Date(), 'yyyyMMdd')}.csv`;
    const blob = Utilities.newBlob(csv, 'text/csv', fileName);
    
    return {
      success: true,
      fileName: fileName,
      blob: blob
    };
    
  } catch (error) {
    logSystem('匯出統計資料失敗：' + error.message, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}
