# 系統更新說明 v2.2.2

## 📋 本次更新內容 (2026-01-26)

### ✅ UI 改進

#### 1. 提交按鈕位置調整
- **改進前**：提交按鈕固定在頁面最底部
- **改進後**：提交按鈕移至「區域選擇」區塊內
- **優點**：
  - 更符合操作邏輯（選完區域後就能看到提交按鈕）
  - 不需要滾動到最底部才能提交
  - 視覺上更清晰，按鈕與區域資訊在同一區塊

### ✅ 資料儲存架構修正

#### 2. 溫濕度資料正確分流
- **問題**：所有資料（包括溫濕度）都寫入 `Inspection_Records` 工作表
- **修正**：
  - **一般點檢項目** → 寫入 `Inspection_Records` 工作表
  - **溫濕度項目** → 寫入 `Temperature_Humidity` 工作表

**資料分流邏輯**：
```
前端提交資料
    ↓
submitInspectionBatch()
    ↓
判斷資料類型
    ├─→ status === 'TEMP' && 有 value → 溫濕度資料
    │   ↓
    │   寫入 Temperature_Humidity 工作表
    │   ↓
    │   自動驗證是否超標
    │   ↓
    │   如果異常 → 發送告警
    │
    └─→ status === 'OK'/'NOK'/'待追蹤' → 一般點檢
        ↓
        寫入 Inspection_Records 工作表
        ↓
        處理待追蹤邏輯
```

#### Temperature_Humidity 工作表結構
| 欄位 | 說明 | 範例 |
|------|------|------|
| 記錄ID | TH開頭的唯一識別碼 | TH20260126AM001 |
| 日期 | 點檢日期 | 2026-01-26 |
| 時段 | 早/晚/其它 | 早 |
| 區域 | 所屬區域 | B1區 |
| 測點名稱 | 溫濕度測點 | 秤料室濕度 |
| 數值 | 測量值 | 45 |
| 單位 | ℃ 或 %RH | %RH |
| 判定結果 | 正常/異常 | 正常 |
| 超標說明 | 異常時的說明 | 超過上限 8.0 |
| 備註 | 點檢人員備註 | 空調已調整 |
| 點檢人員 | 執行人 | 張三 |
| 告警狀態 | Y/N | N |

### 🔧 修改的檔案

#### 1. InspectionForm.html
```html
<!-- 提交按鈕移到區域選擇區塊內 -->
<div class="inspection-card">
  <h6>區域選擇 Area Selection</h6>
  <div class="area-grid">...</div>
  
  <!-- 提交按鈕 -->
  <div class="text-center mt-4">
    <button class="btn btn-primary btn-lg" onclick="submitAllInspection()">
      提交所有點檢資料
    </button>
  </div>
</div>
```

#### 2. DataHandler.gs - submitInspectionBatch()
**核心修改**：
```javascript
// 判斷資料類型並分別處理
if (status === 'TEMP' && value && type) {
  // 溫濕度資料 → Temperature_Humidity 工作表
  const validation = validateTempHumidity(itemName, parseFloat(value), type);
  
  const tempRow = [
    recordId, date, shift, area, itemName,
    parseFloat(value),
    type === 'temperature' ? '℃' : '%RH',
    validation.result,
    validation.message || '',
    note || '',
    inspector,
    validation.alertSent ? 'Y' : 'N'
  ];
  
  tempHumidityRows.push(tempRow);
  
  // 如果異常，發送告警
  if (!validation.isNormal) {
    sendTempHumidityAlert({...});
  }
  
} else if (status && status !== 'TEMP') {
  // 一般點檢 → Inspection_Records 工作表
  const row = [
    recordId, date, shift, area, itemNo, itemName,
    status, previousStatus || '', note || '',
    inspector, new Date()
  ];
  
  inspectionRows.push(row);
  
  // 處理待追蹤邏輯
  if (previousStatus === '待追蹤') {
    updatePendingRecord(...);
  }
}

// 批次寫入各自的工作表
if (inspectionRows.length > 0) {
  recordSheet.getRange(...).setValues(inspectionRows);
}
if (tempHumidityRows.length > 0) {
  tempHumiditySheet.getRange(...).setValues(tempHumidityRows);
}
```

#### 3. DataHandler.gs - validateTempHumidity()
**更新簽章**：
```javascript
// 舊版
function validateTempHumidity(pointName, value)

// 新版
function validateTempHumidity(pointName, value, type)
```

**返回結構優化**：
```javascript
return {
  result: '正常' | '異常',
  isNormal: true | false,
  alertSent: true | false,
  message: '超過上限 8.0',
  level: '警告' | '嚴重',
  deviation: 8.0,
  standard: {
    lower: 30,
    upper: 60,
    unit: '%RH'
  }
};
```

#### 4. SheetInitializer.gs
**移除 Temperature_Humidity 的時間戳記欄位**：
```javascript
// 舊版：13 個欄位（含時間戳記）
const headers = [
  '記錄ID', '日期', '時段', '區域', '測點名稱', '數值', '單位',
  '判定結果', '超標說明', '備註', '點檢人員', '告警狀態', '時間戳記'
];

// 新版：12 個欄位（移除時間戳記）
const headers = [
  '記錄ID', '日期', '時段', '區域', '測點名稱', '數值', '單位',
  '判定結果', '超標說明', '備註', '點檢人員', '告警狀態'
];
```

### 📊 資料流程圖

```
使用者填寫點檢表單
    │
    ├─→ 一般項目：選擇 OK/NOK/待追蹤
    │       ↓
    │   inspectionData[areaCode][itemNo] = {
    │     status: 'OK',
    │     note: '正常'
    │   }
    │
    └─→ 溫濕度項目：輸入數值
            ↓
        inspectionData[areaCode][itemNo] = {
          status: 'TEMP',
          value: 45,
          type: 'humidity',
          note: ''
        }

↓ 點擊「提交所有點檢資料」

submitInspectionBatch(dataArray)
    │
    ├─→ 一般點檢資料
    │   └─→ Inspection_Records 工作表
    │       - 記錄ID：IR20260126AM001
    │       - 狀態：OK/NOK/待追蹤
    │       - 處理待追蹤邏輯
    │
    └─→ 溫濕度資料
        └─→ Temperature_Humidity 工作表
            - 記錄ID：TH20260126AM001
            - 數值 + 單位
            - 自動驗證是否超標
            - 異常時發送告警
```

### 🎯 實際操作範例

#### 點檢 B1區（包含溫濕度項目）

1. **一般項目（前8項）**：
   ```
   1. 物料是否擺放整齊
      → 選擇：OK
      → 備註：已整理完成
   
   2. 地板是否有保持乾淨
      → 選擇：NOK
      → 備註：發現積水
   ```

2. **溫濕度項目（9-10項）**：
   ```
   9. 秤料室濕度 (%RH)
      → 輸入數值：45
      → 系統自動判定：正常（標準30-60）
      → 備註：空調正常
   
   10. 物料室濕度 (%RH)
      → 輸入數值：68
      → 系統自動判定：異常（超過上限8）
      → 發送告警郵件
      → 備註：空調故障，已報修
   ```

3. **提交後的結果**：
   - **Inspection_Records** 新增 8 筆記錄（項目1-8）
   - **Temperature_Humidity** 新增 2 筆記錄（項目9-10）
   - 物料室濕度異常 → 自動發送告警郵件給相關人員

### ✅ 升級步驟

#### 如果您已經部署了舊版本：

1. **備份現有資料**（重要！）
   ```
   選單：🏭 點檢系統 > 📤 匯出資料
   ```

2. **更新程式碼**
   - 開啟 Apps Script 編輯器
   - 更新以下檔案（完整覆蓋）：
     - `InspectionForm.html`
     - `DataHandler.gs`
     - `SheetInitializer.gs`
   - 點擊「💾 儲存」

3. **更新工作表結構**（僅首次需要）
   - 方法一：重新初始化
     ```
     選單：🏭 點檢系統 > 🔄 初始化系統
     注意：不會刪除現有資料
     ```
   
   - 方法二：手動調整
     ```
     開啟 Temperature_Humidity 工作表
     → 刪除「時間戳記」欄位（第13欄）
     → 確保只有12個欄位
     ```

4. **測試功能**
   - 開啟點檢表單
   - 測試一般項目提交
   - 測試溫濕度項目提交
   - 檢查兩個工作表是否正確分流

### ⚠️ 注意事項

#### 相容性說明
- ✅ **向後相容**：舊的點檢記錄不受影響
- ✅ **資料完整性**：所有現有資料保持不變
- ⚠️ **結構更新**：Temperature_Humidity 工作表需要移除時間戳記欄位

#### 資料驗證
更新後請確認：
1. 一般點檢項目出現在 `Inspection_Records`
2. 溫濕度數值出現在 `Temperature_Humidity`
3. 溫濕度異常時收到告警郵件
4. 統計報表正常運作

### 📞 問題排查

#### 問題1：提交後找不到溫濕度資料
**檢查**：
- Temperature_Humidity 工作表是否存在
- 欄位數量是否為 12 個
- 前端是否傳送了 `type` 參數

#### 問題2：溫濕度沒有自動驗證
**檢查**：
- TempHumidity_Criteria 工作表是否有設定標準值
- 測點名稱是否完全一致（包括空格）
- 下限值和上限值是否為數字

#### 問題3：告警郵件沒有發送
**檢查**：
- TempHumidity_Criteria 中「是否啟用告警」是否為 TRUE
- 「告警收件人」是否有填寫 Email
- AlertSystem.gs 中的 sendTempHumidityAlert 函數是否正常

---

**更新版本**：v2.2.2  
**更新日期**：2026-01-26  
**更新內容**：UI優化 + 資料分流修正  
**更新者**：Claude AI
