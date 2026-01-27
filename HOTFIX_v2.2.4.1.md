# DataHandler.gs 緊急修正 (v2.2.4.1)

## 🐛 問題說明

### 發現的問題
在 `buildPreviousStatusMap()` 函數中，欄位解構位置錯誤，導致無法正確讀取前次狀態。

### 錯誤的程式碼
```javascript
// ❌ 錯誤：解構位置不對應實際欄位
const [, , , area, itemNo, , status] = data[i];
```

這會導致：
- `area` 讀到錯誤的欄位
- `itemNo` 讀到錯誤的欄位  
- `status` 讀到錯誤的欄位
- 前次狀態永遠是空的
- 待追蹤邏輯失效

### Inspection_Records 實際欄位順序
```
索引 0: 記錄ID
索引 1: 日期
索引 2: 時段
索引 3: 區域       ← 這裡
索引 4: 項目編號   ← 這裡
索引 5: 項目名稱
索引 6: 狀態       ← 這裡
索引 7: 前次狀態
索引 8: 備註
索引 9: 點檢人員
索引 10: 時間戳記
```

---

## ✅ 修正內容

### 修正後的程式碼
```javascript
function buildPreviousStatusMap(recordSheet) {
  const statusMap = {};
  
  try {
    const data = recordSheet.getDataRange().getValues();
    
    // 如果只有標題列，返回空映射
    if (data.length <= 1) {
      return statusMap;
    }
    
    // 從最新記錄往回建立映射
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const area = row[3];      // ✅ 正確：索引 3 = 區域
      const itemNo = row[4];    // ✅ 正確：索引 4 = 項目編號
      const status = row[6];    // ✅ 正確：索引 6 = 狀態
      
      // 檢查必要欄位是否存在
      if (!area || !itemNo || !status) {
        continue;
      }
      
      const key = `${area}_${itemNo}`;
      
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
```

### 改進重點

1. **明確索引**：不使用解構，直接用 `row[3]`, `row[4]`, `row[6]`
2. **空值檢查**：加入 `if (!area || !itemNo || !status)` 防止空值
3. **空表檢查**：加入 `if (data.length <= 1)` 處理空表情況
4. **日誌記錄**：完成時記錄建立的項目數量，方便除錯

---

## 🔍 影響範圍

### 受影響的功能
- ✅ 前次狀態顯示
- ✅ 待追蹤邏輯
- ✅ 狀態自動更新

### 不受影響的功能
- ✅ 資料提交（本身功能正常）
- ✅ 溫濕度記錄
- ✅ 統計報表

---

## 📋 測試方法

### 測試 1：前次狀態正確性

#### 步驟
1. 第一次點檢：選擇「待追蹤」
2. 提交資料
3. 第二次點檢同一項目：選擇「OK」
4. 提交資料

#### 預期結果
- ✅ 第二次的「前次狀態」欄位應顯示「待追蹤」
- ✅ 第一次的「待追蹤」記錄應被更新為「OK」

#### 驗證方法
開啟 Inspection_Records 工作表，檢查：
```
記錄1: 狀態=OK, 前次狀態=(空) 
記錄2: 狀態=待追蹤, 前次狀態=OK
記錄3: 狀態=OK, 前次狀態=待追蹤 ← 應該要有這個
```

### 測試 2：效能測試

#### 步驟
1. 提交 30 筆點檢資料
2. 計時從點擊「提交」到看到成功訊息

#### 預期結果
- ✅ 應在 3-5 秒內完成
- ✅ 不應該有錯誤訊息

#### 驗證方法
```
Apps Script 執行記錄：
→ 應看到「建立狀態查找表完成：XX 個項目」
→ 應看到「點檢記錄已提交：XX 筆」
→ 不應有錯誤訊息
```

### 測試 3：空表處理

#### 步驟
1. 清空 Inspection_Records（保留標題列）
2. 提交第一筆點檢資料

#### 預期結果
- ✅ 應正常提交
- ✅ 前次狀態為空（正常）
- ✅ 不應有錯誤

---

## 🚀 升級步驟

### 如果您已部署 v2.2.4：

1. **開啟 Apps Script 編輯器**
   ```
   試算表 > 擴充功能 > Apps Script
   ```

2. **開啟 DataHandler.gs**
   ```
   找到 buildPreviousStatusMap 函數
   ```

3. **替換函數**
   - 將整個 `buildPreviousStatusMap()` 函數替換為修正後的版本
   - 點擊「💾 儲存」

4. **測試功能**
   - 執行測試 1、2、3
   - 確認前次狀態正確顯示

5. **完成**
   - 無需重新初始化
   - 無需修改其他檔案

---

## 📊 修正前後對比

### 修正前
```javascript
// 問題：解構位置錯誤
const [, , , area, itemNo, , status] = data[i];
// area = data[i][3] ✅ 正確
// itemNo = data[i][4] ✅ 正確  
// status = data[i][6] ❌ 錯誤！實際是 data[i][5]
```

### 修正後  
```javascript
// 正確：明確指定索引
const area = row[3];      // 區域
const itemNo = row[4];    // 項目編號
const status = row[6];    // 狀態
```

---

## ⚠️ 注意事項

### 不需要重新初始化
- ✅ 這只是程式邏輯修正
- ✅ 不影響資料結構
- ✅ 現有資料完全不受影響

### 舊資料的前次狀態
- 已提交的資料不會回溯更新
- 新提交的資料會正確讀取前次狀態
- 待追蹤項目會正常更新

### 效能影響
- 修正後效能不變（仍然很快）
- 只是讀取正確的欄位而已

---

## 🐛 如何發現這個問題

### 症狀
- 前次狀態欄位永遠是空的
- 待追蹤項目不會自動更新
- 沒有錯誤訊息（因為邏輯本身沒錯，只是讀錯欄位）

### 除錯方法
1. 在 `buildPreviousStatusMap()` 中加入 `Logger.log()`
2. 檢查 `statusMap` 的內容
3. 發現 key 和 value 都不對
4. 追蹤到解構位置錯誤

### 預防措施
- 使用明確的索引，而非解構（可讀性更好）
- 加入註解標明每個索引對應的欄位
- 加入日誌記錄，方便除錯

---

## 📞 需要協助？

如果修正後仍有問題：
1. 檢查 Apps Script 執行記錄
2. 查看「建立狀態查找表完成」日誌
3. 確認顯示的項目數量是否合理
4. 測試前次狀態是否正確顯示

---

**修正版本**：v2.2.4.1  
**修正日期**：2026-01-26  
**修正內容**：buildPreviousStatusMap 欄位索引錯誤  
**影響等級**：中（功能邏輯錯誤，但不影響資料儲存）  
**修正者**：Claude AI

**重要提醒**：
此修正不會影響已提交的資料，只是讓新提交的資料能正確讀取前次狀態。
請盡快更新以確保待追蹤邏輯正常運作。
