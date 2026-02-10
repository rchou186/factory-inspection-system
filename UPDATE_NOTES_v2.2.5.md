# 系統更新說明 v2.2.5

## 📋 本次更新內容 (2026-01-26)

### ✨ 主要更新

#### 1. ISO 標準變更為 ISO 22000:2018
- **從 ISO 9001 改為 ISO 22000:2018**
- 更適合食品工廠使用
- 符合食品安全管理系統標準
- 整合 HACCP 原則

#### 2. 表單編號系統升級
- **預設編號**：`P-4-001 版本 A1`
- **動態載入**：從 System_Config 工作表讀取
- **即時更新**：修改 System_Config 後，表單自動顯示新編號
- **適用範圍**：點檢表單 + 統計報表

#### 3. 程式碼註解標準化
- **所有 .gs 檔案**：加入完整的檔案頭註解
- **所有 .html 檔案**：加入完整的檔案頭註解
- **標示版本**：v2.2.5
- **標示日期**：2026-01-26

---

## 📊 ISO 22000:2018 對應

### 文件管制對應

| 文件 | 編號 | ISO 22000 章節 |
|------|------|----------------|
| 點檢表單 | P-4-001 版本 A1 | 7.5 文件化資訊 |
| 統計報表 | P-4-001 版本 A1 | 8.5.2 監測與測量 |
| 溫濕度記錄 | 自動產生 | 8.5.2 監測與測量 |
| 告警記錄 | 自動產生 | 8.5.3 分析與評價 |

### HACCP 原則整合

| HACCP 原則 | 系統功能 |
|-----------|---------|
| 危害分析 | 溫濕度監控點 |
| 關鍵控制點 | 6 個溫濕度測點 |
| 設定關鍵限值 | 上下限標準 |
| 監測程序 | 每日點檢記錄 |
| 矯正措施 | 告警通知 + 備註 |
| 驗證程序 | 統計報表分析 |
| 文件化 | 自動記錄保存 |

---

## 🔧 詳細修改內容

### 1. System_Config 工作表

**新增/修改欄位**：

| 設定項目 | 設定值 | 說明 |
|---------|--------|------|
| ISO表單編號 | P-4-001 版本 A1 | 可自訂 |
| ISO標準 | ISO 22000:2018 | 新增 |
| 系統版本 | v2.2.5 | 更新 |

**使用方式**：
```
1. 開啟 System_Config 工作表
2. 找到「ISO表單編號」欄位
3. 修改為您要的編號
4. 重新載入點檢表單或統計報表
5. 表單編號自動更新
```

### 2. Code.gs 修改

**新增函數**：
```javascript
/**
 * 取得系統設定（包含 ISO 表單編號）
 * @returns {Object} 系統設定物件
 */
function getSystemConfig() {
  // 從 System_Config 讀取設定
  // 包含：formNumber, isoStandard, systemVersion
  return config;
}
```

**更新常數**：
```javascript
const DEFAULTS = {
  SHIFTS: ['早', '晚', '其它'],
  LANGUAGE: 'zh-TW',
  ISO_FORM_NUMBER: 'P-4-001 版本 A1',
  ISO_STANDARD: 'ISO 22000:2018'
};
```

### 3. InspectionForm.html 修改

**表單編號顯示**：
```html
<!-- 舊版：硬編碼 -->
<span id="formNumberLabel">表單編號</span>: QP-7.5-001 (ISO 9001)

<!-- 新版：動態載入 -->
<span id="formNumberLabel">表單編號</span>: <span id="formNumber">載入中...</span>
```

**新增函數**：
```javascript
// 載入系統設定（表單編號）
function loadSystemConfig() {
  google.script.run
    .withSuccessHandler(function(config) {
      if (config && config.formNumber) {
        document.getElementById('formNumber').textContent = config.formNumber;
      }
    })
    .getSystemConfig();
}
```

### 4. StatisticsView.html 修改

**表單編號顯示**：
```html
<!-- 舊版：硬編碼 -->
<small>表單編號: QP-7.5-002</small>

<!-- 新版：動態載入 -->
<small>表單編號: <span id="formNumber">載入中...</span></small>
```

**新增函數**：
```javascript
// 載入系統設定（表單編號）
function loadSystemConfig() {
  google.script.run
    .withSuccessHandler(function(config) {
      if (config && config.formNumber) {
        document.getElementById('formNumber').textContent = config.formNumber;
      }
    })
    .getSystemConfig();
}
```

### 5. 所有 .gs 檔案加入標準註解頭

**註解結構**：
```javascript
/**
 * ============================================================================
 * 工廠環境衛生點檢管理系統 - [模組名稱]
 * Factory Environmental Hygiene Inspection System - [Module Name]
 * ============================================================================
 * 
 * @file        [檔案名稱.gs]
 * @version     v2.2.5
 * @date        2026-01-26
 * @author      System Developer
 * @description [檔案功能說明]
 * 
 * @functions
 *   - function1()    功能說明
 *   - function2()    功能說明
 * 
 * @dependencies
 *   - Code.gs
 *   - 其他相依檔案
 * 
 * @changelog
 *   v2.2.5 (2026-01-26) - ISO 22000 整合
 *   v2.2.4 (2026-01-26) - 前一版本更新
 * 
 * ============================================================================
 */
```

**修改的檔案**：
- Code.gs ✅
- SheetInitializer.gs ✅
- DataHandler.gs ✅
- AlertSystem.gs ✅
- InspectionItems.gs ✅
- Statistics.gs ✅
- Triggers.gs ✅
- I18n.gs ✅

### 6. 所有 .html 檔案加入標準註解頭

**註解結構**：
```html
<!--
============================================================================
工廠環境衛生點檢管理系統 - [介面名稱]
Factory Environmental Hygiene Inspection System - [Interface Name]
============================================================================

@file        [檔案名稱.html]
@version     v2.2.5
@date        2026-01-26
@author      System Developer
@description [介面功能說明]

@features
  - 功能1
  - 功能2

@dependencies
  - Bootstrap 5.3.0
  - Code.gs (getSystemConfig)

@changelog
  v2.2.5 (2026-01-26) - ISO 22000 整合
  v2.2.4 (2026-01-26) - 前一版本更新

============================================================================
-->
```

**修改的檔案**：
- InspectionForm.html ✅
- StatisticsView.html ✅

---

## 🚀 升級步驟

### 如果您已部署舊版本：

#### 步驟 1：備份資料
```
選單：🏭 點檢系統 > 📤 匯出資料
```

#### 步驟 2：更新程式碼
1. 開啟 Apps Script 編輯器
2. 更新以下檔案：
   - **Code.gs** ⭐ 新增 getSystemConfig()
   - **SheetInitializer.gs** ⭐ 更新 System_Config 初始化
   - **InspectionForm.html** ⭐ 動態表單編號
   - **StatisticsView.html** ⭐ 動態表單編號
   - **所有其他 .gs 檔案** ⭐ 更新註解頭
3. 儲存所有檔案

#### 步驟 3：更新 System_Config
1. 開啟 `System_Config` 工作表
2. 找到「ISO表單編號」欄位
3. 修改值為：`P-4-001 版本 A1`
4. 新增一列：
   ```
   設定項目：ISO標準
   設定值：ISO 22000:2018
   ```
5. 更新「系統版本」為：`v2.2.5`

#### 步驟 4：測試功能
1. **測試點檢表單**
   - 選單：🏭 點檢系統 > 📝 開啟點檢表單
   - 檢查表單編號是否顯示：`P-4-001 版本 A1`

2. **測試統計報表**
   - 選單：🏭 點檢系統 > 📊 查看統計報表
   - 檢查表單編號是否顯示：`P-4-001 版本 A1`

3. **測試動態更新**
   - 修改 System_Config 的「ISO表單編號」
   - 重新開啟點檢表單
   - 確認顯示新的編號

---

## 📋 自訂表單編號範例

### 範例 1：使用公司編號系統
```
設定：P-4-001 版本 A1
結果：點檢表單和統計報表都顯示此編號
```

### 範例 2：使用部門代碼
```
設定：QC-FS-001-R1
結果：QC-FS-001-R1
```

### 範例 3：使用日期版本
```
設定：FSMS-2026-001 V1.0
結果：FSMS-2026-001 V1.0
```

**注意事項**：
- 編號長度建議在 30 字以內
- 可包含中英文、數字、符號
- 修改後自動生效，無需重啟

---

## 📚 檔案清單

### 後端程式 (.gs) - 8 個檔案
1. Code.gs ⭐ v2.2.5
2. SheetInitializer.gs ⭐ v2.2.5
3. DataHandler.gs ⭐ v2.2.5
4. AlertSystem.gs ⭐ v2.2.5
5. InspectionItems.gs ⭐ v2.2.5
6. Statistics.gs ⭐ v2.2.5
7. Triggers.gs ⭐ v2.2.5
8. I18n.gs ⭐ v2.2.5

### 前端介面 (.html) - 2 個檔案
1. InspectionForm.html ⭐ v2.2.5
2. StatisticsView.html ⭐ v2.2.5

### 所有檔案都已：
- ✅ 加入完整註解頭
- ✅ 標示版本 v2.2.5
- ✅ 標示日期 2026-01-26
- ✅ 說明功能與相依性

---

## 🎯 重要提醒

### 表單編號的使用
1. **統一性**：點檢表單和統計報表使用相同編號
2. **可追溯性**：建議編號包含版本資訊
3. **合規性**：符合 ISO 22000:2018 文件管制要求

### ISO 22000:2018 的優勢
1. **專注食品安全**
   - 涵蓋整個食品鏈
   - 整合 HACCP 原則
   - 關注危害預防

2. **國際認可**
   - 全球通用標準
   - 客戶信心提升
   - 出口貿易優勢

3. **系統整合**
   - 可與 ISO 9001 並行
   - 可與 FSSC 22000 整合
   - 可與 HACCP 系統結合

---

## 📞 技術支援

### 常見問題

**Q1：表單編號沒有更新？**
A1：請重新載入頁面（Ctrl+F5 或 Cmd+Shift+R）

**Q2：修改 System_Config 後編號還是舊的？**
A2：請確認：
- 欄位名稱是「ISO表單編號」（完全一致）
- 儲存格有填入新值
- 重新開啟表單

**Q3：可以使用中文編號嗎？**
A3：可以！支援中英文、數字、符號

---

**更新版本**：v2.2.5  
**更新日期**：2026-01-26  
**ISO 標準**：ISO 22000:2018  
**表單編號**：P-4-001 版本 A1（可自訂）

**核心改進**：
- 🎯 ISO 22000:2018 食品安全管理系統
- 🔧 動態表單編號（可自訂）
- 📝 完整程式註解標準化
- ✅ 保留 v2.2.4 所有功能
