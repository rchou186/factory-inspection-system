# 檔案名稱對照表

## 📝 重要說明

由於 Google Apps Script 不允許 .gs 和 .html 檔案使用相同的名稱，因此統計報表的 HTML 檔案名稱已調整。

## 📋 檔案清單

### Google Apps Script 檔案 (.gs)
| 檔案名稱 | 說明 |
|---------|------|
| Code.gs | 系統主程式與入口 |
| SheetInitializer.gs | 工作表初始化模組 |
| InspectionItems.gs | 點檢項目資料（105項+6溫濕度點）|
| DataHandler.gs | 資料處理與提交 |
| **Statistics.gs** | **統計分析與報表模組** |
| AlertSystem.gs | 溫濕度告警系統 |
| Triggers.gs | 自動化觸發器 |
| I18n.gs | 多語系支援 |

### HTML 檔案
| 檔案名稱 | 說明 | 對應頁面 |
|---------|------|----------|
| InspectionForm.html | 點檢表單介面 | page=inspection |
| **StatisticsView.html** | **統計報表介面** | **page=statistics** |

## ⚠️ 注意事項

在 Google Apps Script 編輯器中建立 HTML 檔案時：

1. **InspectionForm.html**
   - 建立時輸入名稱：`InspectionForm`（不含 .html）
   
2. **StatisticsView.html**
   - 建立時輸入名稱：`StatisticsView`（不含 .html）
   - ⚠️ **不要命名為 `Statistics`**，因為已經有 `Statistics.gs` 檔案

## 🔗 檔案引用關係

### Code.gs 中的引用
```javascript
// 點檢表單
template = HtmlService.createTemplateFromFile('InspectionForm');

// 統計報表
template = HtmlService.createTemplateFromFile('StatisticsView');
```

### 網址參數
- 點檢表單：`網址` 或 `網址?page=inspection`
- 統計報表：`網址?page=statistics`

## ✅ 檢查清單

部署時請確認：
- [ ] 已建立 8 個 .gs 檔案
- [ ] 已建立 2 個 HTML 檔案（InspectionForm 和 StatisticsView）
- [ ] 檔案名稱無重複
- [ ] Code.gs 中的引用正確

---

**記住：在 Apps Script 編輯器中建立檔案時，不需要輸入副檔名！**
