依據 [text](test-guide.md) 測試步驟做測試發現以下幾個 Issues:
- `Resource` 與 `Object` 欄位變更之後的背景顏色設定機制保留系統偵測與自動執行，但是不需要用 dialog 通知用戶

依據目錄下的兩個檔案內容做分析，找到 Issues 的問題原因，並且依據指令做後續的處理
1. [text](code.js)  
2. [text](gas-spec-v0.0.7.md)

步驟：
1. 找到產生 Issues 的原因
2. 設計解決方案，修正 code.js 設計並且將更新設計儲存在同一個 code.js 檔案內
3. 將 gas-spec-v0.0.7.md 設計規格的問題修正但是保留檔案原文，將新版的設計規格儲存在 gas-spec-v0.0.8.md 新檔案之中
