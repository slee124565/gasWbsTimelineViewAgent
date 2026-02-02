依據 [text](test-guide.md) 測試步驟做測試發現以下幾個 Issues:
- `Resource` 欄位變化時， 修改的 Cell 對應的同一 Row 的 Background Color 都被修改，我需要調整成只針對 Resource 單一 Cell 的 Background Color 做修改，不要修改同一 Row 的其他 Cell
- `Object` 欄位變化時， 修改的 Cell 對應的同一 Row 的 Background Color 都被修改，我需要調整成只針對 Object 單一 Cell 的 Background Color 做修改，不要修改同一 Row 的其他 Cell

依據目錄下的兩個檔案內容做分析，找到 Issues 的問題原因，並且依據指令做後續的處理
1. [text](code.js)  
2. [text](gas-spec-v0.0.6.md)

步驟：
1. 找到產生 Issues 的原因
2. 設計解決方案，修正 code.js 設計並且將更新設計儲存在同一個 code.js 檔案內
3. 將 gas-spec-v0.0.6.md 設計規格的問題修正但是保留檔案原文，將新版的設計規格儲存在 gas-spec-v0.0.7.md 新檔案之中
