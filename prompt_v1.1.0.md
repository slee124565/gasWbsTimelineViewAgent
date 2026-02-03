由於 Google Sheets 「時間軸」功能無法透過 gas 程式自動設定，
所以修改使用者建構 wbs 流程是從複製 google sheets 範例檔案開始，
透過範例檔案已存在的 `wbs` & `TW_HOLIDYAS` & `Gantt-Chart` & `Resource-Chart` 提供使用者設定的便利性，
讓使用者從開放權限的範例檔案[wbs-to-gantt-and-resource-chart-view](https://docs.google.com/spreadsheets/d/1o6I1fYqk0xD9SdPk_qnZDObbcgyKA6gajaSGauc-VlM/edit?usp=sharing) 開始，
針對舊版 [text](gas-spec-v1.0.1.md) 規格做調整，
設計使用者的複製範例與修改流程，
並且依據你的設計流程調整以下檔案內容
1. 設計新版 goolge sheets 規格檔案 [text](gas-spec-v1.1.0.md)
2. 依據新版規格 [text](gas-spec-v1.1.0.md) 調整 [text](code.js)
3. 依據新版規格 [text](gas-spec-v1.1.0.md) 調整 [text](README.md)
