# Google Sheets 規格說明 (簡化版)

本文件定義 `wbs` 與 `holidays-tw` 之結構，作為自動化系統實作之唯一依據。

---

## Sheet 1：`wbs`
**用途定位** 任務層級主表。本版本移除 ObjectID/TaskID，改以 **[Object + TaskTitle]** 作為識別唯一任務的複合主鍵。

### 一、欄位結構定義

#### （A）輸入欄位（人工或外部系統填寫）

| Column Name | 型別 | 必填 | 說明 |
|---|---|---|---|
| Object | String | 是 | 目標名稱。與 TaskTitle 組合後須具備唯一性 |
| TaskTitle | String | 是 | 任務標題。同一個 Object 下不可重複 |
| TaskDescription-1 | String | 是 | 任務詳細描述 |
| StartDate | Date | 否 | 計畫起始日，預設為今天 |
| WorkDays | Integer | 否 | 任務所需工作日數（正整數），預設為 5 |
| Resource | String | 否 | 執行人員 |
| TaskStatus | Enum | 否 | 狀態（NotStarted / InProgress / Done / Blocked），預設為 NotStarted |
| DoneDate | Date | 否 | 任務實際完成日 |

#### （B）衍生欄位（系統自動計算）

| Column Name | 型別 | 說明 |
|---|---|---|
| DueDate | Date | 依 WorkDays 與工作日規則計算之預計完成日 |
| TaskDescription-2 | String | 組字規則：`[人員]-描述` |

---

### 二、識別與唯一性規則

1. **複合主鍵 (Primary Key)**：系統判定唯一任務的邏輯由 `Object` + `TaskTitle` 組成。
2. **防重機制**：自動化腳本應在資料寫入前檢查同一 `Object` 下是否已存在相同的 `TaskTitle`。

---

### 三、DueDate 計算規則

1. **起始點判定**：
   - 若 `StartDate` 為工作日，則為 `FirstWorkingDate`。
   - 若 `StartDate` 為週末或 `holidays-tw` 假日，則 `FirstWorkingDate` 為其後第一個工作日。
2. **工作日累加**：
   - 從 `FirstWorkingDate` 開始計算，僅計入工作日。
   - 排除週末與 `holidays-tw` 表列日期。
   - 若 `WorkDays = 1`，則 `DueDate` 即為 `FirstWorkingDate`。

---

### 四、一致性與組字邏輯

- **Done 狀態連動**：當 `TaskStatus` 為 "Done" 時，`DoneDate` 必須有日期值；反之亦然。
- **描述組字 (TaskDescription-2)**：
  - 若 `Resource` 有值：`[Resource]-TaskDescription-1`。
  - 若 `Resource` 為空：`[未指派]-TaskDescription-1`。

---

## Sheet 2：`holidays-tw`
**用途定位** 定義非工作日清單，僅影響 `DueDate` 計算。

| Column Name | 型別 | 必填 | 說明 |
|---|---|---|---|
| date | Date | 是 | 假日日期（唯一值） |
| holiday-name | String | 是 | 假日名稱 |

---
【END】