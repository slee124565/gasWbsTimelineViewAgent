
# Google Sheets 規格說明

本文件定義兩張工作表（Sheets）的**完整且穩定規格**，作為後續公式、自動化腳本（Apps Script）、或系統實作的唯一依據。

---

## Sheet 1：`wbs`
**用途定位**  
`wbs` 為任務層（Task-level）的主表，是整個 WBS 排程、責任歸屬、狀態追蹤與目標彙總的唯一事實來源。

---

### 一、欄位結構定義

#### （A）輸入欄位（人工或外部系統填寫）

| Column Name | 型別 | 必填 | 說明 |
|---|---|---|---|
| ObjectID | String | 是 | 目標唯一識別碼，例如 `O-01`、`O-02` |
| Object | String | 是 | 目標名稱（人類可讀） |
| TaskID | String | 是 | 任務識別碼，格式 `{ObjectID}-T-{NNN}`，例如 `O-01-T-010` |
| TaskTitle | String | 是 | 任務簡短標題 |
| StartDate | Date | 是 | 計畫起始日，保留原始輸入，不做校正 |
| WorkDays | Integer | 是 | 任務所需工作日數（正整數） |
| Resource | String | 否 | 執行人員，來自受控名單；可空 |
| TaskDescription-1 | String | 是 | 任務詳細描述（內容本體） |
| TaskStatus | Enum | 是 | 任務狀態（例如：NotStarted / InProgress / Done / Blocked） |
| DoneDate | Date | 否 | 任務實際完成日（允許假日或週末） |

---

#### （B）衍生欄位（系統自動計算）

| Column Name | 型別 | 說明 |
|---|---|---|
| DueDate | Date | 依 WorkDays 與工作日規則計算之預計完成日 |
| TaskDescription-2 | String | 任務描述（含人員），依組字規則產生 |

---

### 二、TaskID 編碼與排序規則

- **TaskID 格式**  
  `{ObjectID}-T-{NNN}`  
  例：`O-01-T-001`、`O-01-T-010`、`O-01-T-015`

- **編碼慣例**
  - 同一 ObjectID 底下，`NNN` 為三位數字
  - 預設以 **+010 跳號** 新增（001 → 010 → 020）
  - 允許插入（例如 015）
  - 同一 ObjectID 範圍內 TaskID 不可重複

- **排序規則**
  - 先依 ObjectID 分組
  - 再依 TaskID 中的 `NNN` 以「數值」排序（不可用字串排序）

---

### 三、DueDate 計算契約（正式定義）

定義一個計算用概念（不一定是實體欄位）：

- **FirstWorkingDate**  
  StartDate 當天或之後，第一個可工作的日期。

計算規則：

1. 若 StartDate 為工作日  
   → FirstWorkingDate = StartDate
2. 若 StartDate 為週末或 `holidays-tw` 中的假日  
   → FirstWorkingDate = StartDate 之後的第一個工作日
3. WorkDays 僅從 FirstWorkingDate 開始累加
4. 僅計算工作日，跳過：
   - 週末
   - `holidays-tw.date`

邊界條件：

- WorkDays = 1  
  → DueDate = FirstWorkingDate
- DoneDate 不參與 DueDate 計算

---

### 四、TaskStatus 與 DoneDate 一致性規則

- TaskStatus = Done  
  → DoneDate 必須有值
- DoneDate 有值  
  → TaskStatus 必須為 Done（或由系統自動修正）
- DoneDate 不得早於 StartDate（除非明確允許歷史回填）

---

### 五、Resource 與 TaskDescription-2 組字規則

- Resource 為受控名單欄位
- Resource 有值時：  
  `[Resource]-TaskDescription-1`
- Resource 為空時：  
  `[未指派]-TaskDescription-1`

---

## Sheet 2：`holidays-tw`
**用途定位**  
公司自訂行事曆，用於判斷工作日與非工作日，僅影響 DueDate 計算。

---

### 一、欄位結構定義

| Column Name | 型別 | 必填 | 說明 |
|---|---|---|---|
| date | Date | 是 | 假日日期（必須唯一） |
| holiday-name | String | 是 | 假日名稱 |

---

### 二、使用規則

- `holidays-tw.date` 中的所有日期，視為**非工作日**
- 不影響：
  - StartDate 合法性
  - DoneDate 合法性
- 僅用於：
  - FirstWorkingDate 判定
  - DueDate 工作日累加計算

---

## 補充說明（規格穩定性）

- `wbs` 與 `holidays-tw` 為後續所有公式、Apps Script、API、資料庫 schema 的基礎規格
- 未來若新增彙總表（例如 Object 層完成狀態），不得回寫或破壞 `wbs` 的事實資料

--- 

【END】
