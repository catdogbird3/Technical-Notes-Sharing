## Gmail 支出記錄自動化技術筆記

### 一、專案概述

* **目的**：透過 Google Apps Script，自動從 Gmail 中擷取收據郵件，並將交易日期、廠商、金額等資訊，記錄到 Google 試算表中，減少手動輸入花費的麻煩。
* **功能**：

  1. 自動標記符合條件的收據郵件
  2. 定時或手動擷取近 14 天的收據記錄
  3. 將結果寫入試算表，並跳出提醒已新增筆數

### 二、準備工作

1. **Google 試算表**：建立或確認有一份名稱為 `Expenses` 的試算表，第一行依序輸入：

   ```
   Date | Vendor | Amount | Thread URL | Thread ID
   ```
2. **Gmail 標籤**：在 Gmail 中建立一個名為 `receipts` 的標籤，用於套用在收據郵件。也可交由程式自動建立。
3. **Apps Script 專案**：

   * 在試算表上方選單：`延伸功能` → `Apps Script`，開啟程式碼編輯器。
   * 刪除預設程式，貼上下方完整程式碼。

### 三、程式碼結構與說明

```javascript
// 1. onOpen：試算表開啟時，新增「Expense Logger」選單
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Expense Logger')
    .addItem('Fetch Now', 'fetchExpenses')          // 手動擷取收據
    .addItem('Setup Daily Trigger', 'createDailyTrigger')  // 設定定時觸發
    .addItem('Auto-Label Receipts', 'autoLabelReceipts')   // 自動標籤收據
    .addToUi();
}

// 2. onInstall：專案首次安裝時，呼叫 onOpen 並建立定時觸發器
function onInstall(e) {
  onOpen(e);
  createDailyTrigger();
}

// 3. fetchExpenses：抓取近 14 天標記為 receipts 的郵件，提取並寫入試算表
function fetchExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Expenses');
  if (!sheet) {
    sheet = ss.insertSheet('Expenses');
    sheet.appendRow(['Date','Vendor','Amount','Thread URL','Thread ID']);
  }
  const query = 'label:receipts newer_than:14d';
  const threads = GmailApp.search(query);
  let count = 0;
  threads.forEach(thread => {
    const id = thread.getId();
    if (sheet.createTextFinder(id).matchEntireCell(true).findNext()) return;
    thread.getMessages().forEach(msg => {
      const date   = msg.getDate();
      const body   = msg.getPlainBody();
      const match  = body.match(/(?:NT\$|\$)\s?([0-9,]+\.?[0-9]{0,2})/i);
      const amount = match ? match[1].replace(/,/g,'') : '';
      const vendor = msg.getFrom().replace(/.*<|>.*/g,'');
      const url    = 'https://mail.google.com/mail/u/0/#inbox/'+id;
      sheet.appendRow([date, vendor, amount, url, id]);
      count++;
    });
  });
  SpreadsheetApp.getUi().alert(`Fetched ${count} new records.`);
}

// 4. autoLabelReceipts：自動標記最近 14 天標題含「receipt」或「收據」的郵件
function autoLabelReceipts() {
  let label = GmailApp.getUserLabelByName('receipts');
  if (!label) label = GmailApp.createLabel('receipts');
  const query = 'subject:(receipt OR 收據) newer_than:14d';
  const threads = GmailApp.search(query);
  threads.forEach(thread=>thread.addLabel(label));
  SpreadsheetApp.getUi().alert(`Applied label to ${threads.length} threads.`);
}

// 5. createDailyTrigger：建立每天 08:00 自動執行 fetchExpenses 的觸發器
function createDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==='fetchExpenses')
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('fetchExpenses')
    .timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Daily trigger set for 08:00.');
}
```

### 四、操作流程

1. **貼上程式碼** → 按下儲存
2. **重新整理試算表** → 右上角可看到「Expense Logger」選單
![image](https://hackmd.io/_uploads/ryRPv2Jrxe.png)

4. **首次執行**：

   * 點選「Auto-Label Receipts」以自動套標
   * 點選「Fetch Now」檢查是否有資料匯入
4. **自動化**

   * 點選「Setup Daily Trigger」或由 `onInstall` 自動設定，確保每天上午 8 點自動更新記錄
![image](https://hackmd.io/_uploads/BkmsDhyrxx.png)

### 五、常見問題

* **為何抓不到任何郵件？**

  1. 確認 Gmail 標籤「receipts」已有套用於收據郵件
  2. 可手動測試 `autoLabelReceipts()` 產生標籤後，再執行 `fetchExpenses()`
* **要抓不同來源的收據？**
  修改 `autoLabelReceipts()` 中的 `query`，例如 `from:amazon.com`、`subject:Invoice` 等。

---

> 完成以上設定後，只要有符合條件的收據郵件，就能自動記錄到試算表，讓你輕鬆管理花費！
