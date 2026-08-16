function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('個人財務分析儀表板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 提供給 index.html 內部呼叫讀取試算表資料的函數
function getSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("記帳資料") || ss.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  
  if (rows.length <= 1) return [];

  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] && !row[3]) continue;

    let dateStr = row[0];
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    let timeStr = row[1];
    if (row[1] instanceof Date) {
      timeStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "HH:mm");
    }

    data.push({
      id: i,
      date: String(dateStr || ""),
      time: String(timeStr || ""),
      account: String(row[2] || "其他"),
      name: String(row[3] || "未命名項目"),
      category: String(row[4] || "未分類"),
      amount: Number(row[5]) || 0,
      note: String(row[6] || "")
    });
  }
  return data;
}
