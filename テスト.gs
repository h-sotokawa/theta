function createTestData() {
  const sheetName = "テストデータシート";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const headers = [
    "タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり証No.", "預り機の写真", "備考"
  ];
  const statuses = ["受付中", "処理中", "完了"];

  // Set headers
  sheet.appendRow(headers);

  for (let i = 0; i < 100; i++) {
    const timestamp = new Date(Date.now() - Math.floor(Math.random() * 10000000000));
    const email = `test${i + 1}@example.com`;
    const status = statuses[Math.floor(Math.random() * statuses.length)];
    const customerName = `顧客${i + 1}`;
    const receiptNo = `R-${('0000' + (i + 1)).slice(-5)}`;
    const photoLink = `https://example.com/photo${i + 1}.jpg`;
    const remarks = `備考${i + 1}`;

    sheet.appendRow([timestamp, email, status, customerName, receiptNo, photoLink, remarks]);
  }
}