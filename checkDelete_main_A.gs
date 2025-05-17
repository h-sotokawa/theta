function checkDeleteMainA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('main');
  
  // D列のデータを取得
  const lastRow = sheet.getLastRow();
  const dColumnData = sheet.getRange(1, 4, lastRow, 1).getValues();
  
  // A列のデータを取得
  const aColumnData = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  // 削除対象の行を特定
  for (let i = 0; i < lastRow; i++) {
    const dValue = dColumnData[i][0];
    const aValue = aColumnData[i][0];
    
    // 条件チェック
    // 1. D列に"貸出"が含まれない
    // 2. A列が"預かり機のステータス"でない
    if (!dValue.toString().includes('貸出') && aValue !== '預かり機のステータス') {
      // A列のデータを削除（空文字に設定）
      sheet.getRange(i + 1, 1).setValue('');
    }
  }
} 