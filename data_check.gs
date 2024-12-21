function main() {
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    Logger.log('スプレッドシートIDが設定されていません。スクリプトプロパティに「SPREADSHEET_ID」を設定してください。');
    return;
  }
  const sheets = getAllSheets(spreadsheetId);
  const correctColumnOrder = ["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり証No.", "備考"];
  const excludedSheetsPattern = /^(main|QRコード|.*_log)$/;
  const logSheet = getOrCreateLogSheet(spreadsheetId, 'data_check_log');

  sheets.forEach(sheet => {
    if (excludedSheetsPattern.test(sheet.getName())) {
      logMessage(logSheet, `シート「${sheet.getName()}」は処理の対象外です。`);
      return;
    }

    const headers = getHeaders(sheet);
    logMessage(logSheet, `シート「${sheet.getName()}」の現在のヘッダー: ${headers.join(', ')}`);

    if (!arraysEqual(headers, correctColumnOrder)) {
      logMessage(logSheet, `シート「${sheet.getName()}」の列の順番を修正します。`);
      rearrangeColumns(sheet, headers, correctColumnOrder);
      logMessage(logSheet, `シート「${sheet.getName()}」の列の順番を「${correctColumnOrder.join(', ')}」に整頓しました。`);
    } else {
      logMessage(logSheet, `シート「${sheet.getName()}」の列の順番は正しいです。`);
    }
  });
}

function getSpreadsheetId() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SPREADSHEET_ID');
}

function getAllSheets(spreadsheetId) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  return spreadsheet.getSheets();
}

function getHeaders(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return range.getValues()[0];
}

function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) {
    return false;
  }
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) {
      return false;
    }
  }
  return true;
}

function rearrangeColumns(sheet, currentHeaders, correctOrder) {
  const columnMap = {};
  currentHeaders.forEach((header, index) => {
    columnMap[header] = index + 1;
  });

  correctOrder.forEach((header, targetIndex) => {
    if (header in columnMap) {
      const currentIndex = columnMap[header] - 1;
      if (currentIndex !== targetIndex) {
        sheet.moveColumns(sheet.getRange(1, currentIndex + 1, sheet.getMaxRows()), targetIndex + 1);
      }
    }
  });
}

function setSpreadsheetId(spreadsheetId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SPREADSHEET_ID', spreadsheetId);
}

function getOrCreateLogSheet(spreadsheetId, logSheetName) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let logSheet = spreadsheet.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(logSheetName);
    logSheet.appendRow(["タイムスタンプ", "メッセージ"]);
  }
  return logSheet;
}

function logMessage(logSheet, message) {
  const timestamp = new Date();
  logSheet.appendRow([timestamp, message]);
}

// メイン関数を実行
main();
