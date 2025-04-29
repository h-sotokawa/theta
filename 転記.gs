function transferDataMain() {
  try {
    const rows = listColumnData();

    if (rows.length === 0) {
      Logger.log('転記元データが存在しません。転記処理をスキップします。');
      writeLogToSheet('転記元データが存在しません。転記処理をスキップします。');
      return;
    }

    Logger.log('転記元データの資産管理番号一覧:');
    rows.forEach((row, index) => {
      Logger.log(`行 ${index + 1}: 資産管理番号: ${row[0]}`);
    });
    writeLogToSheet(`転記元データの資産管理番号一覧: ${rows.map(row => row[0]).join(', ')}`);

    const unprocessedRows = transferToSpreadsheetDestination(rows);
    logDifferences(rows);

    if (unprocessedRows.length > 0) {
      Logger.log(`未処理の行数: ${unprocessedRows.length}`);
      unprocessedRows.forEach((row, index) => {
        Logger.log(`未処理の行 ${index + 1}: 資産管理番号: ${row[0]}, 理由: 対応する転記先の行が見つかりませんでした。`);
      });
      writeLogToSheet(`未処理の行数: ${unprocessedRows.length}`);
    } else {
      Logger.log('すべての行が正常に処理されました。');
      writeLogToSheet('すべての行が正常に処理されました。');
    }
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
    writeLogToSheet('エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
    sendErrorNotification(error);
  }
}

function listColumnData() {
  // スクリプトプロパティからスプレッドシートのIDを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetSourceId = scriptProperties.getProperty('SPREADSHEET_ID_SOURCE');

  if (!spreadsheetSourceId) {
    throw new Error('スプレッドシート(SPREADSHEET_ID_SOURCE)のIDが設定されていません。スクリプトプロパティに「SPREADSHEET_ID_SOURCE」を正しく設定してください。');
  }

  // スプレッドシート(SPREADSHEET_ID_SOURCE)の「main」シートを取得
  const spreadsheetSource = SpreadsheetApp.openById(spreadsheetSourceId);
  const sourceSheet = spreadsheetSource.getSheetByName('main');

  if (!sourceSheet) {
    throw new Error('スプレッドシート(SPREADSHEET_ID_SOURCE)に「main」シートが存在しません。');
  }

  // 「main」シートのデータを取得
  const dataRange = sourceSheet.getDataRange();
  const values = dataRange.getValues();

  // ヘッダーの内容をチェック
  checkHeader(values[0]);

  // B列の内容がASから始まる行だけリスト化し、A列を除外
  const rows = [];
  for (let row = 0; row < values.length; row++) {
    if (values[row][1] && values[row][1].startsWith('AS')) { // B列がASから始まる場合のみ追加
      const formattedRow = values[row].slice(1).map(formatValue); // A列を除外するためにslice(1)
      rows.push(formattedRow);
    }
  }

  return rows;
}

function checkHeader(headerRow) {
  // 期待されるヘッダー内容を定義
  const expectedHeaders = [
    '', '資産管理番号', '', 'ステータス', '顧客名', '', '日付', '担当者', '備考', 'お預かり証No.'
  ];

  const columnLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

  // ヘッダーが期待される内容かどうかをチェック
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (expectedHeaders[i] && headerRow[i] !== expectedHeaders[i]) {
      throw new Error(`ヘッダーの${columnLetters[i]}列は「${expectedHeaders[i]}」である必要がありますが、実際の値は「${headerRow[i]}」です。期待される内容: 「${expectedHeaders.join('、')}」`);
    }
  }
}

// 日時を整形して文字列に変換
function formatValue(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    const year = value.getFullYear();
    const month = ('0' + (value.getMonth() + 1)).slice(-2);
    const day = ('0' + value.getDate()).slice(-2);
    return `${year}/${month}/${day}`;
  }
  return value.toString();
}

function transferToSpreadsheetDestination(rows) {
  // スクリプトプロパティからスプレッドシートのIDを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetDestinationId = scriptProperties.getProperty('SPREADSHEET_ID_DESTINATION');

  if (!spreadsheetDestinationId) {
    throw new Error('スプレッドシート(SPREADSHEET_ID_DESTINATION)のIDが設定されていません。スクリプトプロパティに「SPREADSHEET_ID_DESTINATION」を正しく設定してください。');
  }

  // シート名をスクリプトプロパティから取得(スクリプトプロパティには拠点を入力する想定)
  const sheetName = scriptProperties.getProperty('LOCATION_MAPPING');
  if (!sheetName) {
    throw new Error('LOCATION_MAPPINGが設定されていません。スクリプトプロパティに「LOCATION_MAPPING」を正しく設定してください。');
  }

  // スプレッドシート(SPREADSHEET_ID_DESTINATION)を取得
  const spreadsheetDestination = SpreadsheetApp.openById(spreadsheetDestinationId);
  const allSheets = spreadsheetDestination.getSheets().map(sheet => sheet.getName());

  // 転記先のシートが存在するか確認
  if (!allSheets.includes(sheetName)) {
    throw new Error(`スプレッドシート(SPREADSHEET_ID_DESTINATION)に「${sheetName}」シートが存在しません。存在するシート名: ${allSheets.join(', ')}`);
  }

  const destinationSheet = spreadsheetDestination.getSheetByName(sheetName);

  // B列の4行目以降のデータ範囲を取得
  const dataRangeB = destinationSheet.getRange(4, 2, destinationSheet.getLastRow() - 3, 1);
  const valuesB = dataRangeB.getValues();

  // valuesBの内容をログに出力
  Logger.log('転記先シートのB列の内容:');
  valuesB.forEach((value, index) => {
    Logger.log(`行 ${index + 4}: ${value[0]}`);
  });

  // 転記先シートのB列をハッシュマップに変換して、高速な検索を可能にする
  const valuesBMap = new Map();
  valuesB.forEach((value, index) => {
    if (value[0]) {
      valuesBMap.set(value[0], index + 4);
    }
  });

  let unprocessedRows = [];
  let processedRowCount = 0;
  let newRows = [];

  // 転記元の各行について処理
  for (const rowData of rows) {
    const assetNumber = rowData[0];
    const targetRow = valuesBMap.get(assetNumber);

    if (targetRow !== undefined) {
      // 既存の行に対する転記処理
      destinationSheet.getRange(targetRow, 1).setValue(rowData[6] || '');
      destinationSheet.getRange(targetRow, 11).setValue(rowData[2] || '');

      if (rowData[2] === "代替貸出") {
        destinationSheet.getRange(targetRow, 12).setValue(rowData[3] || '');
        destinationSheet.getRange(targetRow, 13).setValue(rowData[5] || '');
        destinationSheet.getRange(targetRow, 15).setValue(rowData[8] || '');
        destinationSheet.getRange(targetRow, 16).setValue(rowData[7] || '');
        destinationSheet.getRange(targetRow, 14).setValue("有");
      } else {
        destinationSheet.getRange(targetRow, 12).clearContent();
        destinationSheet.getRange(targetRow, 13).clearContent();
        destinationSheet.getRange(targetRow, 15).clearContent();
        destinationSheet.getRange(targetRow, 14).clearContent();
      }
      processedRowCount++;
    } else {
      // 新しい代替機のデータを追加
      newRows.push(rowData);
    }
  }

  // 新しい代替機のデータを追加
  if (newRows.length > 0) {
    const lastRow = destinationSheet.getLastRow();
    const startRow = lastRow + 1;
    
    newRows.forEach((rowData, index) => {
      const currentRow = startRow + index;
      
      // 新しい行にデータを転記
      destinationSheet.getRange(currentRow, 1).setValue(rowData[6] || '');
      destinationSheet.getRange(currentRow, 2).setValue(rowData[0]); // 資産管理番号
      destinationSheet.getRange(currentRow, 11).setValue(rowData[2] || '');

      if (rowData[2] === "代替貸出") {
        destinationSheet.getRange(currentRow, 12).setValue(rowData[3] || '');
        destinationSheet.getRange(currentRow, 13).setValue(rowData[5] || '');
        destinationSheet.getRange(currentRow, 15).setValue(rowData[8] || '');
        destinationSheet.getRange(currentRow, 16).setValue(rowData[7] || '');
        destinationSheet.getRange(currentRow, 14).setValue("有");
      }

      // 新しい行の追加をログに記録
      Logger.log(`新しい代替機を追加しました: 資産管理番号 ${rowData[0]}`);
      writeLogToSheet(`新しい代替機を追加しました: 資産管理番号 ${rowData[0]}`);
    });
  }

  return unprocessedRows;
}

function logDifferences(rows) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetDestinationId = scriptProperties.getProperty('SPREADSHEET_ID_DESTINATION');
  const sheetName = scriptProperties.getProperty('LOCATION_MAPPING');

  if (!spreadsheetDestinationId || !sheetName) {
    throw new Error('スプレッドシートの情報が設定されていません。');
  }

  const spreadsheetDestination = SpreadsheetApp.openById(spreadsheetDestinationId);
  const destinationSheet = spreadsheetDestination.getSheetByName(sheetName);
  const dataRangeB = destinationSheet.getRange(4, 2, destinationSheet.getLastRow() - 3, 1);
  const valuesB = dataRangeB.getValues().map(row => row[0]);

  Logger.log('転記元データに含まれるが転記先シートに存在しない資産管理番号:');
  rows.forEach(row => {
    if (!valuesB.includes(row[0])) {
      Logger.log(`資産管理番号: ${row[0]}`);
    }
  });

  Logger.log('転記先シートに含まれるが転記元データに存在しない資産管理番号:');
  valuesB.forEach(value => {
    if (!rows.some(row => row[0] === value)) {
      Logger.log(`資産管理番号: ${value}`);
    }
  });
}

function writeLogToSheet(logMessage) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetSourceId = scriptProperties.getProperty('SPREADSHEET_ID_SOURCE');

  if (!spreadsheetSourceId) {
    throw new Error('スプレッドシート(SPREADSHEET_ID_SOURCE)のIDが設定されていません。');
  }

  const spreadsheetSource = SpreadsheetApp.openById(spreadsheetSourceId);
  let logSheet = spreadsheetSource.getSheetByName('transferDataMain_log');

  if (!logSheet) {
    logSheet = spreadsheetSource.insertSheet('transferDataMain_log');
    logSheet.appendRow(['日時', 'ログメッセージ']);
  }

  const now = new Date();
  logSheet.appendRow([now.toLocaleString(), logMessage]);

  rotateLog(logSheet);
}

function rotateLog(sheet) {
  const maxYears = 3;
  const today = new Date();
  const cutoffDate = new Date(today.getFullYear() - maxYears, today.getMonth(), today.getDate());

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const filteredData = data.filter(row => {
    const logDate = new Date(row[0]);
    return logDate >= cutoffDate;
  });

  sheet.clear();
  sheet.appendRow(headers);
  if (filteredData.length > 0) {
    sheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
}

function sendErrorNotification(error) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const recipient = scriptProperties.getProperty('ERROR_NOTIFICATION_EMAIL');
  const location = scriptProperties.getProperty('LOCATION_MAPPING');

  if (!recipient) {
    throw new Error('エラーメール通知先のメールアドレスが設定されていません。スクリプトプロパティに「ERROR_NOTIFICATION_EMAIL」を正しく設定してください。');
  }

  const subject = '転記処理スクリプトエラー通知：'+location;
  const body = 'エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack;
  MailApp.sendEmail(recipient, subject, body);
}