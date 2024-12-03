function transferDataMain() {
  try {
    // スプレッドシートAからデータをリスト化
    const rows = listColumnData();
    // スプレッドシートBにデータを転記
    transferToSpreadsheetDestination(rows);
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
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
    if (values[row][1] !== "" && values[row][1] !== null && values[row][1].startsWith('AS')) { // B列がASから始まる場合のみ追加
      const formattedRow = values[row].slice(1).map(formatValue); // A列を除外するためにslice(1)
      rows.push(formattedRow);
    }
  }

  // 行ごとのリスト化結果をログに出力
  Logger.log(rows);
  return rows;
}

function checkHeader(headerRow) {
  const expectedHeaders = [
    '', '資産管理番号', '', 'ステータス', '顧客名', '', '日付', '担当者', '備考', 'お預かり証No.'
  ];

  const columnLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

  for (let i = 0; i < expectedHeaders.length; i++) {
    if (expectedHeaders[i] && headerRow[i] !== expectedHeaders[i]) {
      throw new Error(`ヘッダーの${columnLetters[i]}列は「${expectedHeaders[i]}」である必要がありますが、実際の値は「${headerRow[i]}」です。期待される内容: 「${expectedHeaders.join('、')}」`);
    }
  }
}

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

  // スプレッドシート(SPREADSHEET_ID_DESTINATION)を取得
  const spreadsheetDestination = SpreadsheetApp.openById(spreadsheetDestinationId);
  const firstSheetDestination = spreadsheetDestination.getSheets()[0]; // デフォルトのシートを取得

  if (!firstSheetDestination) {
    throw new Error('スプレッドシート(SPREADSHEET_ID_DESTINATION)にデフォルトのシートが存在しません。');
  }

  const dataRangeB = firstSheetDestination.getRange(1, 2, firstSheetDestination.getLastRow(), 1); // B列の範囲を取得
  const valuesB = dataRangeB.getValues();

  for (const rowData of rows) {
    const assetNumber = rowData[0]; // リストの1番目が資産管理番号
    let targetRow = -1;

    for (let row = 0; row < valuesB.length; row++) {
      if (valuesB[row][0] === assetNumber) {
        targetRow = row + 1; // シート上の行番号は1から始まる
        break;
      }
    }

    if (targetRow !== -1) {
      // 日付をフォーマットする関数
      const formatDate = (date) => {
        if (Object.prototype.toString.call(date) === '[object Date]') {
          const year = date.getFullYear();
          const month = ('0' + (date.getMonth() + 1)).slice(-2);
          const day = ('0' + date.getDate()).slice(-2);
          return `${year}/${month}/${day}`;
        }
        return date; // 日付以外のデータはそのまま返す
      };
      
      // 各列にデータを転記
      firstSheetDestination.getRange(targetRow, 1).setValue(rowData[6] || ''); // A列にリスト6番目の内容を転記
      firstSheetDestination.getRange(targetRow, 11).setValue(rowData[2] || ''); // K列にリスト2番目の内容を転記
      firstSheetDestination.getRange(targetRow, 12).setValue(rowData[3] || ''); // L列にリスト3番目の内容を転記
      firstSheetDestination.getRange(targetRow, 13).setValue(rowData[5] || ''); // M列にリスト5番目の内容を転記
      firstSheetDestination.getRange(targetRow, 15).setValue(rowData[8] || ''); // O列にリスト8番目の内容を転記
      firstSheetDestination.getRange(targetRow, 16).setValue(rowData[7] || ''); // P列にリスト9番目の内容を転記

      // K列の内容が"代替貸出"の場合、N列に"有"を入力
      if (rowData[2] === "代替貸出") {
        firstSheetDestination.getRange(targetRow, 14).setValue("有"); // J列に"有"を転記
      }
    }
  }
}

function sendErrorNotification(error) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const recipient = scriptProperties.getProperty('ERROR_NOTIFICATION_EMAIL');

  if (!recipient) {
    throw new Error('エラーメール通知先のメールアドレスが設定されていません。スクリプトプロパティに「ERROR_NOTIFICATION_EMAIL」を正しく設定してください。');
  }

  const subject = 'スクリプトエラー通知';
  const body = 'エラーが発生しました' + error.message;
  MailApp.sendEmail(recipient, subject, body);
}
