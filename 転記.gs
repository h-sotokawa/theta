function transferDataMain() {
  try {
    // 転記元シートからデータをリスト化
    const rows = listColumnData();
    
    // rowsが空でないかをチェックし、空の場合は転記をスキップ
    if (rows.length === 0) {
      Logger.log('転記元データが存在しません。転記処理をスキップします。');
      sendLogNotification();
      return;
    }

    // rowsのすべての内容の資産管理番号をログに記録
    Logger.log('転記元データの資産管理番号一覧:');
    rows.forEach((row, index) => {
      Logger.log(`行 ${index + 1}: 資産管理番号: ${row[0]}`);
    });

    // 転記先シートにデータを転記
    const unprocessedRows = transferToSpreadsheetDestination(rows);

    // 転記元と転記先の差異をログに出力
    logDifferences(rows);

    // 処理した行数と実際の行数が一致しない場合、未処理の行をログに残す
    if (unprocessedRows.length > 0) {
      Logger.log(`未処理の行数: ${unprocessedRows.length}`);
      unprocessedRows.forEach((row, index) => {
        Logger.log(`未処理の行 ${index + 1}: 資産管理番号: ${row[0]}, 理由: 対応する転記先の行が見つかりませんでした。`);
      });
    } else {
      Logger.log('すべての行が正常に処理されました。');
    }

    // ログをメールで送信
    sendLogNotification();
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
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
  const valuesBMap = new Map(); // 使用する理由: Mapを使うことで検索時間がO(1)に短縮され、パフォーマンスが向上するため
  valuesB.forEach((value, index) => {
    if (value[0]) {
      valuesBMap.set(value[0], index + 4); // シート上の行番号は4から始まる
    }
  });

  let unprocessedRows = []; // 処理されなかった行を格納する配列
  let processedRowCount = 0;

  // 転記元の各行について処理
  for (const rowData of rows) {
    const assetNumber = rowData[0]; // 資産管理番号（B列の値）
    const targetRow = valuesBMap.get(assetNumber); // ハッシュマップから行番号を取得

    if (targetRow !== undefined) { // 修正: targetRowがundefinedでないことを確認
      // 各列にデータを転記
      destinationSheet.getRange(targetRow, 1).setValue(rowData[6] || ''); // A列にリスト6番目の内容を転記
      destinationSheet.getRange(targetRow, 11).setValue(rowData[2] || ''); // K列にリスト2番目の内容を転記

      if (rowData[2] === "代替貸出") {
        destinationSheet.getRange(targetRow, 12).setValue(rowData[3] || ''); // L列にリスト3番目の内容を転記
        destinationSheet.getRange(targetRow, 13).setValue(rowData[5] || ''); // M列にリスト5番目の内容を転記
        destinationSheet.getRange(targetRow, 15).setValue(rowData[8] || ''); // O列にリスト8番目の内容を転記
        destinationSheet.getRange(targetRow, 14).setValue("有"); // N列に"有"を転記
      } else {
        destinationSheet.getRange(targetRow, 12).clearContent(); // L列の内容を削除
        destinationSheet.getRange(targetRow, 13).clearContent(); // M列の内容を削除
        destinationSheet.getRange(targetRow, 15).clearContent(); // O列の内容を削除
      }
      processedRowCount++;
    } else {
      // 転記できなかった行を記録
      Logger.log(`未処理の資産管理番号: ${assetNumber}`); // 追加: 未処理の行をログに出力
      unprocessedRows.push(rowData);
    }
  }

  // 処理されなかった行を返す
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

function sendLogNotification() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const recipient = scriptProperties.getProperty('NOTIFICATION_EMAIL');
  const location = scriptProperties.getProperty('LOCATION_MAPPING');

  if (!recipient) {
    throw new Error('ログメール通知先のメールアドレスが設定されていません。スクリプトプロパティに「NOTIFICATION_EMAIL」を正しく設定してください。');
  }

  const log = Logger.getLog();
  const subject = '転記処理スクリプト実行結果通知：' + location;
  const body = 'スクリプトの実行結果を通知します:\n' + log;
  MailApp.sendEmail(recipient, subject, body);
}

function sendErrorNotification(error) {
  // スクリプトプロパティからエラーメールの通知先を取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const recipient = scriptProperties.getProperty('ERROR_NOTIFICATION_EMAIL');
  const location = scriptProperties.getProperty('LOCATION_MAPPING');

  if (!recipient) {
    throw new Error('エラーメール通知先のメールアドレスが設定されていません。スクリプトプロパティに「ERROR_NOTIFICATION_EMAIL」を正しく設定してください。');
  }

  // エラーメールを送信
  const subject = '転記処理スクリプトエラー通知：'+location;
  const body = 'エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack;
  MailApp.sendEmail(recipient, subject, body);
}
