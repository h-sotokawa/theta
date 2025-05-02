function transferDataMain(sourceId, destinationId, location) {
  try {
    let allRows = [];
    let sourceIds = [];

    // 大阪の場合は複数のソースからデータを取得
    if (location === '大阪') {
      sourceIds = sourceId; // sourceIdは配列として渡される
      for (const id of sourceIds) {
        const rows = listColumnData(id);
        allRows = allRows.concat(rows);
        Logger.log(`ソースID ${id} から ${rows.length} 行のデータを取得しました。`);
        writeLogToSheet(id, `[開始] ソースID ${id} から ${rows.length} 行のデータを取得しました。`);
      }
    } else {
      // その他の拠点は単一のソースからデータを取得
      sourceIds = [sourceId];
      allRows = listColumnData(sourceId);
      Logger.log(`ソースID ${sourceId} から ${allRows.length} 行のデータを取得しました。`);
      writeLogToSheet(sourceId, `[開始] ソースID ${sourceId} から ${allRows.length} 行のデータを取得しました。`);
    }

    if (allRows.length === 0) {
      Logger.log('転記元データが存在しません。転記処理をスキップします。');
      sourceIds.forEach(id => {
        writeLogToSheet(id, '[終了] 転記元データが存在しません。転記処理をスキップします。');
      });
      return;
    }

    Logger.log('転記元データの資産管理番号一覧:');
    allRows.forEach((row, index) => {
      Logger.log(`行 ${index + 1}: 資産管理番号: ${row[0]}`);
    });
    sourceIds.forEach(id => {
      writeLogToSheet(id, `転記元データの資産管理番号一覧: ${allRows.map(row => row[0]).join(', ')}`);
    });

    const unprocessedRows = transferToSpreadsheetDestination(allRows, destinationId, location, sourceIds[0]);
    logDifferences(allRows, destinationId, location);

    if (unprocessedRows.length > 0) {
      Logger.log(`未処理の行数: ${unprocessedRows.length}`);
      unprocessedRows.forEach((row, index) => {
        Logger.log(`未処理の行 ${index + 1}: 資産管理番号: ${row[0]}, 理由: 対応する転記先の行が見つかりませんでした。`);
      });
      sourceIds.forEach(id => {
        writeLogToSheet(id, `[終了] 未処理の行数: ${unprocessedRows.length}`);
      });
    } else {
      Logger.log('すべての行が正常に処理されました。');
      sourceIds.forEach(id => {
        writeLogToSheet(id, '[終了] すべての行が正常に処理されました。');
      });
    }

    // 最後にログのローテーションを実行
    try {
      sourceIds.forEach(id => {
        const spreadsheet = SpreadsheetApp.openById(id);
        const logSheet = spreadsheet.getSheetByName('transferDataMain_log');
        if (logSheet) {
          rotateLog(logSheet);
        }
      });
    } catch (rotateError) {
      Logger.log('ログのローテーション中にエラーが発生しました: ' + rotateError.message);
    }

  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
    sourceIds.forEach(id => {
      writeLogToSheet(id, '[エラー] ' + error.message);
    });
    sendErrorNotification(error, location);
  }
}

function listColumnData(spreadsheetId) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sourceSheet = spreadsheet.getSheetByName('main');

  if (!sourceSheet) {
    throw new Error('「main」シートが存在しません。');
  }

  const dataRange = sourceSheet.getDataRange();
  const values = dataRange.getValues();

  checkHeader(values[0]);

  const rows = [];
  for (let row = 0; row < values.length; row++) {
    if (values[row][1] && values[row][1] !== "資産管理番号") {
      const formattedRow = values[row].slice(1).map(formatValue);
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

function transferToSpreadsheetDestination(rows, destinationId, location, sourceId) {
  try {
    // スクリプトプロパティからスプレッドシートのIDを取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const spreadsheetDestinationId = scriptProperties.getProperty('SPREADSHEET_ID_DESTINATION');

    if (!spreadsheetDestinationId) {
      throw new Error('スプレッドシート(SPREADSHEET_ID_DESTINATION)のIDが設定されていません。スクリプトプロパティに「SPREADSHEET_ID_DESTINATION」を正しく設定してください。');
    }

    // シート名を設定オブジェクトから取得
    const sheetName = location;
    if (!sheetName) {
      throw new Error('LOCATION_MAPPINGが設定されていません。設定オブジェクトに「location」を正しく設定してください。');
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

    // 転記先シートのB列をMapに変換
    const valuesBMap = new Map();
    valuesB.forEach((value, index) => {
      if (value[0]) {
        valuesBMap.set(value[0], index + 4);
      }
    });

    // 転記元の資産管理番号をMapに変換
    const sourceAssetNumbers = new Map();
    rows.forEach(row => {
      if (row[0]) {
        sourceAssetNumbers.set(row[0], true);
      }
    });

    let unprocessedRows = [];
    let processedRowCount = 0;
    let newRows = [];
    let deletedRows = [];

    // 転記元の各行について処理
    for (const rowData of rows) {
      const assetNumber = rowData[0];
      const targetRow = valuesBMap.get(assetNumber);

      if (targetRow !== undefined) {
        // 既存の行に対する転記処理（必要な列のみ更新）
        if (rowData[2] === "代替貸出") {
          // 代替貸出の場合
          destinationSheet.getRange(targetRow, 1).setValue(rowData[6] || '');  // A列：担当者
          destinationSheet.getRange(targetRow, 11).setValue(rowData[2] || ''); // K列：ステータス
          destinationSheet.getRange(targetRow, 12).setValue(rowData[3] || ''); // L列：貸出先
          destinationSheet.getRange(targetRow, 13).setValue(rowData[5] || ''); // M列：貸出日
          destinationSheet.getRange(targetRow, 15).setValue(rowData[8] || ''); // O列：お預かり証No.
          destinationSheet.getRange(targetRow, 16).setValue(rowData[7] || ''); // P列：備考
          destinationSheet.getRange(targetRow, 14).setValue(rowData[8] ? "有" : ""); // N列：ユーザー機有
        } else {
          // 代替貸出でない場合
          destinationSheet.getRange(targetRow, 1).setValue(rowData[6] || '');  // A列：担当者
          destinationSheet.getRange(targetRow, 11).setValue(rowData[2] || ''); // K列：ステータス
          destinationSheet.getRange(targetRow, 16).setValue(rowData[7] || ''); // P列：備考
          // 代替機関連の列をクリア
          destinationSheet.getRange(targetRow, 12).clearContent();  // L列：顧客名
          destinationSheet.getRange(targetRow, 13).clearContent();  // M列：貸出先
          destinationSheet.getRange(targetRow, 15).clearContent();  // O列：お預かり証No.
          destinationSheet.getRange(targetRow, 14).clearContent();  // N列：ユーザー機有
        }

        // 追加の列を転記
        destinationSheet.getRange(targetRow, 3).setValue(rowData[9] || '');  // C列：型番
        destinationSheet.getRange(targetRow, 4).setValue(rowData[10] || '');  // D列：シリアル
        destinationSheet.getRange(targetRow, 5).setValue(rowData[11] || '');  // E列：ソフト
        destinationSheet.getRange(targetRow, 6).setValue(rowData[12] || '');  // F列：OS

        processedRowCount++;
      } else {
        // 新しい代替機のデータを追加
        newRows.push(rowData);
      }
    }

    // 転記先に存在するが転記元に存在しない行を削除
    valuesB.forEach((value, index) => {
      const assetNumber = value[0];
      if (assetNumber && !sourceAssetNumbers.has(assetNumber)) {
        deletedRows.push({
          row: index + 4,
          assetNumber: assetNumber
        });
      }
    });

    // 削除する行を逆順にソート（下から削除することで行番号のずれを防ぐ）
    deletedRows.sort((a, b) => b.row - a.row);

    // 行を削除
    deletedRows.forEach(item => {
      destinationSheet.deleteRow(item.row);
      if (location === '大阪') {
        sourceId.forEach(id => {
          writeLogToSheet(id, `削除された行: 資産管理番号 ${item.assetNumber} (行 ${item.row})`);
        });
      } else {
        writeLogToSheet(sourceId, `削除された行: 資産管理番号 ${item.assetNumber} (行 ${item.row})`);
      }
    });

    // 新しい代替機のデータを追加
    if (newRows.length > 0) {
      const lastRow = destinationSheet.getLastRow();
      const startRow = lastRow + 1;
      
      newRows.forEach((rowData, index) => {
        const currentRow = startRow + index;
        
        // 新しい行にデータを転記
        destinationSheet.getRange(currentRow, 1).setValue(rowData[6] || '');
        destinationSheet.getRange(currentRow, 2).setValue(rowData[0]); // 資産管理番号
        destinationSheet.getRange(currentRow, 3).setValue(rowData[9] || ''); // C列：型番
        destinationSheet.getRange(currentRow, 4).setValue(rowData[10] || ''); // D列：シリアル
        destinationSheet.getRange(currentRow, 5).setValue(rowData[11] || ''); // E列：ソフト
        destinationSheet.getRange(currentRow, 6).setValue(rowData[12] || ''); // F列：OS
        destinationSheet.getRange(currentRow, 11).setValue(rowData[2] || '');

        if (rowData[2] === "代替貸出") {
          destinationSheet.getRange(currentRow, 12).setValue(rowData[3] || '');
          destinationSheet.getRange(currentRow, 13).setValue(rowData[5] || '');
          destinationSheet.getRange(currentRow, 15).setValue(rowData[8] || '');
          destinationSheet.getRange(currentRow, 16).setValue(rowData[7] || '');
          destinationSheet.getRange(currentRow, 14).setValue(rowData[8] ? "有" : "");
        }

        if (location === '大阪') {
          sourceId.forEach(id => {
            writeLogToSheet(id, `新しい代替機を追加しました: 資産管理番号 ${rowData[0]}`);
          });
        } else {
          writeLogToSheet(sourceId, `新しい代替機を追加しました: 資産管理番号 ${rowData[0]}`);
        }
      });
    }

    return unprocessedRows;
  } catch (error) {
    Logger.log('転記処理中にエラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
    throw error;
  }
}

function logDifferences(rows, destinationId, location) {
  // スプレッドシート(SPREADSHEET_ID_DESTINATION)を取得
  const spreadsheetDestination = SpreadsheetApp.openById(destinationId);
  const destinationSheet = spreadsheetDestination.getSheetByName(location);
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

/**
 * ログを記録する関数
 */
function writeLogToSheet(spreadsheetId, logMessage) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let logSheet = spreadsheet.getSheetByName('transferDataMain_log');
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('transferDataMain_log');
      logSheet.appendRow(['日時', 'ログメッセージ']);
    }

    // 固定フォーマットで日時を取得（YYYY/MM/DD HH:mm:ss）
    const timezone = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), timezone, 'yyyy/MM/dd HH:mm:ss');
    logSheet.appendRow([now, logMessage]);

    // ログのローテーションを実行
    rotateLog(logSheet);
  } catch (error) {
    Logger.log('ログ記録エラー: ' + error);
    throw error;
  }
}

/**
 * ログをローテート（1年より古い行を削除）
 */
function rotateLog(sheet) {
  try {
    const maxYears = 1;  // 1年間保持
    const cutoffDate = new Date();
    cutoffDate.setFullYear(cutoffDate.getFullYear() - maxYears);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();  // ヘッダー行

    // YYYY/MM/DD HH:mm:ss 形式で日付を比較
    const filteredData = data.filter(row => {
      const dateStr = row[0].toString();
      const [datePart, timePart] = dateStr.split(' ');
      const [year, month, day] = datePart.split('/').map(Number);
      const [hours, minutes, seconds] = timePart.split(':').map(Number);
      const logDate = new Date(year, month - 1, day, hours, minutes, seconds);
      return logDate >= cutoffDate;
    });

    // クリアしてヘッダー＋残存データを再書き込み
    sheet.clearContents();
    sheet.appendRow(headers);
    if (filteredData.length > 0) {
      sheet.getRange(2, 1, filteredData.length, headers.length)
           .setValues(filteredData);
    }
  } catch (error) {
    Logger.log('ローテートエラー: ' + error);
  }
}

function sendErrorNotification(error, location) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const recipient = scriptProperties.getProperty('ERROR_NOTIFICATION_EMAIL');

  if (!recipient) {
    throw new Error('エラーメール通知先のメールアドレスが設定されていません。スクリプトプロパティに「ERROR_NOTIFICATION_EMAIL」を正しく設定してください。');
  }

  const subject = '転記処理スクリプトエラー通知：'+location;
  const body = 'エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack;
  MailApp.sendEmail(recipient, subject, body);
}