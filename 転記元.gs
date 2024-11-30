function listColumnData() {
  // スクリプトプロパティからスプレッドシートのIDを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetAId = scriptProperties.getProperty('SPREADSHEET_A_ID');

  if (!spreadsheetAId) {
    throw new Error('スプレッドシートAのIDが設定されていません。スクリプトプロパティに「SPREADSHEET_A_ID」を設定してください。');
  }

  // スプレッドシートAの「main」シートを取得
  const spreadsheetA = SpreadsheetApp.openById(spreadsheetAId);
  const mainSheet = spreadsheetA.getSheetByName('main');

  if (!mainSheet) {
    throw new Error('スプレッドシートAに「main」シートが存在しません。');
  }

  // 「main」シートのデータを取得
  const dataRange = mainSheet.getDataRange();
  const values = dataRange.getValues();

  // A列にデータが入っている場合はエラーを返す
  for (let row = 0; row < values.length; row++) {
    if (values[row][0] !== "" && values[row][0] !== null) {
      throw new Error('A列にデータが入っています。A列は空である必要があります。');
    }
  }

  // B列の内容がASから始まる行だけリスト化し、A列を除外
  const rows = [];
  for (let row = 0; row < values.length; row++) {
    if (values[row][1] !== "" && values[row][1] !== null && values[row][1].startsWith('AS')) { // B列がASから始まる場合のみ追加
      const formattedRow = values[row].slice(1).map(value => { // A列を除外するためにslice(1)
        if (Object.prototype.toString.call(value) === '[object Date]') {
          const year = value.getFullYear();
          const month = ('0' + (value.getMonth() + 1)).slice(-2);
          const day = ('0' + value.getDate()).slice(-2);
          return `${year}/${month}/${day}`;
        }
        return value.toString();
      });
      rows.push(formattedRow);
    }
  }

  // 行ごとのリスト化結果をログに出力
  Logger.log(rows);
  return rows;
}

function transferToSpreadsheetB(rows) {
  // スクリプトプロパティからスプレッドシートのIDを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetBId = scriptProperties.getProperty('SPREADSHEET_B_ID');

  if (!spreadsheetBId) {
    throw new Error('スプレッドシートBのIDが設定されていません。スクリプトプロパティに「SPREADSHEET_B_ID」を設定してください。');
  }

  // スプレッドシートBを取得
  const spreadsheetB = SpreadsheetApp.openById(spreadsheetBId);
  const firstSheetB = spreadsheetB.getSheets()[0]; // デフォルトのシートを取得

  if (!firstSheetB) {
    throw new Error('スプレッドシートBにデフォルトのシートが存在しません。');
  }

  for (let i = 0; i < rows.length; i++) {
    const rowData = rows[i];
    const assetNumber = rowData[0]; // リストの1番目が資産管理番号

    // スプレッドシートB内のB列で資産管理番号を検索
    const dataRangeB = firstSheetB.getRange(1, 2, firstSheetB.getLastRow(), 1); // B列の範囲を取得
    const valuesB = dataRangeB.getValues();
    let targetRow = -1;

    for (let row = 0; row < valuesB.length; row++) {
      if (valuesB[row][0] === assetNumber) {
        targetRow = row + 1; // シート上の行番号は1から始まる
        break;
      }
    }

    if (targetRow !== -1) {
      // 各列にデータを転記
      firstSheetB.getRange(targetRow, 1).setValue(rowData[6] || ''); // A列にリスト6番目の内容を転記
      firstSheetB.getRange(targetRow, 7).setValue(rowData[2] || ''); // G列にリスト2番目の内容を転記
      firstSheetB.getRange(targetRow, 8).setValue(rowData[3] || ''); // H列にリスト3番目の内容を転記
      firstSheetB.getRange(targetRow, 9).setValue(rowData[5] || ''); // I列にリスト5番目の内容を転記
      firstSheetB.getRange(targetRow, 11).setValue(rowData[8] || ''); // K列にリスト8番目の内容を転記
      firstSheetB.getRange(targetRow, 12).setValue(rowData[7] || ''); // L列にリスト9番目の内容を転記

      // G列の内容が"代替貸出"の場合、J列に"有"を入力
      if (rowData[2] === "代替貸出") {
        firstSheetB.getRange(targetRow, 10).setValue("有"); // J列に"有"を転記
      }
    }
  }
}

function main() {
  try {
    // スプレッドシートAからデータをリスト化
    const rows = listColumnData();
    // スプレッドシートBにデータを転記
    transferToSpreadsheetB(rows);
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
  }
}
