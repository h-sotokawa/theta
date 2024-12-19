function checkAndFixColumnOrder_main() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    throw new Error("スクリプトプロパティに 'SPREADSHEET_ID' が設定されていません。");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getActiveSheet();
  const logSheet = spreadsheet.getSheetByName("header_check_on_main_sheet_log") || createLogSheet(spreadsheet);
  const correctColumnOrder = [
    "機種名", "資産管理番号", "拠点管理番号", "ステータス", 
    "顧客名", "ユーザー機シリアル", "日付", 
    "担当者", "備考", "お預かり証No."
  ];

  // ログシートのローテーション
  rotateLog(logSheet);

  // 現在のヘッダーを取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!arraysEqual(headers, correctColumnOrder)) {
    logMessage(logSheet, `シート「${sheet.getName()}」の列の順番を修正します。`);
    logColumnChanges(logSheet, headers, correctColumnOrder);
    rearrangeColumns(sheet, headers, correctColumnOrder);
    logMessage(logSheet, `シート「${sheet.getName()}」の列の順番を「${correctColumnOrder.join(', ')}」に整頓しました。`);
  } else {
    logMessage(logSheet, `シート「${sheet.getName()}」の列の順番は正しいです。`);
  }
}

/**
 * 配列が等しいかをチェック
 * @param {Array} array1 - 配列1
 * @param {Array} array2 - 配列2
 * @return {boolean} 等しければ true
 */
function arraysEqual(array1, array2) {
  if (array1.length !== array2.length) return false;
  for (let i = 0; i < array1.length; i++) {
    if (array1[i] !== array2[i]) return false;
  }
  return true;
}

/**
 * 列を正しい順序に並び替える
 * @param {Sheet} sheet - 操作対象のシート
 * @param {Array} headers - 現在のヘッダー
 * @param {Array} correctOrder - 正しい列の順序
 */
function rearrangeColumns(sheet, headers, correctOrder) {
  const columnIndexMap = headers.reduce((map, header, index) => {
    map[header] = index + 1; // 1-based index
    return map;
  }, {});

  const rearrangedData = [correctOrder]; // ヘッダー行を追加
  const data = sheet.getDataRange().getValues().slice(1); // ヘッダー以外のデータを取得

  for (const row of data) {
    const newRow = correctOrder.map(header => row[columnIndexMap[header] - 1] || "");
    rearrangedData.push(newRow);
  }

  // シートをクリアして並び替え後のデータを設定
  sheet.clear();
  sheet.getRange(1, 1, rearrangedData.length, rearrangedData[0].length).setValues(rearrangedData);
}

/**
 * ログメッセージを記録
 * @param {Sheet} logSheet - ログ用のシート
 * @param {string} message - 記録するメッセージ
 */
function logMessage(logSheet, message) {
  const timestamp = new Date();
  logSheet.appendRow([timestamp, message]);
}

/**
 * 列の変更内容を記録
 * @param {Sheet} logSheet - ログ用のシート
 * @param {Array} headers - 現在のヘッダー
 * @param {Array} correctOrder - 正しい列の順序
 */
function logColumnChanges(logSheet, headers, correctOrder) {
  const changes = headers.map((header, index) => {
    const newIndex = correctOrder.indexOf(header);
    return `${header}: ${index + 1} -> ${newIndex + 1}`;
  });
  logMessage(logSheet, `列の変更内容: ${changes.join(', ')}`);
}

/**
 * ログ用シートを作成
 * @param {Spreadsheet} spreadsheet - 対象のスプレッドシート
 * @return {Sheet} 作成されたログシート
 */
function createLogSheet(spreadsheet) {
  const logSheet = spreadsheet.insertSheet("header_check_on_main_sheet_log");
  logSheet.appendRow(["タイムスタンプ", "メッセージ"]);
  return logSheet;
}

/**
 * ログシートをローテーション
 * @param {Sheet} logSheet - ログ用のシート
 */
function rotateLog(logSheet) {
  const oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

  const data = logSheet.getDataRange().getValues();
  const header = data[0]; // ヘッダー行
  const filteredData = data.filter((row, index) => {
    if (index === 0) return true; // ヘッダーは保持
    const logDate = new Date(row[0]);
    return logDate >= oneYearAgo;
  });

  // シートをクリアして保持するデータを再設定
  logSheet.clear();
  logSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
}
