function listAllSheetNames() {
  try {
    // スプレッドシートの取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シートの取得とログ出力
    var sheets = ss.getSheets();
    Logger.log("=== シート一覧の確認 ===");
    
    if (sheets.length === 0) {
      Logger.log("エラー: スプレッドシートにシートがありません。");
      return;
    }

    sheets.forEach(function(sheet, index) {
      Logger.log("シート %d: %s", index + 1, sheet.getName());
    });

    Logger.log("総シート数: %d", sheets.length);
  } catch (e) {
    Logger.log("エラー発生: %s", e.toString());
  }
}
