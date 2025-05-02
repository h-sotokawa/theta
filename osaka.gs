function runOsaka() {
  try {
    // Logger.log('大阪の転記処理を開始します。');
    
    const config = Config.getLocationConfig('osaka');
    // Logger.log('大阪の設定を取得しました:');
    // Logger.log('ソースID: ' + config.sourceIds.join(', '));
    // Logger.log('転記先ID: ' + config.destinationId);
    // Logger.log('拠点: ' + config.location);

    // 各ソースIDが有効か確認
    config.sourceIds.forEach(id => {
      if (!id) {
        throw new Error('ソースIDが設定されていません。');
      }
      try {
        const spreadsheet = SpreadsheetApp.openById(id);
        const sheets = spreadsheet.getSheets();
        // Logger.log(`ソースID ${id} のスプレッドシートにアクセスできました。`);
        // Logger.log(`シート一覧: ${sheets.map(sheet => sheet.getName()).join(', ')}`);
      } catch (e) {
        throw new Error(`ソースID ${id} のスプレッドシートにアクセスできません: ${e.message}`);
      }
    });

    // 転記先IDが有効か確認
    try {
      const destinationSpreadsheet = SpreadsheetApp.openById(config.destinationId);
      const sheets = destinationSpreadsheet.getSheets();
      // Logger.log(`転記先ID ${config.destinationId} のスプレッドシートにアクセスできました。`);
      // Logger.log(`シート一覧: ${sheets.map(sheet => sheet.getName()).join(', ')}`);
    } catch (e) {
      throw new Error(`転記先ID ${config.destinationId} のスプレッドシートにアクセスできません: ${e.message}`);
    }

    // Logger.log('転記処理を開始します。');
    transferDataMain(config.sourceIds, config.destinationId, config.location);
    // Logger.log('転記処理が完了しました。');
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message + '\nスタックトレース: ' + error.stack);
  }
}
