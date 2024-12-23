function checkAndFixColumnsFormulas() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('main');
    if (!sheet) {
      Logger.log('シート「main」が見つかりません');
      return;
    }

    const logSheetName = 'check_formula_of_mainSheet_log';
    let logSheet = spreadsheet.getSheetByName(logSheetName);

    // ログシートが存在しない場合は作成
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(logSheetName);
      logSheet.appendRow(['タイムスタンプ', '列', '行', '現在の数式', '想定される数式', '結果']);
    }

    // 3ヶ月以上前のログを削除（ローテート）
    const threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
    const logData = logSheet.getDataRange().getValues();
    if (logData.length > 1) {
      const newLogData = logData.filter((row, index) => {
        if (index === 0) return true; // ヘッダーを維持
        const logDate = new Date(row[0]);
        return logDate >= threeMonthsAgo;
      });
      logSheet.clear();
      logSheet.getRange(1, 1, newLogData.length, newLogData[0].length).setValues(newLogData);
    }

    const lastRow = sheet.getLastRow(); // 最終行を取得

    // 検証対象の列（D列: 4列目 から J列: 10列目）
    const columns = [4, 5, 6, 7, 8, 9, 10];

    columns.forEach(column => {
      const baseFormula = sheet.getRange(4, column).getFormula().trim(); // 基準となる数式を取得

      if (!baseFormula) {
        Logger.log(`列 ${String.fromCharCode(64 + column)} の基準数式が存在しません。スキップします。`);
        logSheet.appendRow([
          new Date(),
          String.fromCharCode(64 + column),
          '',
          '',
          '',
          '基準数式が存在しません（スキップ）'
        ]);
        return;
      }

      // 現在の列の数式を取得
      const formulas = sheet.getRange(1, column, lastRow, 1).getFormulas();

      for (let i = 0; i < lastRow; i++) {
        const currentRow = i + 1;
        const currentFormula = (formulas[i][0] || '').toString().trim(); // 現在の数式をプレーンテキストとして取得しトリム
        const expectedFormula = baseFormula.replace(/4/g, currentRow).trim(); // 基準数式を現在の行に調整

        let result;
        if (!currentFormula) {
          // 数式が入っていない場合
          result = '数式が入力されていません';
          Logger.log(`列 ${String.fromCharCode(64 + column)} 行 ${currentRow}: ${result}`);
        } else if (currentFormula === expectedFormula) {
          // 数式が正しい場合
          result = '数式が正しいです。修正は不要です';
          Logger.log(`列 ${String.fromCharCode(64 + column)} 行 ${currentRow}: ${result}`);
        } else {
          // 数式が異なる場合
          result = '数式が異なります。修正しました';
          Logger.log(`列 ${String.fromCharCode(64 + column)} 行 ${currentRow}: ${result}`);
          sheet.getRange(currentRow, column).setFormula(expectedFormula); // 修正
        }

        // ログを記録
        logSheet.appendRow([
          new Date(),
          String.fromCharCode(64 + column),
          currentRow,
          `"${currentFormula || '（空白）'}"`,
          `"${expectedFormula}"`,
          result
        ]);
      }
    });

    Logger.log('対象列の数式を確認・修正しました。');
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
  }
}
