/* 定数定義 */
const CONSTANTS = {
  MAIN_SHEET_NAME: 'main',
  LOG_SHEET_NAME: 'check_formula_of_mainSheet_log',
  LOG_HEADERS: ['タイムスタンプ', '列', '行', '基準数式', '現在の数式', '想定される数式', '結果'],
  TARGET_COLUMNS: [4, 5, 6, 7, 8, 9, 10],
  LOG_RETENTION_DAYS: 14,    // ログ保持期間（日数）を2週間（14日）に設定
  START_ROW: 4
};

function checkAndFixColumnsFormulas() {
  Logger.log('===== 処理開始: 対象列の数式チェック・修正処理 =====');
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONSTANTS.MAIN_SHEET_NAME);
    if (!sheet) {
      Logger.log(`シート「${CONSTANTS.MAIN_SHEET_NAME}」が見つかりません`);
      return;
    }

    const logSheet = initializeLogSheet(spreadsheet);
    rotateLog(logSheet);

    // スプレッドシートのログにも開始を記録
    const startTime = new Date();
    logSheet.appendRow([startTime, '', '', '', '', '', '処理開始']);

    const lastRow = sheet.getLastRow();
    const batchUpdates = [];
    const logEntries = [];

    for (const column of CONSTANTS.TARGET_COLUMNS) {
      const columnResult = processColumn(sheet, column, lastRow);
      if (columnResult) {
        batchUpdates.push(...columnResult.updates);
        logEntries.push(...columnResult.logs);
      }
    }

    // 処理終了ログを追加
    logEntries.push([new Date(), '', '', '', '', '', '処理終了']);

    // バッチ処理で更新を適用
    if (batchUpdates.length > 0) {
      applyBatchUpdates(sheet, batchUpdates);
    }

    // ログを一括で追加 (開始ログは既に記録済)
    if (logEntries.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length)
        .setValues(logEntries);
    }

    Logger.log('対象列の数式を確認・修正しました。');
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
  } finally {
    Logger.log('===== 処理終了: 対象列の数式チェック・修正処理 =====');
  }
}

function initializeLogSheet(spreadsheet) {
  let logSheet = spreadsheet.getSheetByName(CONSTANTS.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(CONSTANTS.LOG_SHEET_NAME);
    logSheet.appendRow(CONSTANTS.LOG_HEADERS);
  }
  return logSheet;
}

function rotateLog(logSheet) {
  const cutoffDate = new Date();
  // ログ保持期間を2週間（14日）に設定
  cutoffDate.setDate(cutoffDate.getDate() - CONSTANTS.LOG_RETENTION_DAYS);
  
  const logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) return;

  const newLogData = logData.filter((row, index) => {
    if (index === 0) return true; // ヘッダーを維持
    const logDate = new Date(row[0]);
    return logDate >= cutoffDate;
  });

  logSheet.clear();
  if (newLogData.length > 0) {
    logSheet.getRange(1, 1, newLogData.length, newLogData[0].length).setValues(newLogData);
  } else {
    logSheet.appendRow(CONSTANTS.LOG_HEADERS);
  }
}

function processColumn(sheet, column, lastRow) {
  try {
    const baseFormula = sheet.getRange(CONSTANTS.START_ROW, column).getFormula().trim();
    if (!baseFormula) {
      Logger.log(`列 ${String.fromCharCode(64 + column)} の基準数式が存在しません。スキップします。`);
      return {
        updates: [],
        logs: [[
          new Date(),
          String.fromCharCode(64 + column),
          '',
          '',
          '',
          '',
          '基準数式が存在しません（スキップ）'
        ]]
      };
    }

    Logger.log(`列 ${String.fromCharCode(64 + column)} の基準数式: ${baseFormula}`);
    const updates = [];
    const logs = [];
    const formulas = sheet.getRange(CONSTANTS.START_ROW, column, lastRow - CONSTANTS.START_ROW + 1, 1).getFormulas();

    for (let i = 0; i < formulas.length; i++) {
      const currentRow = i + CONSTANTS.START_ROW;
      const currentFormula = (formulas[i][0] || '').toString().trim();
      const expectedFormula = baseFormula.replace(/(\$?[A-Z]+\$?)4\b/g, (m, col) => `${col}${currentRow}`).trim();

      let result;
      if (!currentFormula) {
        result = '数式が入力されていません';
      } else if (currentFormula === expectedFormula) {
        result = '数式が正しいです。修正は不要です';
      } else {
        result = '数式が異なります。修正しました';
        updates.push({ row: currentRow, column: column, formula: expectedFormula });
      }

      Logger.log(`列 ${String.fromCharCode(64 + column)} 行 ${currentRow}: ${result}`);
      logs.push([
        new Date(),
        String.fromCharCode(64 + column),
        currentRow,
        `"${baseFormula}"`,
        `"${currentFormula || '（空白）'}"`,
        `"${expectedFormula}"`,
        result
      ]);
    }

    return { updates, logs };
  } catch (error) {
    Logger.log(`列 ${String.fromCharCode(64 + column)} の処理中にエラーが発生しました: ${error.message}`);
    return { updates: [], logs: [] };
  }
}

function applyBatchUpdates(sheet, updates) {
  const batchSize = 100;
  for (let i = 0; i < updates.length; i += batchSize) {
    const batch = updates.slice(i, i + batchSize);
    batch.forEach(update => {
      sheet.getRange(update.row, update.column).setFormula(update.formula);
    });
  }
}
