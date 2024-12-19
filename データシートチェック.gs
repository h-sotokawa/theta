// スプレッドシート内の列順を検証し修正する関数
function organizeDataSheets() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var PREFIX = scriptProperties.getProperty('PREFIX');
  var SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');

  if (!PREFIX || !SPREADSHEET_ID) {
    Logger.log("PREFIXまたはSPREADSHEET_IDが指定されていません。");
    return;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets().filter(sheet => sheet.getName().includes(PREFIX));

  var correctHeaders = ["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり証No.", "備考"];
  var logSheet = ensureLogSheet(ss, 'organize_dataSheets_log');

  sheets.forEach(sheet => {
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
      logAction(logSheet, sheet.getName(), "シートにデータがないためスキップ");
      return;
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (arraysEqual(headers, correctHeaders)) {
      logAction(logSheet, sheet.getName(), "列順は正しい状態でした");
      return;
    }

    rearrangeColumns(sheet, headers, correctHeaders, logSheet);
  });
}

function rearrangeColumns(sheet, headers, correctHeaders, logSheet) {
  var moveActions = [];
  correctHeaders.forEach((header, targetIndex) => {
    var currentIndex = headers.indexOf(header);
    if (currentIndex !== -1 && currentIndex !== targetIndex) {
      try {
        if (currentIndex !== targetIndex) {
          var sourceIndex = currentIndex + 1;
          var destinationIndex = targetIndex + 1;
          if (sourceIndex < destinationIndex) {
            destinationIndex++;
          }
          sheet.moveColumns(sourceIndex, destinationIndex);
          moveActions.push(`列「${header}」を${sourceIndex}列目から${destinationIndex - (sourceIndex < destinationIndex ? 1 : 0)}列目に移動`);
        }
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      } catch (e) {
        logAction(logSheet, sheet.getName(), `列移動エラー: ${e.message}`);
      }
    }
  });

  if (moveActions.length > 0) {
    logAction(logSheet, sheet.getName(), moveActions.join("; "));
  } else {
    logAction(logSheet, sheet.getName(), "列順を修正しましたが、移動操作はありませんでした");
  }
}

function ensureLogSheet(ss, logSheetName) {
  var logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    logSheet.appendRow(["処理日時", "シート名", "アクション"]);
  }
  return logSheet;
}

function logAction(logSheet, sheetName, action) {
  logSheet.appendRow([new Date(), sheetName, action]);
}

function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  for (var i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }
  return true;
}
