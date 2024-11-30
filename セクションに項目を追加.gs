// フォルダIDを指定してください
const FOLDER_ID = "<YOUR_FOLDER_ID>";

function main() {
  const logSheet = createLogSheet();
  addFieldToFormsInFolder(logSheet);
}

function addFieldToFormsInFolder(logSheet) {
  try {
    logSheet.appendRow(["スクリプトの実行を開始しました。", new Date()]);
    
    // 指定したフォルダを取得
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_FORMS);
    
    let formCount = 0;
    while (files.hasNext()) {
      formCount++;
      const file = files.next();
      const formId = file.getId();
      try {
        const form = FormApp.openById(formId);
        logSheet.appendRow([`Form '${form.getTitle()}' の処理を開始します。`, new Date()]);

        // フォームのセクションを取得
        const items = form.getItems();
        let section2Index = -1;
        let sectionCount = 0;
        
        // セクションブレークを数える
        for (let i = 0; i < items.length; i++) {
          if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
            sectionCount++;
            if (sectionCount === 2) {
              section2Index = i;
              break;
            }
          }
        }

        // セクション2が存在する場合に「預かり証No.」を追加
        if (section2Index !== -1) {
          form.addTextItem()
              .setTitle("預かり証No.")
              .setHelpText("預かり証の番号を入力してください。")
              .moveToPageBreakItem(items[section2Index]);
          logSheet.appendRow([`Form '${form.getTitle()}' に「預かり証No.」を追加しました。`, new Date()]);
        } else {
          logSheet.appendRow([`Form '${form.getTitle()}' にはセクション2が見つかりませんでした。`, new Date()]);
        }
      } catch (e) {
        logSheet.appendRow([`Form ID '${formId}' の処理中にエラーが発生しました: ${e}`, new Date()]);
      }
    }
    logSheet.appendRow([`対象のフォーム数: ${formCount}`, new Date()]);
    logSheet.appendRow(["スクリプトの実行を終了しました。", new Date()]);
  } catch (e) {
    logSheet.appendRow([`スクリプト全体でエラーが発生しました: ${e}`, new Date()]);
  }
}

function createLogSheet() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const logFile = folder.createFile("LogSheet", "");
  const logSpreadsheet = SpreadsheetApp.openById(logFile.getId());
  const sheet = logSpreadsheet.getSheets()[0];
  sheet.appendRow(["メッセージ", "タイムスタンプ"]);
  return sheet;
}
