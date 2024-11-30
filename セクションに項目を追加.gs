// フォルダIDを指定してください
const FOLDER_ID = "<YOUR_FOLDER_ID>";

function main() {
  addFieldToFormsInFolder();
}

function addFieldToFormsInFolder() {
  try {
    console.log("スクリプトの実行を開始しました。", new Date());
    
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
        console.log(`Form '${form.getTitle()}' の処理を開始します。`, new Date());

        // フォームのセクションを取得
        const items = form.getItems();
        let targetIndex = Math.min(items.length, 3); // 4番目のインデックスに移動するため、0ベースで3

        const newItem = form.addTextItem()
            .setTitle("預かり証No.")
            .setHelpText("預かり証の番号を入力してください。");
        form.moveItem(newItem.getIndex(), targetIndex);
        console.log(`Form '${form.getTitle()}' の4番目に「預かり証No.」を追加しました。`, new Date());
      } catch (e) {
        console.log(`Form ID '${formId}' の処理中にエラーが発生しました: ${e}`, new Date());
      }
    }
    console.log(`対象のフォーム数: ${formCount}`, new Date());
    console.log("スクリプトの実行を終了しました。", new Date());
  } catch (e) {
    console.log(`スクリプト全体でエラーが発生しました: ${e}`, new Date());
  }
}
