function main() {
  addFieldToFormsInFolder();
}

function addFieldToFormsInFolder() {
  try {
    console.log("スクリプトの実行を開始しました。", new Date());
    
    // スクリプトプロパティからフォルダIDを取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const folderId = scriptProperties.getProperty("FOLDER_ID");
    if (!folderId) {
      throw new Error("フォルダIDが設定されていません。スクリプトプロパティに 'FOLDER_ID' を設定してください。");
    }
    
    // 指定したフォルダを取得
    const folder = DriveApp.getFolderById(folderId);
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
        let section2StartIndex = -1;
        let sectionCount = 0;
        
        // セクションブレークを数えることでセクション2の開始インデックスを特定
        for (let i = 0; i < items.length; i++) {
          if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
            sectionCount++;
            if (sectionCount === 2) {
              section2StartIndex = i;
              break;
            }
          }
        }

        // セクション2が存在する場合に「預かり証No.」を追加
        if (section2StartIndex !== -1) {
          let section2EndIndex = section2StartIndex;
          for (let i = section2StartIndex + 1; i < items.length; i++) {
            if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
              break;
            }
            section2EndIndex = i;
          }

          // 新しい質問を追加し、セクション2の最後に移動
          const newItem = form.addTextItem()
              .setTitle("預かり証No.")
              .setHelpText("預かり証の番号を入力してください。");
          form.moveItem(newItem.getIndex(), section2EndIndex + 1);
          console.log(`Form '${form.getTitle()}' に「預かり証No.」を追加しました。`, new Date());
        } else {
          console.log(`Form '${form.getTitle()}' にはセクション2が見つかりませんでした。`, new Date());
        }
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
