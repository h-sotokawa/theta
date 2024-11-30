// フォルダIDを指定してください
const FOLDER_ID = "<YOUR_FOLDER_ID>";

function addFieldToFormsInFolder() {
  // 指定したフォルダを取得
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType(MimeType.GOOGLE_FORMS);

  while (files.hasNext()) {
    const file = files.next();
    const formId = file.getId();
    const form = FormApp.openById(formId);

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
      Logger.log(`Form '${form.getTitle()}' に「預かり証No.」を追加しました。`);
    } else {
      Logger.log(`Form '${form.getTitle()}' にはセクション2が見つかりませんでした。`);
    }
  }
} 
