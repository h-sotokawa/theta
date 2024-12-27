function onOpen() {
  let ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  let menu = ui.createMenu("【メニュー】");  // Uiクラスからメニューを作成する
  menu.addItem('実行する', 'main');
  menu.addToUi();                            // メニューをUiクラスに追加する
}

function dataSheetColumnCheckMain() {

  const prefix = PropertiesService.getScriptProperties().getProperty("PREFIX")?.split(",");
  console.log(prefix);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const items = ["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり機の製造番号", "備考", "お預かり証No."];

  sheets.forEach(sheet => {

    const sheetName = sheet.getName();
    const flag = prefix.some(x => sheetName.includes(x.trim()));
    if (flag === false) return;
    sheet.activate();
    console.log(`sheetName: ${sheetName}`);

    let index = 0;
    items.forEach(key => {
      const [header] = sheet.getDataRange().getDisplayValues();
      const column = header.indexOf(key) + 1;
      if (column <= 0) return;
      index++;
      const aiNotion = sheet.getRange(1, column).getA1Notation();
      const rng = sheet.getRange(1, column);
      // console.log(`key: ${key}`);
      // console.log(`index: ${index}`);
      // console.log(`column: ${column}`);
      if (index === column) return;
      sheet.moveColumns(rng, index);

    });

  });

}
