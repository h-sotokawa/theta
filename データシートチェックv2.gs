function onOpen() {
  let ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  let menu = ui.createMenu("【メニュー】");  // Uiクラスからメニューを作成する
  menu.addItem('実行する', 'main');
  menu.addToUi();                            // メニューをUiクラスに追加する
}

function main() {

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

// function main_() {

//   const prefix = PropertiesService.getScriptProperties().getProperty("PREFIX").split(",");
//   console.log(prefix);

//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheets = ss.getSheets();

//   sheets.forEach(sheet => {

//     sheet.activate();
//     const sheetName = sheet.getName();
//     const flag = prefix.some(x => sheetName.includes(x.trim()));
//     if (flag === false) return;

//     console.log(`sheetName: ${sheetName}`);

//     const [header, ...values] = sheet.getDataRange().getDisplayValues();
//     const output = [];
//     const items1 = ["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり機の製造番号", "備考", "お預かり証No."];
//     const items2 = [];

//     header.forEach(x => {
//       if (!items1.includes(x)) items2.push(x);
//     })
//     const newHeader = [...items1, ...items2];

//     // console.log(header);
//     // console.log(items2);
//     // console.log(newHeader);

//     for (const value of values) {

//       const valueObj = {};
//       header.forEach((x, index) => valueObj[x] = value[index]);
//       // console.log(valueObj);
//       const rowValue = [];
//       newHeader.forEach(x => rowValue.push(valueObj[x]));
//       output.push(rowValue);
//     }

//     sheet.clearContents();
//     sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
//     sheet.getRange(2, 1, output.length, output[0].length).setValues(output);


//   });

// }

