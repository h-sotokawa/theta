function onButtonClick() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('フォーム作成');

  // シートが見つからなければエラーメッセージを表示して終了
  if (!sheet) {
    SpreadsheetApp.getUi().alert('シート「フォーム作成」が見つかりません。「フォーム作成」という名前でシートを作成してください。');
    return;
  }
  
  // FORM_TITLE をシートの特定のセルから取得（例：C2セル）
  const formTitle = sheet.getRange('C2').getValue();
  
  // ポップアップを表示して確認
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`「${formTitle}」 でフォームを作成しますか？作成には30秒ほど時間がかかります`, ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.NO) {
    return; // ユーザーが「NO」を選択した場合、処理を中止
  }
  
  // SPREADSHEET_ID を自動取得
  const spreadsheetId = spreadsheet.getId();
  
  // フォルダIDも特定のセルから取得（例：C3セル）
  const folderId = sheet.getRange('C3').getValue();

  // フォーム作成処理を実行
  createGoogleFormWithEmailCollection(formTitle, folderId, spreadsheetId);
}

function createGoogleFormWithEmailCollection(formTitle, folderId, spreadsheetId) {
  if (!formTitle || !folderId || !spreadsheetId) {
    logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'フォームタイトル、フォルダID、またはスプレッドシートIDが設定されていません。FORM_TITLE、FOLDER_ID、SPREADSHEET_IDを設定してください。FORM_TITLEはC2セルに入力(端末の名前になります)、フォルダーIDはフォームを作成したいドライブのフォルダのURL末尾30文字程度の英文字をC3セルに入力してください');
    return;
  }
  
  // フォルダ内に同名のファイルが存在するか確認
  const parentFolder = DriveApp.getFolderById(folderId);
  const files = parentFolder.getFilesByName(formTitle);
  if (files.hasNext()) {
    logToSheet(SpreadsheetApp.getActiveSpreadsheet(), '指定したフォルダ内に同名のファイルが存在します。異なるタイトルを指定してください。');
    return;
  }
  
  // フォームを作成し基本設定を行う
  const form = createForm(formTitle, folderId);
  
  // 既存のスプレッドシートにリンク
  linkResponseSpreadsheet(form, spreadsheetId);
  
  // 質問を追加
  addQuestions(form);

  // 回答毎にセクションに移動する設定を追加
  setSectionNavigation(form);
  
  // フォームの編集URLと送信URLをログに出力
  logFormUrls(form);
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'renameResponseSheet の処理を開始するまで 1 秒待機します...');
  Utilities.sleep(1000);
  // renameResponseSheet(spreadsheetId, formTitle);
  
  // フォーム送信リンクをQRコードとして出力
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'フォーム送信リンクのQRコードを生成中...（QuickChartを使用）');
generateQrCode(form.getPublishedUrl(), formTitle);
}

function createForm(formTitle, folderId) {
  // 指定されたフォルダにフォームを作成
  const parentFolder = DriveApp.getFolderById(folderId);
  const form = FormApp.create(formTitle);
  const file = DriveApp.getFileById(form.getId());
  parentFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  // メールアドレス収集を有効に設定
  form.setCollectEmail(true);
  
  // フォームの説明を設定
  form.setDescription(formTitle + 'の管理用フォームです。');
  
  return form;
}

function linkResponseSpreadsheet(form, spreadsheetId) {
  // 既存のスプレッドシートを取得し、フォームとスプレッドシートをリンク
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // フォームのリンク先としてスプレッドシートを設定
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), '回答用スプレッドシートのURL: ' + spreadsheet.getUrl());
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'フォームの回答シートがスプレッドシートにリンクされました。');
}

// 動作しないので使わない
function renameResponseSheet(spreadsheetId, formTitle) {
  // スプレッドシートのシート名を変更
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheets = spreadsheet.getSheets();
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), '取得したすべてのシート名:');
  for (const sheet of sheets) {
    logToSheet(SpreadsheetApp.getActiveSpreadsheet(), sheet.getName());
  }
  let sheetRenamed = false;
  for (const sheet of sheets) {
    if (sheet.getName().startsWith('フォームの回答 ')) {
      sheet.setName(formTitle); // シート名をフォームタイトルに変更
      logToSheet(SpreadsheetApp.getActiveSpreadsheet(), `シート名を「${formTitle}」に変更しました。`);
      sheetRenamed = true;
      break;
    }
  }
  if (!sheetRenamed) {
    logToSheet(SpreadsheetApp.getActiveSpreadsheet(), '対象のシートが見つからなかったため、シート名を変更できませんでした。');
  }
}

function addQuestions(form) {
  addStatusQuestion(form);
  addLendSectionQuestions(form);
  addReturnSectionQuestions(form);
  addRepairSectionQuestions(form);
  addBorrowSectionQuestions(form);
  addConfirmationSectionQuestions(form);
}

function addStatusQuestion(form) {
  // 質問1: 単一選択の質問
  const statusItem = form.addMultipleChoiceItem()
      .setTitle('ステータス')
      .setChoiceValues(['設定や修理など一時移動', '(管理者用)端末棚卸', '回収(車載返却含む)', '代替貸出', '社外に持ち出し(社内含む)'])
      .setRequired(true);
}

function addLendSectionQuestions(form) {
  // 貸出セクションを追加
  form.addPageBreakItem().setTitle('貸出セクション');
  
  // 貸出に関する質問を追加
  form.addTextItem()
      .setTitle('顧客名')
      .setRequired(true);
  form.addTextItem()
      .setTitle('預かり機の製造番号')
      .setRequired(false);
  form.addTextItem()
      .setTitle('備考')
      .setRequired(false);
  form.addTextItem()
      .setTitle('お預かり証No.')
      .setHelpText('お預かり証No.を入力してください。(お預かり証がない場合は省略)') 
      .setRequired(false);

  // 貸出セクションの最後で送信するように設定
  form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function addReturnSectionQuestions(form) {
  // 回収セクションを追加
  form.addPageBreakItem().setTitle('回収セクション');
  
  // 回収に関する質問を追加
  form.addMultipleChoiceItem()
      .setTitle('返却方法を選択してください')
      .setChoiceValues(['オフィスに返却', '社外に持ち出し'])
      .setRequired(true);
  form.addMultipleChoiceItem()
      .setTitle('修理の必要性')
      .setChoiceValues(['なし', 'あり'])
      .setRequired(true);
  form.addTextItem()
      .setTitle('症状')
      .setRequired(false);
  form.addTextItem()
      .setTitle('備考（回収）')
      .setRequired(false);
  
  // 回収セクションの最後で送信するように設定
  form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function addRepairSectionQuestions(form) {
  // 修理・設定セクションを追加
  form.addPageBreakItem().setTitle('修理・設定セクション');
  
  // 修理・設定に関する質問を追加
  form.addMultipleChoiceItem()
      .setTitle('修理ステータス')
      .setChoiceValues(['修理中', '修理完了', '設定中'])
      .setRequired(true);
  form.addTextItem()
      .setTitle('修理報告書のリンク (Google ドライブなどにアップロードしてリンクを共有)')
      .setRequired(false);
  form.addTextItem()
      .setTitle('症状・理由')
      .setRequired(false);
  form.addTextItem()
      .setTitle('修理内容')
      .setRequired(false);
  
  // 修理・設定セクションの最後で送信するように設定
  form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function addBorrowSectionQuestions(form) {
  // 社外持ち出しセクションを追加
  form.addPageBreakItem().setTitle('社外持ち出しセクション');
  
  // 社外持ち出しに関する質問を追加
  form.addTextItem()
      .setTitle('理由')
      .setRequired(true);
  
  // 社外持ち出しセクションの最後で送信するように設定
  form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function addConfirmationSectionQuestions(form) {
  // 預かり証の確認に関する質問を追加
  form.addTextItem()
      .setTitle('預かり証の確認内容を入力してください')
      .setRequired(true);
}

function setSectionNavigation(form) {
  // ステータス質問アイテムを取得
  const statusItem = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem();
  
  // 各選択肢に対する遷移先を設定
  const lendSection = form.getItems(FormApp.ItemType.PAGE_BREAK).find(item => item.getTitle() === '貸出セクション').asPageBreakItem();
  const repairSection = form.getItems(FormApp.ItemType.PAGE_BREAK).find(item => item.getTitle() === '修理・設定セクション').asPageBreakItem();
  const borrowSection = form.getItems(FormApp.ItemType.PAGE_BREAK).find(item => item.getTitle() === '社外持ち出しセクション').asPageBreakItem();
  const returnSection = form.getItems(FormApp.ItemType.PAGE_BREAK).find(item => item.getTitle() === '回収セクション').asPageBreakItem();
  const choices = [
    statusItem.createChoice('設定や修理など一時移動', repairSection),
    statusItem.createChoice('(管理者用)端末棚卸', FormApp.PageNavigationType.SUBMIT),
    statusItem.createChoice('回収(車載返却含む)', returnSection),
    statusItem.createChoice('代替貸出', lendSection),
    statusItem.createChoice('社外に持ち出し(社内含む)', borrowSection)
  ];
  statusItem.setChoices(choices);

  // 回収セクションの最初の質問で「社外に持ち出し」を選択した場合、質問1に遷移するように設定
  const returnItem = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE).find(item => item.getTitle() === '返却方法を選択してください').asMultipleChoiceItem();
  const returnChoices = [
    returnItem.createChoice('オフィスに返却', FormApp.PageNavigationType.SUBMIT),
    returnItem.createChoice('社外に持ち出し', FormApp.PageNavigationType.RESTART)
  ];
  returnItem.setChoices(returnChoices);
}

function logFormUrls(form) {
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'フォームの編集リンク: ' + form.getEditUrl());
  logToSheet(SpreadsheetApp.getActiveSpreadsheet(), 'フォームの送信リンク: ' + form.getPublishedUrl());
}

function logToSheet(spreadsheet, message) {
  let logSheet = spreadsheet.getSheetByName('createForm_script_log');
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('createForm_script_log');
  }
  logSheet.appendRow([new Date(), message]);
}

function generateQrCode(url, formTitle) {
  const qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(url)}&size=200`;
  const response = UrlFetchApp.fetch(qrCodeUrl);
  const blob = response.getBlob().setName(`${formTitle}_QR_Code.png`);

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let qrSheet = spreadsheet.getSheetByName('QRコード');
  
  // シートが存在しない場合は作成してヘッダーを追加
  if (!qrSheet) {
    qrSheet = spreadsheet.insertSheet('QRコード');
    qrSheet.getRange("A1").setValue("作成日");
    qrSheet.getRange("B1").setValue("管理番号");
    qrSheet.getRange("C1").setValue("QRコード");
  }

  const lastRow = qrSheet.getLastRow() + 1;

  try {
    // 日付と管理番号の設定
    const creationDate = new Date();
    qrSheet.getRange(lastRow, 1).setValue(creationDate);
    qrSheet.getRange(lastRow, 2).setValue(formTitle || "未設定");

    // QRコード画像のサイズ
    const imageWidth = 200;  // ピクセル
    const imageHeight = 200; // ピクセル

    // 列幅と行の高さの調整
    qrSheet.setColumnWidth(3, imageWidth + 10);    // 3列目の幅をQRコード+10の幅に設定
    qrSheet.setRowHeight(lastRow, imageHeight + 10); // 現在の行の高さをQRコード+10の高さに設定

    // QRコード画像の挿入
    qrSheet.insertImage(blob, 3, lastRow);

    // セルの範囲を取得して枠線を設定
    const range = qrSheet.getRange(lastRow, 1, 1, 3); // 作成日、管理番号、QRコードの列範囲
    range.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ログの保存
    logToSheet(spreadsheet, 'QRコードと情報をシートに追加しました。');
  } catch (e) {
    logToSheet(spreadsheet, 'QRコードの追加に失敗しました: ' + e.message);
  }
}


