// ログシートの作成と管理
function createLogSheet_notification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Notification_log');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('Notification_log');
    // ヘッダー行を設定
    logSheet.getRange('A1:E1').setValues([['タイムスタンプ', 'イベント', 'ステータス', 'メール送信先', 'メッセージ']]);
    logSheet.getRange('A1:E1').setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
  
  return logSheet;
}

// ログを記録する関数
function writeLog_notification(event, status, recipient, message) {
  const logSheet = createLogSheet_notification();
  const timestamp = new Date();
  const logData = [[timestamp, event, status, recipient, message]];
  logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 5).setValues(logData);
}

// フォーム送信時のトリガーを設定する関数
function createFormSubmitTrigger_notification() {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onFormSubmit_notification') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 新しいトリガーを作成
    ScriptApp.newTrigger('onFormSubmit_notification')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
    
    writeLog_notification('トリガー設定', '成功', 'システム', 'フォーム送信トリガーを設定しました');
  } catch (error) {
    writeLog_notification('トリガー設定', 'エラー', 'システム', 'エラー: ' + error.toString());
  }
}

// フォーム送信時の処理
function onFormSubmit_notification(e) {
  try {
    writeLog_notification('データ監視', '開始', 'システム', '最新データの監視を開始します');
    
    // スプレッドシートから回答を取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('main');
    if (!sheet) {
      writeLog_notification('データ監視', 'エラー', 'システム', 'mainシートが見つかりません');
      return;
    }
    
    // G列（日付）の最終行を取得
    const lastRow = sheet.getLastRow();
    const dateRange = sheet.getRange(2, 7, lastRow - 1, 1).getValues(); // G列のデータ（ヘッダー行を除く）
    
    // 最新の日付を探す
    let latestDate = new Date(0);
    let latestRowIndex = -1;
    
    dateRange.forEach((date, index) => {
      if (date[0] instanceof Date && !isNaN(date[0])) {
        if (date[0] > latestDate) {
          latestDate = date[0];
          latestRowIndex = index + 2; // ヘッダー行を考慮して+2
        }
      }
    });
    
    if (latestRowIndex === -1) {
      writeLog_notification('データ監視', 'エラー', 'システム', '有効な日付が見つかりません');
      return;
    }
    
    // 最新行のデータを取得
    const latestData = sheet.getRange(latestRowIndex, 1, 1, 11).getValues()[0];
    
    // メール送信先を取得
    const recipientEmail = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL');
    
    if (!recipientEmail) {
      writeLog_notification('データ監視', 'エラー', 'システム', 'メール送信先が設定されていません');
      return;
    }
    
    // 型番の表示を設定
    const modelNumber = latestData[10] ? latestData[10] : "!!登録されていません!! ※mainシートK列を確認してください";
    
    // ステータスに"貸出"が含まれるかチェック
    const isRental = latestData[3].toString().includes("貸出");
    
    // メール本文を作成
    let message = 
      `【${latestData[2]}】のステータスが変更されました。\n\n` +
      `型番：${modelNumber}\n` +
      `ステータス：${latestData[3]}\n` +
      `変更日時：${Utilities.formatDate(latestData[6], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}\n` +
      `変更者：${latestData[7]}`;
    
    // ステータスが"貸出"を含む場合のみ備考と預かり証No.を追加
    if (isRental) {
      message += `\n備考：${latestData[8]}\n` +
                `預かり証No.：${latestData[9] || "未設定"}`;
    }
    
    // メールの件名と本文を作成
    const subject = '代替機 : ステータス変更通知';
    const spreadsheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
    const body = message + 
      `\n\n詳細はスプレッドシートでご確認ください。\n` +
      `URL：${spreadsheetUrl}`;
    
    writeLog_notification('データ監視', '送信準備', recipientEmail, 'メール送信を試みます');
    
    // メールを送信
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: body
    });
    
    writeLog_notification('データ監視', '成功', recipientEmail, 'メール送信が完了しました');
    
  } catch (error) {
    writeLog_notification('データ監視', 'エラー', 'システム', 'エラー: ' + error.toString());
  }
}

// フォームの回答シートを取得する関数
function getFormResponseSheet(ss) {
  try {
    // フォームを取得
    const form = FormApp.getActiveForm();
    if (!form) {
      writeLog('シート検索', 'エラー', 'システム', 'アクティブなフォームが見つかりません。フォームを開いてから実行してください。');
      return null;
    }

    const formTitle = form.getTitle();
    writeLog_notification('シート検索', '情報', 'システム', `検索対象のシート名: "${formTitle}"`);
    
    // フォームのタイトルと同じ名前のシートを探す
    const sheet = ss.getSheetByName(formTitle);
    if (sheet) {
      writeLog_notification('シート検索', '成功', 'システム', `フォームの回答シート "${formTitle}" を見つけました`);
      return sheet;
    }
    
    // シートが見つからない場合、全てのシート名をログに記録
    const allSheets = ss.getSheets();
    const sheetNames = allSheets.map(sheet => sheet.getName());
    writeLog_notification('シート検索', 'エラー', 'システム', 
      `フォームの回答シート "${formTitle}" が見つかりません。\n` +
      `検索対象のシート名: "${formTitle}"\n` +
      `利用可能なシート: ${sheetNames.join(', ')}`
    );
    
    return null;
  } catch (error) {
    writeLog_notification('シート検索', 'エラー', 'システム', 
      `フォームの取得中にエラーが発生しました: ${error.toString()}\n` +
      `エラーの詳細: ${error.stack}`
    );
    return null;
  }
}

// 利用可能なシート名を確認する関数
function checkAvailableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  writeLog_notification('シート確認', '情報', 'システム', `利用可能なシート: ${sheetNames.join(', ')}`);
  return sheetNames;
}

// フォームの情報を確認する関数
function checkFormInfo() {
  try {
    const form = FormApp.getActiveForm();
    if (!form) {
      writeLog_notification('フォーム確認', 'エラー', 'システム', 'アクティブなフォームが見つかりません。フォームを開いてから実行してください。');
      return null;
    }

    const formTitle = form.getTitle();
    writeLog_notification('フォーム確認', '情報', 'システム', 
      `フォーム情報:\n` +
      `タイトル: ${formTitle}\n` +
      `ID: ${form.getId()}\n` +
      `URL: ${form.getPublishedUrl()}`
    );
    return formTitle;
  } catch (error) {
    writeLog_notification('フォーム確認', 'エラー', 'システム', 
      `フォームの取得中にエラーが発生しました: ${error.toString()}\n` +
      `エラーの詳細: ${error.stack}`
    );
    return null;
  }
}

// フォームとスプレッドシートの関連付けを確認する関数
function checkFormSpreadsheetConnection() {
  try {
    const form = FormApp.getActiveForm();
    if (!form) {
      writeLog_notification('接続確認', 'エラー', 'システム', 'アクティブなフォームが見つかりません。フォームを開いてから実行してください。');
      return;
    }

    const destinationType = form.getDestinationType();
    const destinationId = form.getDestinationId();
    
    writeLog_notification('接続確認', '情報', 'システム', 
      `フォームとスプレッドシートの接続情報:\n` +
      `接続タイプ: ${destinationType}\n` +
      `接続先ID: ${destinationId}`
    );

    if (destinationType !== FormApp.DestinationType.SPREADSHEET) {
      writeLog_notification('接続確認', 'エラー', 'システム', 'フォームがスプレッドシートに接続されていません。');
      return;
    }

    const ss = SpreadsheetApp.openById(destinationId);
    const sheet = ss.getSheetByName(form.getTitle());
    
    if (sheet) {
      writeLog_notification('接続確認', '成功', 'システム', 
        `フォームとスプレッドシートの接続が確認できました。\n` +
        `フォームタイトル: ${form.getTitle()}\n` +
        `シート名: ${sheet.getName()}`
      );
    } else {
      writeLog_notification('接続確認', 'エラー', 'システム', 
        `フォームとスプレッドシートは接続されていますが、対応するシートが見つかりません。\n` +
        `フォームタイトル: ${form.getTitle()}\n` +
        `利用可能なシート: ${ss.getSheets().map(s => s.getName()).join(', ')}`
      );
    }
  } catch (error) {
    writeLog_notification('接続確認', 'エラー', 'システム', 
      `接続確認中にエラーが発生しました: ${error.toString()}\n` +
      `エラーの詳細: ${error.stack}`
    );
  }
}

// テスト実行用の関数
function testFormSubmit() {
  onFormSubmit();
}

// スクリプトプロパティの設定を確認する関数
function checkScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  const email = properties.getProperty('NOTIFICATION_EMAIL');
  
  if (!email) {
    writeLog_notification('設定確認', 'エラー', 'システム', 'メール送信先が設定されていません');
  } else {
    writeLog_notification('設定確認', '成功', email, '現在の通知メールアドレスを確認しました');
  }
}

// スクリプトプロパティを設定する関数
function setNotificationEmail(email) {
  if (!email) {
    writeLog_notification('メール設定', 'エラー', 'システム', 'メールアドレスが指定されていません');
    return;
  }
  
  PropertiesService.getScriptProperties().setProperty('NOTIFICATION_EMAIL', email);
  writeLog_notification('メール設定', '成功', email, '通知メールアドレスを設定しました');
} 