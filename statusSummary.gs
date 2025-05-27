/**
 * 各拠点の預り機ステータス集計スクリプト
 * 各拠点のmainシートのA列から預り機ステータスを取得し、
 * 種別列から機器種別を取得して、SPREADSHEET_ID_DESTINATIONの"サマリー"シートに集計結果を記入する
 */

// ステータス定数定義
const STATUS_CONSTANTS = {
  STATUSES: [
    '1.代替機を貸し出しているが修理が完了していつでも返却できる台数',
    '2.商談や金額の問題で返却出来ない台数',
    '3.代替機を貸し出し中だが、お客様より代替機を使うので返却を拒否されている台数',
    '4.HW延長保守にて貸し出している台数 ※OS入れ替えやサービス終了を含む'
  ],
  DEVICE_TYPES: ['SV', 'CL', 'プリンタ', 'その他'],
  MAIN_SHEET_NAME: 'main',
  SUMMARY_SHEET_NAME: 'サマリー',
  STATUS_COLUMN: 1, // A列
  TYPE_COLUMN: null, // 種別列（動的に検索）
  HEADER_TEXT: '預かり機のステータス',
  TYPE_HEADER_TEXT: '代替機種別',
  // プリンタ用スプレッドシートIDのプロパティ名
  PRINTER_SPREADSHEET_IDS: {
    OSAKA: 'SPREADSHEET_ID_SOURCE_OSAKA_PRINTER',
    HYOGO: 'SPREADSHEET_ID_SOURCE_HYOGO_PRINTER'
  }
};

/**
 * メイン実行関数：全拠点のステータスを集計してサマリーシートに記入
 */
function aggregateAllLocationStatusMain() {
  console.log('===== ステータス集計処理開始 =====');
  
  try {
    // 集計結果を格納するオブジェクト
    const aggregatedData = initializeAggregatedData();
    
    // 各拠点のステータスを集計
    const locations = Object.keys(Config.LOCATIONS);
    
    for (const locationKey of locations) {
      console.log(`拠点「${locationKey}」の処理を開始`);
      
      try {
        const locationData = aggregateLocationStatus(locationKey);
        
        // 拠点データを全体集計に追加
        for (const status of STATUS_CONSTANTS.STATUSES) {
          for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
            aggregatedData.total[status][deviceType] += locationData[status][deviceType] || 0;
          }
        }
        aggregatedData.byLocation[locationKey] = locationData;
        
        console.log(`拠点「${locationKey}」の集計完了:`, locationData);
        
      } catch (error) {
        console.error(`拠点「${locationKey}」の処理中にエラーが発生:`, error.message);
        // エラーが発生した拠点はスキップして続行
        continue;
      }
    }
    
    // サマリーシートに結果を記入
    writeToSummarySheet(aggregatedData);
    
    console.log('===== ステータス集計処理完了 =====');
    console.log('集計結果:', aggregatedData);
    
  } catch (error) {
    console.error('ステータス集計処理中にエラーが発生:', error.message);
    throw error;
  }
}

/**
 * 集計データの初期化
 */
function initializeAggregatedData() {
  const data = {
    total: {},
    byLocation: {}
  };
  
  // ステータス×種別のマトリックスを初期化
  for (const status of STATUS_CONSTANTS.STATUSES) {
    data.total[status] = {};
    for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
      data.total[status][deviceType] = 0;
    }
  }
  
  return data;
}

/**
 * 指定拠点のステータス×種別マトリックスを集計
 * @param {string} locationKey - 拠点キー（osaka, kobe, himeji）
 * @returns {Object} ステータス×種別の集計結果
 */
function aggregateLocationStatus(locationKey) {
  const locationConfig = Config.getLocationConfig(locationKey);
  const statusCounts = {};
  
  // ステータス×種別のマトリックスを初期化
  for (const status of STATUS_CONSTANTS.STATUSES) {
    statusCounts[status] = {};
    for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
      statusCounts[status][deviceType] = 0;
    }
  }
  
  // 大阪の場合は複数のスプレッドシートを処理
  if (locationKey === 'osaka') {
    // 通常の大阪のスプレッドシートを処理
    for (const sourceId of locationConfig.sourceIds) {
      if (sourceId) {
        const sheetCounts = getStatusCountsFromSheet(sourceId);
        // 各ステータス×種別の数を加算
        for (const status of STATUS_CONSTANTS.STATUSES) {
          for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
            statusCounts[status][deviceType] += sheetCounts[status][deviceType] || 0;
          }
        }
      }
    }
    
    // 大阪プリンタ用のスプレッドシートも処理
    const scriptProperties = PropertiesService.getScriptProperties();
    const osakaPrinterId = scriptProperties.getProperty(STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.OSAKA);
    if (!osakaPrinterId) {
      throw new Error(`大阪プリンタ用スプレッドシートID「${STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.OSAKA}」が設定されていません`);
    }
    
    try {
      const printerCounts = getStatusCountsFromSheet(osakaPrinterId);
      // 各ステータス×種別の数を加算
      for (const status of STATUS_CONSTANTS.STATUSES) {
        for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
          statusCounts[status][deviceType] += printerCounts[status][deviceType] || 0;
        }
      }
      console.log(`大阪プリンタ用スプレッドシート ${osakaPrinterId} の処理が完了しました`);
    } catch (error) {
      throw new Error(`大阪プリンタ用スプレッドシート ${osakaPrinterId} の処理中にエラーが発生しました: ${error.message}`);
    }
  } else if (locationKey === 'kobe') {
    // 神戸の場合は通常のスプレッドシートと兵庫プリンタを処理
    if (locationConfig.sourceId) {
      const sheetCounts = getStatusCountsFromSheet(locationConfig.sourceId);
      // 各ステータス×種別の数を加算
      for (const status of STATUS_CONSTANTS.STATUSES) {
        for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
          statusCounts[status][deviceType] += sheetCounts[status][deviceType] || 0;
        }
      }
    }
    
    // 兵庫プリンタ用のスプレッドシートも処理
    const scriptProperties = PropertiesService.getScriptProperties();
    const hyogoPrinterId = scriptProperties.getProperty(STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.HYOGO);
    if (!hyogoPrinterId) {
      throw new Error(`兵庫プリンタ用スプレッドシートID「${STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.HYOGO}」が設定されていません`);
    }
    
    try {
      const printerCounts = getStatusCountsFromSheet(hyogoPrinterId);
      // 各ステータス×種別の数を加算
      for (const status of STATUS_CONSTANTS.STATUSES) {
        for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
          statusCounts[status][deviceType] += printerCounts[status][deviceType] || 0;
        }
      }
      console.log(`兵庫プリンタ用スプレッドシート ${hyogoPrinterId} の処理が完了しました`);
    } catch (error) {
      throw new Error(`兵庫プリンタ用スプレッドシート ${hyogoPrinterId} の処理中にエラーが発生しました: ${error.message}`);
    }
  } else {
    // その他の拠点は単一のスプレッドシート
    if (locationConfig.sourceId) {
      const sheetCounts = getStatusCountsFromSheet(locationConfig.sourceId);
      // 各ステータス×種別の数を加算
      for (const status of STATUS_CONSTANTS.STATUSES) {
        for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
          statusCounts[status][deviceType] += sheetCounts[status][deviceType] || 0;
        }
      }
    }
  }
  
  return statusCounts;
}

/**
 * 指定されたスプレッドシートのmainシートからステータス×種別のマトリックスを集計
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} ステータス×種別の集計結果
 */
function getStatusCountsFromSheet(spreadsheetId) {
  const statusCounts = {};
  
  // ステータス×種別のマトリックスを初期化
  for (const status of STATUS_CONSTANTS.STATUSES) {
    statusCounts[status] = {};
    for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
      statusCounts[status][deviceType] = 0;
    }
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const mainSheet = spreadsheet.getSheetByName(STATUS_CONSTANTS.MAIN_SHEET_NAME);
    
    if (!mainSheet) {
      throw new Error(`スプレッドシート ${spreadsheetId} に「${STATUS_CONSTANTS.MAIN_SHEET_NAME}」シートが見つかりません`);
    }
    
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) {
      console.log(`スプレッドシート ${spreadsheetId} のmainシートにデータがありません`);
      return statusCounts;
    }
    
    // ヘッダー行から「代替機種別」列を検索
    const headerRow = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const typeColumnIndex = headerRow.findIndex(header => header === STATUS_CONSTANTS.TYPE_HEADER_TEXT);
    
    if (typeColumnIndex === -1) {
      throw new Error(`スプレッドシート ${spreadsheetId} に「${STATUS_CONSTANTS.TYPE_HEADER_TEXT}」列が見つかりません。実際のヘッダー: ${headerRow.join(', ')}`);
    }
    
    const typeColumn = typeColumnIndex + 1; // 1ベースのインデックスに変換
    
    // A列（ステータス）と種別列のデータを取得（ヘッダー行を除く）
    const statusData = mainSheet.getRange(2, STATUS_CONSTANTS.STATUS_COLUMN, lastRow - 1, 1).getValues();
    const typeData = mainSheet.getRange(2, typeColumn, lastRow - 1, 1).getValues();
    
    // 各行のステータスと種別をチェックしてカウント
    for (let i = 0; i < statusData.length; i++) {
      const statusValue = statusData[i][0];
      const typeValue = typeData[i][0];
      
      if (statusValue && typeof statusValue === 'string') {
        const trimmedStatus = statusValue.trim();
        
        // 定義されたステータスと一致するかチェック（短縮形も対応）
        for (const status of STATUS_CONSTANTS.STATUSES) {
          const shortStatus = getShortStatusName(status);
          if (trimmedStatus === status || trimmedStatus === shortStatus) {
            // 種別を判定
            let deviceType = 'その他'; // デフォルト値
            
            if (typeValue && typeof typeValue === 'string') {
              const trimmedType = typeValue.trim();
              
              // 定義された種別と一致するかチェック
              for (const type of STATUS_CONSTANTS.DEVICE_TYPES) {
                if (trimmedType === type) {
                  deviceType = type;
                  break;
                }
              }
            }
            
            statusCounts[status][deviceType]++;
            break;
          }
        }
      }
    }
    
    console.log(`スプレッドシート ${spreadsheetId} の集計結果:`, statusCounts);
    
  } catch (error) {
    console.error(`スプレッドシート ${spreadsheetId} の処理中にエラー:`, error.message);
    throw error; // エラーを再スローして上位に伝播
  }
  
  return statusCounts;
}

/**
 * サマリーシートに集計結果を記入
 * @param {Object} aggregatedData - 集計されたデータ
 */
function writeToSummarySheet(aggregatedData) {
  try {
    // 設定から宛先スプレッドシートIDを取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const destinationId = scriptProperties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    if (!destinationId) {
      throw new Error('SPREADSHEET_ID_DESTINATIONが設定されていません');
    }
    
    const destinationSpreadsheet = SpreadsheetApp.openById(destinationId);
    let summarySheet = destinationSpreadsheet.getSheetByName(STATUS_CONSTANTS.SUMMARY_SHEET_NAME);
    
    // サマリーシートが存在しない場合は作成
    if (!summarySheet) {
      summarySheet = createSummarySheet(destinationSpreadsheet);
    }
    
    // 現在の日時を取得
    const currentDateTime = new Date();
    
    // サマリーシートにデータを記入
    updateSummarySheetData(summarySheet, aggregatedData, currentDateTime);
    
    console.log('サマリーシートへの記入が完了しました');
    
  } catch (error) {
    console.error('サマリーシートへの記入中にエラー:', error.message);
    throw error;
  }
}

/**
 * サマリーシートを作成（画像の通りの構造）
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 * @returns {Sheet} 作成されたサマリーシート
 */
function createSummarySheet(spreadsheet) {
  const summarySheet = spreadsheet.insertSheet(STATUS_CONSTANTS.SUMMARY_SHEET_NAME);
  
  // 画像の通りの構造でシートを初期化
  setupSummarySheetStructure(summarySheet);
  
  console.log('サマリーシートを作成しました');
  
  return summarySheet;
}

/**
 * サマリーシートの構造を設定（画像の通り）
 * @param {Sheet} summarySheet - サマリーシート
 */
function setupSummarySheetStructure(summarySheet) {
  let currentRow = 2; // 1行目を空ける
  
  // 各ステータスごとにセクションを作成
  for (let statusIndex = 0; statusIndex < STATUS_CONSTANTS.STATUSES.length; statusIndex++) {
    const status = STATUS_CONSTANTS.STATUSES[statusIndex];
    
    // ステータスタイトル行
    summarySheet.getRange(currentRow, 2).setValue(status);
    summarySheet.getRange(currentRow, 2, 1, 5).merge();
    summarySheet.getRange(currentRow, 2).setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setBackground('#f0f0f0');
    currentRow++;
    
    // 種別ヘッダー行
    const typeHeaders = ['', 'SV', 'CL', 'プリンタ', 'その他'];
    summarySheet.getRange(currentRow, 1, 1, typeHeaders.length).setValues([typeHeaders]);
    summarySheet.getRange(currentRow, 1, 1, typeHeaders.length).setFontWeight('bold');
    summarySheet.getRange(currentRow, 1, 1, typeHeaders.length).setBackground('#e6f3ff');
    currentRow++;
    
    // 拠点行（大阪、神戸、姫路、合計）
    const locations = ['大阪', '神戸', '姫路', '合計'];
    for (const location of locations) {
      summarySheet.getRange(currentRow, 1).setValue(location);
      if (location === '合計') {
        summarySheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
        summarySheet.getRange(currentRow, 1, 1, 5).setBackground('#fff2cc');
      }
      currentRow++;
    }
    
    // セクション間のスペース
    if (statusIndex < STATUS_CONSTANTS.STATUSES.length - 1) {
      currentRow++;
    }
  }
  
  // 列幅を調整
  summarySheet.setColumnWidth(1, 80);  // 拠点列
  summarySheet.setColumnWidth(2, 120); // ステータス/SV列
  summarySheet.setColumnWidth(3, 80);  // CL列
  summarySheet.setColumnWidth(4, 80);  // プリンタ列
  summarySheet.setColumnWidth(5, 80);  // その他列
}

/**
 * サマリーシートのデータを更新（画像の構造に合わせて）
 * @param {Sheet} summarySheet - サマリーシート
 * @param {Object} aggregatedData - 集計データ
 * @param {Date} currentDateTime - 現在日時
 */
function updateSummarySheetData(summarySheet, aggregatedData, currentDateTime) {
  // 既存のシートをクリアして再構築
  summarySheet.clear();
  setupSummarySheetStructure(summarySheet);
  
  let currentRow = 2; // 1行目を空ける
  
  // 各ステータスごとにデータを記入
  for (const status of STATUS_CONSTANTS.STATUSES) {
    // ステータスタイトル行をスキップ
    currentRow++;
    
    // 種別ヘッダー行をスキップ
    currentRow++;
    
    // 拠点データを記入
    const locationKeys = Object.keys(Config.LOCATIONS);
    const locationNames = ['大阪', '神戸', '姫路'];
    
    // 各拠点のデータを記入
    for (let i = 0; i < locationNames.length; i++) {
      const locationKey = locationKeys[i];
      const locationData = aggregatedData.byLocation[locationKey] || {};
      const statusData = locationData[status] || {};
      
      // 各種別の数値を記入
      for (let j = 0; j < STATUS_CONSTANTS.DEVICE_TYPES.length; j++) {
        const deviceType = STATUS_CONSTANTS.DEVICE_TYPES[j];
        const count = statusData[deviceType] || 0;
        summarySheet.getRange(currentRow, j + 2).setValue(count);
      }
      currentRow++;
    }
    
    // 合計行を記入
    const totalStatusData = aggregatedData.total[status] || {};
    for (let j = 0; j < STATUS_CONSTANTS.DEVICE_TYPES.length; j++) {
      const deviceType = STATUS_CONSTANTS.DEVICE_TYPES[j];
      const count = totalStatusData[deviceType] || 0;
      summarySheet.getRange(currentRow, j + 2).setValue(count);
    }
    currentRow++;
    
    // セクション間のスペース
    currentRow++;
  }
  
  console.log('サマリーシートのデータ更新が完了しました');
}

/**
 * 手動実行用：特定拠点のステータス×種別集計テスト
 * @param {string} locationKey - テストする拠点キー
 */
function testLocationStatus(locationKey = 'osaka') {
  console.log(`=== 拠点「${locationKey}」のテスト実行 ===`);
  
  try {
    const result = aggregateLocationStatus(locationKey);
    console.log('集計結果:', result);
    
    // 合計を計算
    let total = 0;
    for (const status of STATUS_CONSTANTS.STATUSES) {
      for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
        total += result[status][deviceType] || 0;
      }
    }
    console.log('合計台数:', total);
    
    // ステータス別の詳細表示
    for (const status of STATUS_CONSTANTS.STATUSES) {
      console.log(`${status}:`);
      for (const deviceType of STATUS_CONSTANTS.DEVICE_TYPES) {
        const count = result[status][deviceType] || 0;
        console.log(`  ${deviceType}: ${count}台`);
      }
    }
    
  } catch (error) {
    console.error('テスト実行中にエラー:', error.message);
  }
}

/**
 * ステータスの短縮形を取得
 * @param {string} fullStatus - 完全なステータス名
 * @returns {string} 短縮されたステータス名
 */
function getShortStatusName(fullStatus) {
  const statusMap = {
    '1.代替機を貸し出しているが修理が完了していつでも返却できる台数': '1.返却可能',
    '2.商談や金額の問題で返却出来ない台数': '2.商談/金銭的な理由により返却不可',
    '3.代替機を貸し出し中だが、お客様より代替機を使うので返却を拒否されている台数': '3.お客様にて返却拒否',
    '4.HW延長保守にて貸し出している台数 ※OS入れ替えやサービス終了を含む': '4.HW延長保守にて貸出'
  };
  
  return statusMap[fullStatus] || fullStatus;
}

/**
 * 手動実行用：mainシートのヘッダー構造を確認
 * @param {string} spreadsheetId - 確認するスプレッドシートID（省略時は大阪の最初のシート）
 */
function debugMainSheetHeaders(spreadsheetId = null) {
  try {
    // スプレッドシートIDが指定されていない場合は大阪の最初のシートを使用
    if (!spreadsheetId) {
      const osakaConfig = Config.getLocationConfig('osaka');
      spreadsheetId = osakaConfig.sourceIds[0];
    }
    
    console.log(`=== スプレッドシート ${spreadsheetId} のヘッダー確認 ===`);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const mainSheet = spreadsheet.getSheetByName(STATUS_CONSTANTS.MAIN_SHEET_NAME);
    
    if (!mainSheet) {
      console.error(`「${STATUS_CONSTANTS.MAIN_SHEET_NAME}」シートが見つかりません`);
      return;
    }
    
    // ヘッダー行を取得
    const headerRow = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    
    console.log('ヘッダー行の内容:');
    headerRow.forEach((header, index) => {
      console.log(`  列${index + 1} (${String.fromCharCode(65 + index)}): "${header}"`);
    });
    
    // 「代替機種別」列の検索結果
    const typeColumnIndex = headerRow.findIndex(header => header === STATUS_CONSTANTS.TYPE_HEADER_TEXT);
    if (typeColumnIndex !== -1) {
      console.log(`「${STATUS_CONSTANTS.TYPE_HEADER_TEXT}」列が見つかりました: 列${typeColumnIndex + 1} (${String.fromCharCode(65 + typeColumnIndex)})`);
    } else {
      console.log(`「${STATUS_CONSTANTS.TYPE_HEADER_TEXT}」列が見つかりませんでした`);
    }
    
  } catch (error) {
    console.error('ヘッダー確認中にエラー:', error.message);
  }
}

/**
 * 手動実行用：プリンタ用スプレッドシートIDの設定状況を確認
 */
function debugPrinterSpreadsheetIds() {
  console.log('=== プリンタ用スプレッドシートID設定確認 ===');
  
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // 大阪プリンタのスプレッドシートID確認
    const osakaPrinterId = scriptProperties.getProperty(STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.OSAKA);
    console.log(`大阪プリンタ (${STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.OSAKA}):`, osakaPrinterId || '未設定');
    
    if (osakaPrinterId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(osakaPrinterId);
        const mainSheet = spreadsheet.getSheetByName(STATUS_CONSTANTS.MAIN_SHEET_NAME);
        if (mainSheet) {
          const lastRow = mainSheet.getLastRow();
          console.log(`  - アクセス可能、データ行数: ${lastRow - 1}行`);
        } else {
          console.log(`  - アクセス可能だが「${STATUS_CONSTANTS.MAIN_SHEET_NAME}」シートが見つかりません`);
        }
      } catch (error) {
        console.error(`  - アクセスエラー: ${error.message}`);
      }
    }
    
    // 兵庫プリンタのスプレッドシートID確認
    const hyogoPrinterId = scriptProperties.getProperty(STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.HYOGO);
    console.log(`兵庫プリンタ (${STATUS_CONSTANTS.PRINTER_SPREADSHEET_IDS.HYOGO}):`, hyogoPrinterId || '未設定');
    
    if (hyogoPrinterId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(hyogoPrinterId);
        const mainSheet = spreadsheet.getSheetByName(STATUS_CONSTANTS.MAIN_SHEET_NAME);
        if (mainSheet) {
          const lastRow = mainSheet.getLastRow();
          console.log(`  - アクセス可能、データ行数: ${lastRow - 1}行`);
        } else {
          console.log(`  - アクセス可能だが「${STATUS_CONSTANTS.MAIN_SHEET_NAME}」シートが見つかりません`);
        }
      } catch (error) {
        console.error(`  - アクセスエラー: ${error.message}`);
      }
    }
    
    // 設定されているすべてのスクリプトプロパティを表示（デバッグ用）
    console.log('\n=== 全スクリプトプロパティ ===');
    const allProperties = scriptProperties.getProperties();
    Object.keys(allProperties).forEach(key => {
      console.log(`${key}: ${allProperties[key]}`);
    });
    
  } catch (error) {
    console.error('プリンタスプレッドシートID確認中にエラー:', error.message);
  }
}

/**
 * 手動実行用：プリンタ用スプレッドシートの接続テスト
 */
function testPrinterSpreadsheetConnection() {
  console.log('=== プリンタ用スプレッドシート接続テスト ===');
  
  try {
    // 大阪プリンタのテスト
    console.log('大阪プリンタのテスト開始...');
    try {
      const osakaResult = aggregateLocationStatus('osaka');
      console.log('大阪プリンタ接続テスト: 成功');
      console.log('大阪集計結果:', osakaResult);
    } catch (error) {
      console.error('大阪プリンタ接続テスト: 失敗 -', error.message);
    }
    
    // 神戸（兵庫プリンタ）のテスト
    console.log('\n神戸（兵庫プリンタ）のテスト開始...');
    try {
      const kobeResult = aggregateLocationStatus('kobe');
      console.log('兵庫プリンタ接続テスト: 成功');
      console.log('神戸集計結果:', kobeResult);
    } catch (error) {
      console.error('兵庫プリンタ接続テスト: 失敗 -', error.message);
    }
    
    // 姫路のテスト（プリンタなし）
    console.log('\n姫路のテスト開始...');
    try {
      const himejiResult = aggregateLocationStatus('himeji');
      console.log('姫路接続テスト: 成功');
      console.log('姫路集計結果:', himejiResult);
    } catch (error) {
      console.error('姫路接続テスト: 失敗 -', error.message);
    }
    
  } catch (error) {
    console.error('プリンタスプレッドシート接続テスト中にエラー:', error.message);
  }
} 