# 転記処理スクリプト　ドキュメント


## 概要

本スクリプトは、Google Apps Script (GAS) を用いて、指定した「転記元」スプレッドシート（以下「転記元シート」）からデータを取得し、「転記先」スプレッドシート（以下「転記先シート」）へデータを転記する処理を自動化するものです。  
スクリプトは転記処理に加え、データの整合性チェック、ログの出力、メールによる通知（成功・エラー報告）も行います。



## ヒューマンエラーが介入しやすいポイントまとめ

1. **スクリプトプロパティの未設定・誤設定**  
   - `SPREADSHEET_ID_SOURCE`, `SPREADSHEET_ID_DESTINATION`, `LOCATION_MAPPING`, `NOTIFICATION_EMAIL`, `ERROR_NOTIFICATION_EMAIL`  
   これらが未設定、もしくは別環境や別スプレッドシートのIDを誤って入力すると、シート取得エラーや通知メール未送信の原因となります。

2. **ヘッダー行の変更・不一致**  
   - 転記元シート(`main`シート)のヘッダー（B列「資産管理番号」やD列「ステータス」など）が期待値から変更されると、スクリプトがエラーを返します。  
   - ユーザーが意図せずカラム名を変更したり、列の順序を変えたりした場合に、エラーが発生します。

3. **転記先シート名の不整合**  
   - `LOCATION_MAPPING`に設定したシート名が転記先スプレッドシート上に存在しない場合、処理がエラーとなります。  
   - シート名を変更・削除・リネームした際にプロパティを更新し忘れるミスが起きやすいです。

4. **資産管理番号の形式変更**  
   - 転記元データが「AS」始まりであることを前提としているため、資産管理番号のフォーマットをユーザーが変更した場合、転記対象から漏れる可能性があります。

5. **日付データの不備**  
   - 転記元シートに想定外の形式（テキスト形式や不正な日付文字列）が入力されると、日付整形処理で問題が生じる可能性があります。

6. **通知用メールアドレスの間違い**  
   - `NOTIFICATION_EMAIL`および`ERROR_NOTIFICATION_EMAIL`に誤ったメールアドレスを設定すると、ログやエラー通知が届かず、問題の発見が遅れることがあります。

---


## 前提条件・設定項目

スクリプトは以下のスクリプトプロパティを使用します。  
スクリプトプロパティは、`プロジェクトの設定` -> `スクリプトプロパティ` より設定します。

| プロパティキー               | 説明                                                                 | 必須 | 例                     |
|------------------------------|----------------------------------------------------------------------|------|------------------------|
| SPREADSHEET_ID_SOURCE        | 転記元スプレッドシートのID                                           | 必須 | `1xyz...sourceID`      |
| SPREADSHEET_ID_DESTINATION   | 転記先スプレッドシートのID                                           | 必須 | `1abc...destinationID` |
| LOCATION_MAPPING              | 転記先スプレッドシート内でデータを反映させるシート名。拠点名などを想定 | 必須 | `Tokyo`                |
| NOTIFICATION_EMAIL           | 転記処理成功時などのログ通知メール送信先アドレス                      | 必須 | `example@example.com`  |
| ERROR_NOTIFICATION_EMAIL      | エラー発生時のエラーメール通知先アドレス                              | 必須 | `error@example.com`    |

※ 上記のスクリプトプロパティが正しく設定されていないとエラーが発生します。

## スクリプトの動作フロー

1. **転記元シートのデータ取得 (listColumnData関数)**  
   - 転記元スプレッドシート（IDは`SPREADSHEET_ID_SOURCE`）にある「main」シートを開く。  
   - ヘッダー行を検証し、想定通りのカラム名であることをチェック。  
   - 「B列」に`"AS"`で始まる資産管理番号を持つ行を抽出し、その行データをリストとして取得する。
   - この際、A列は除外され、日付がある場合は`YYYY/MM/DD`形式に整形。

2. **データの検証およびログ出力**  
   - 転記元データが空の場合は処理をスキップし、ログを残した上でメール通知。  
   - 転記元データが存在すれば、取得した資産管理番号の一覧をログ出力。

3. **転記先シートへの転記 (transferToSpreadsheetDestination関数)**  
   - スクリプトプロパティから`SPREADSHEET_ID_DESTINATION`および`LOCATION_MAPPING`を取得し、該当シートを特定。  
   - 転記先シートのB列(4行目以降)を読み込み、資産管理番号の一覧をMapに格納して高速検索を可能にする。  
   - 転記元データの各行について、転記先シートで該当する資産管理番号の行を特定し、以下の列に値を設定またはクリアする：  
     - A列: `rowData[6]`  
     - K列: `rowData[2]`  
     - 代替貸出の場合(`rowData[2] === "代替貸出"`)のみ、L列、M列、O列を設定し、N列に「有」を設定。それ以外はこれらの列をクリア。  
   - 転記先に存在しない資産管理番号は「未処理行」としてログに残す。

4. **差分ログの出力 (logDifferences関数)**  
   - 転記元データと転記先シートの資産管理番号の差分をログ出力（転記元にあるが転記先にないもの、転記先にあるが転記元にないものを表示）。

5. **処理結果の通知 (sendLogNotification関数, sendErrorNotification関数)**  
   - 処理成功時または未処理行がある場合でも、ログ全体をメール通知する。  
   - エラー発生時はエラーログをエラー通知先アドレスに送信する。

## 関数一覧

### transferDataMain

**役割**: 全体的な転記処理フローを統括するメイン関数。

**処理内容**:  
- `listColumnData()`で転記元データ取得  
- データ存在チェック  
- ログ出力  
- `transferToSpreadsheetDestination()`で転記先へ反映  
- `logDifferences()`で差異ログ出力  
- すべての処理完了後、`sendLogNotification()`でログメール送信  
- エラー発生時は`sendErrorNotification(error)`でエラーメール送信

### listColumnData

**役割**: 転記元シート(main)のデータを取得し、資産管理番号(`"AS"始まり`)の行を抽出して返す。

**処理内容**:  
- スクリプトプロパティから`SPREADSHEET_ID_SOURCE`を取得し開く。  
- `main`シートを取得、ヘッダー検証(`checkHeader()`呼び出し)。  
- B列が`"AS"`で始まる行のみ抽出。A列を除いたデータを戻り値として返す。  
- 日付型の場合は`formatValue()`で`YYYY/MM/DD`形式の文字列に整形。

**戻り値**: `rows` (配列: 各行は列ごとの値の配列)

### checkHeader

**役割**: `main`シートのヘッダー行が期待されたカラム名であるかを検証。

**処理内容**:  
- 期待されるヘッダー配列と実際のヘッダーを比較し、不一致があればエラーをスロー。

### formatValue

**役割**: 値が日付の場合、`YYYY/MM/DD`形式に整形し、文字列として返す。

### transferToSpreadsheetDestination

**役割**: 転記元データ(`rows`)を転記先シートに反映する。

**処理内容**:  
- `SPREADSHEET_ID_DESTINATION`および`LOCATION_MAPPING`を使用して転記先シートを特定。  
- 転記先シートのB列(4行目以降)を取得し、資産管理番号をキーとしたMapを作成。  
- 転記元データを1行ずつ処理し、対応する資産管理番号行が見つかれば特定列に値を設定。また、「代替貸出」の場合は特定列に追加情報を設定。  
- 該当行が見つからなかったものは「未処理行」として`unprocessedRows`に格納し返す。

**戻り値**: `unprocessedRows` (処理されなかった行データの配列)

### logDifferences

**役割**: 転記元と転記先で食い違っている資産管理番号をログ出力する。

**処理内容**:  
- 転記先シートの資産管理番号一覧を取得。  
- 転記元にあるが転記先にないもの、転記先にあるが転記元にないものをそれぞれログに記録。

### sendLogNotification

**役割**: 処理結果のログをメールで通知する。

**処理内容**:  
- `NOTIFICATION_EMAIL` と `LOCATION_MAPPING` を使用し、処理結果ログをメール送信。

### sendErrorNotification

**役割**: エラー発生時にエラー内容をメールで通知する。

**処理内容**:  
- `ERROR_NOTIFICATION_EMAIL` と `LOCATION_MAPPING` を使用し、エラー内容とスタックトレースをメール送信。

## 想定される実行方法

- Google Apps Scriptエディタ上で `transferDataMain()` 関数を実行するか、トリガーを設定して自動的に定期実行することを想定。  
- トリガー設定例：GASプロジェクトで「トリガー」から`transferDataMain`を毎日、または毎時実行するように設定。

## エラーハンドリング

- スクリプトプロパティ未設定、ヘッダー不一致、シート未存在などの環境エラー時はエラーをスローし、`transferDataMain()`から`catch`されて`sendErrorNotification()`で通知。
- 未処理行があった場合はログに記録されるが、スクリプトとしてはエラーとせずメールでログ通知される。



# mainシート列順整理スクリプト ドキュメント

## ヒューマンエラーが介入しそうなポイント

1. **スクリプトプロパティ設定ミス**  
   - `PREFIX`・`SPREADSHEET_ID`が未設定、または誤ったIDや値を設定している場合、意図したシートを取得できない・対象シートが見つからないなどのエラーが発生します。

2. **シート内のヘッダー名変更**  
   - 対象シートの列ヘッダーが想定通りでない場合は自動的に列移動しますが、想定外のヘッダー追加や、必要なヘッダー削除・変更があった場合、正しい列順に戻せない可能性があります。

3. **ログシートの命名衝突・手動変更**  
   - `organize_dataSheets_log`という名前のシートを本スクリプトで作成・使用します。ユーザーが手動でこのシートの名前を変更・削除すると、正しくログが記録されなくなります。

---

## 概要

本スクリプトは、指定したスプレッドシート内で、特定のプレフィックス(`PREFIX`)を含むシートを取得し、それらシートの列順を特定の標準形に揃えます。  
標準の列順は以下の通りです:

```
["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり証No.", "備考"]
```

この処理を行うことで、シート間のデータ整合性を保ち、後続処理やデータ解析がスムーズになります。また、処理結果は`organize_dataSheets_log`シートに記録され、トラブルシューティングや監査に利用できます。

## 前提条件・設定項目

スクリプトは以下のスクリプトプロパティを用いて動作します。

| プロパティキー | 説明                                  | 必須 | 例                     |
|----------------|---------------------------------------|------|------------------------|
| PREFIX         | 対象となるシート名に含まれる接頭辞     | 必須 | "Data_"                |
| SPREADSHEET_ID | 列順を整理したいスプレッドシートのID   | 必須 | `1xyz...spreadsheetID` |

これらの値が正しく設定されていない場合、処理対象シートを正しく取得できず、スクリプトは期待通りに動作しません。

## スクリプトの動作フロー

1. **スクリプトプロパティの読み込み**  
   - `PREFIX`と`SPREADSHEET_ID`を取得し、指定がなければログを出して終了。

2. **対象スプレッドシートおよびシート取得**  
   - `SPREADSHEET_ID`を元にスプレッドシートを開き、`PREFIX`を含む名称のシートを抽出。

3. **各シートに対する列順チェック**  
   - シートの1行目からヘッダー行を取得。  
   - 事前定義した`correctHeaders`（正しい列順）と比較し、`arraysEqual()`関数で完全一致するか確認。

4. **列順修正**  
   - 一致しない場合、`rearrangeColumns()`関数で列を移動し、定義された正しい順序に揃える。  
   - 移動操作があった場合は詳細をログシート(`organize_dataSheets_log`)に記録。

5. **ログ出力**  
   - `ensureLogSheet()`関数でログシートがなければ作成し、`logAction()`関数を使って各シートごとの処理結果を記録。

## 主な関数の説明

### organizeDataSheets

**役割**: 全体的な列順整理処理のメインエントリーポイント。

**処理内容**:  
- スクリプトプロパティから`PREFIX`と`SPREADSHEET_ID`を取得。  
- 指定スプレッドシートのうち、シート名に`PREFIX`を含むシートを抽出。  
- 正しいヘッダー順序(`correctHeaders`)とシートの現状ヘッダーを比較し、不一致なら`rearrangeColumns()`で並び替えを実行。  
- 結果をログシートに記録。

### rearrangeColumns

**役割**: 列順が不正なシートの列を正しい順序に移動する。

**処理内容**:  
- 期待するヘッダー順序(`correctHeaders`)を元に、現在のヘッダー(`headers`)との位置差を確認。  
- `moveColumns()`メソッドを使用し、必要に応じて列を正しい位置に移動。  
- 操作結果（移動内容）をログ記録。

**注意点**:  
- 列移動処理中に列のインデックスが変化するため、処理途中でヘッダー情報(`headers`)を再取得しながら進めます。  
- 列移動に失敗した場合は、エラーメッセージをログシートに記録。

### ensureLogSheet

**役割**: 処理ログを記録するシートが存在するかチェックし、なければ作成する。

**処理内容**:  
- `logSheetName`として定義されたシート名(`organize_dataSheets_log`)を元に、ログシートの存在確認。  
- 存在しなければ新規作成し、ヘッダー行「処理日時」「シート名」「アクション」を設定。  
- ログシートの`Sheet`オブジェクトを返す。

### logAction

**役割**: ログシートに処理結果やエラー内容を記録する。

**処理内容**:  
- 日時(`new Date()`)・シート名・アクション内容を1行として追記。

### arraysEqual

**役割**: 配列同士が完全一致するかを判定するユーティリティ関数。

**処理内容**:  
- 長さが同じかチェックし、要素が1つでも異なれば`false`を返す。  
- 完全一致の場合は`true`を返す。

## 想定される実行方法

- Google Apps Scriptエディタ上で `organizeDataSheets()` 関数を実行するか、定期的な実行トリガーを設定することで、定期的にシート列順を整理できます。

## エラーハンドリング

- `PREFIX`や`SPREADSHEET_ID`が未設定の場合、ログ出力にとどめ、処理は行われません。  
- 列移動中にエラーが発生した場合は、その内容をログシートに記録します。  
- 必要に応じて、エラー通知メールなどを統合することも可能です（本スクリプトには標準機能として実装されていません）。



# 代替機管のデータシート用　列順チェック＆修正スクリプト ドキュメント

## ヒューマンエラーが介入しそうなポイント

1. **`SPREADSHEET_ID`の設定ミス**  
   - スクリプトプロパティに`SPREADSHEET_ID`が設定されていない、もしくは誤ったIDを設定していると処理そのものが開始できません。
   
2. **列ヘッダー名の変更・削除**  
   - コード中で定義されている`correctColumnOrder`の列名が、対象シートで改変されると、列順整頓が想定通りに動作せず、データの不整合を引き起こします。  
   - 列名は維持するか、変更した場合はコード側も更新する必要があります。

3. **ログシート`header_check_on_main_sheet_log`の手動削除・編集**  
   - ログシートが存在しなければ再作成しますが、手動で名前変更や削除を行うとログの蓄積・ローテーションが正常に行えない場合があります。

4. **日付データを含むログローテーション処理**  
   - ログローテーションでは1年以上前のログを削除します。日付フォーマットやシートタイムゾーンが期待通りでない場合、意図せずデータが保持/削除される可能性があります。

---

## 概要

このスクリプトは、指定したスプレッドシート（`SPREADSHEET_ID`）のアクティブシート（ユーザーが最後に表示していたシート）に対して、事前定義した「正しい列順」に整える処理を実行します。

また、  
- 現在の列順をチェックして、必要に応じて列を正しい並び順に修正します。  
- 列順が変更された場合は、どの列がどの位置からどこへ移動したかをログシートに記録します。  
- ログシートは1年以上前のデータを削除し、古いログを自動的にローテーションします。

## 前提条件・設定項目

スクリプトは以下のスクリプトプロパティを用います。

| プロパティキー   | 説明                                                  | 必須 | 例                    |
|------------------|-------------------------------------------------------|------|-----------------------|
| SPREADSHEET_ID   | 列順を整えたいスプレッドシートのID                     | 必須 | `1xyz...spreadsheetID`|

これらを正しく設定しないと、処理が開始できずエラーとなります。

## スクリプトの処理フロー

1. **スクリプトプロパティから`SPREADSHEET_ID`取得**  
   - 未設定または不正な場合はエラー発生。

2. **ターゲットスプレッドシート・アクティブシート取得**  
   - `SPREADSHEET_ID`を使用してスプレッドシートを開き、`getActiveSheet()`でカレントシートを取得。

3. **ログシートの準備**  
   - `header_check_on_main_sheet_log` という名前のログシートを取得または作成。  
   - ログシートのローテーション処理(`rotateLog`)を行い、1年以上前のログを削除。

4. **列順チェック**  
   - アクティブシートのヘッダー行を取得し、定義済みの正しい列順(`correctColumnOrder`)と比較。  
   - 不一致の場合、ログに記録し、`rearrangeColumns()`で列順を修正し、変更内容をログ出力。

5. **一致の場合**  
   - 正しい列順であれば、その旨をログに記録して終了。

## 主な関数説明

### checkAndFixColumnOrder_main

**役割**: 一連の列チェックおよび修正処理のメイン関数。

**処理概要**:  
- スクリプトプロパティから`SPREADSHEET_ID`取得・シートオープン  
- アクティブシートとログシートの用意  
- 列順比較後、一致しなければ`rearrangeColumns()`で並び替え、ログ出力  
- 一致すればログに「正しい」と記録

### arraysEqual

**役割**: 2つの配列が同一であるか厳密にチェックするユーティリティ関数。

### rearrangeColumns

**役割**: 現在の列順を正しい順序に合わせて再配置する。

**処理内容**:  
- 現在のヘッダーと正しい列順を比較してインデックスマップを作成。  
- データ本体も正しい順序に並び替え、新たなシート構成をセットし直す。  
- 列移動後はログを記録。

### logMessage

**役割**: ログシートに日時付きでメッセージを追記する。

### logColumnChanges

**役割**: 列変更前後のインデックスをログに記録する。  
- 現在のヘッダー配列をもとに、正しい列順のインデックスへどう移動したかを出力。

### createLogSheet

**役割**: 存在しない場合、ログシート`header_check_on_main_sheet_log`を作成。  
- ログのヘッダー行（タイムスタンプ、メッセージ）を設定。

### rotateLog

**役割**: ログシートから1年以上前のログを削除し、古いログを定期的にクリーンアップする。  
- シート上の全データを取得し、日付が1年以上前の行を除去した上でリセット。

## 想定される実行方法

- 手動でGASエディタから `checkAndFixColumnOrder_main()` を実行する。  
- または、タイマー・トリガー設定(Google Apps Scriptのトリガー機能)により、定期的に実行し、列順が乱れていないかを定期的に監視することも可能。

## エラーハンドリング

- `SPREADSHEET_ID`未設定時はスクリプト冒頭でエラーを投げる。  
- 列順整頓中に問題が発生した場合（例えば列名が存在しない等）には、エラーがスローされ、スクリプトは停止する可能性がある。必要に応じて`try-catch`でエラー発生時にメール通知などの機能を追加するとよい。



# フォーム作成スクリプト ドキュメント

## ヒューマンエラーが介入しやすいポイント

1. **「フォーム作成」シートの存在確認・入力値**  
   - 「フォーム作成」というシートが存在しない場合、エラーが発生して処理中断します。  
   - シート上のセル(C2, C3)にフォームタイトルやフォルダIDを正しく記載する必要があります。  
   - FORM_TITLE(C2セル)・FOLDER_ID(C3セル)が空だったり、誤ったIDを入力した場合、フォーム作成が正しく行われません。

2. **既存フォルダとの重複タイトル**  
   - 同じフォルダ内に同名ファイルが存在する場合、処理が中断されます。ユーザーはフォームタイトルを変更する必要があります。

3. **スプレッドシートとの連携**  
   - `SPREADSHEET_ID`は自動取得しますが、スクリプトが実行されるスプレッドシートが回答記録先として想定しているため、環境を移動する際や複製する際には注意が必要です。  

4. **必須回答と入力必須項目**  
   - 質問項目の必須設定や初期値が妥当でない場合、フォームが意図した通りに動作しない可能性があります。質問項目の増減・修正の際はコード側の反映が必須。

5. **ログシート、QRコードシートの手動編集**  
   - `createForm_script_log`および `QRコード` シートが手動で削除・改名されると、ログやQRコード追加処理に支障が出ます。

---

## 概要

本スクリプトは、Google スプレッドシートとGoogle フォームを連携し、以下のことを自動化します：

- スプレッドシート内の「フォーム作成」シートで設定されたフォームタイトルとフォルダIDを読み込み  
- 指定フォルダ配下に新たなフォームを作成し、回答をこのスプレッドシートにリンク  
- フォームには、様々な状況(貸出・回収・修理など)に応じた質問項目とセクション遷移を自動生成  
- 作成後、フォームURLをスプレッドシートの`createForm_script_log`シートに記録し、QRコードを`QRコード`シートに生成・挿入

## 前提条件・設定項目

- 「フォーム作成」という名前のシートが存在し、  
  - C2セル：フォームタイトル  
  - C3セル：フォームを配置するドライブフォルダのID  
  を設定していること。  
- スクリプトは当該スプレッドシート上で動作し、`onButtonClick()`をトリガーとしてフォームを生成する想定。

## 処理フロー

1. **onButtonClick関数**:  
   - 「フォーム作成」シートの存在確認  
   - C2セルからフォームタイトル、C3セルからフォルダID取得  
   - ユーザーに対しポップアップで作成確認（YES/NO）  
   - YESの場合、`createGoogleFormWithEmailCollection()`を呼び出しフォーム作成処理へ。

2. **createGoogleFormWithEmailCollection関数**:  
   - フォームタイトル・フォルダID・スプレッドシートIDを確認  
   - フォルダ内の重複ファイル名チェック  
   - `createForm()`でフォーム作成  
   - `linkResponseSpreadsheet()`でフォームの回答を当該スプレッドシートにリンク  
   - `addQuestions()`で必要な質問とセクションを追加  
   - `setSectionNavigation()`でステータスに応じたページ遷移ロジックを設定  
   - フォームの編集URL、送信URLをログに記録  
   - `generateQrCode()`でフォーム送信URLをQRコード化し、`QRコード`シートに挿入。

3. **addQuestions関数群**:  
   - `addStatusQuestion()`: ステータス選択の複数選択項目追加  
   - `addLendSectionQuestions()`: 貸出セクション作成  
   - `addReturnSectionQuestions()`: 回収セクション作成  
   - `addRepairSectionQuestions()`: 修理・設定セクション作成  
   - `addBorrowSectionQuestions()`: 社外持ち出しセクション作成  
   - `addConfirmationSectionQuestions()`: 最終確認用テキスト質問追加

4. **setSectionNavigation関数**:  
   - ステータス選択肢に応じて各セクションへ遷移するロジック設定  
   - 回収方法によってフォームを最初からやり直す(社外持ち出し選択時)といった特殊遷移も設定。

5. **ログ処理**:  
   - `logToSheet()`で全ての重要なイベント（作成状況、URL、エラーなど）を`createForm_script_log`シートに記録。

6. **QRコード生成**:  
   - `generateQrCode()`でフォーム送信URLからQRコードを生成  
   - `QRコード`シートに画像挿入・日付・管理番号を記録。

## 主な関数詳細

### onButtonClick
**役割**: ボタンクリック時のエントリーポイント。シート存在確認、ユーザー確認ダイアログ表示、フォーム作成処理開始。

### createGoogleFormWithEmailCollection
**役割**: 必要なID類のチェック、フォーム新規作成、回答先連携、質問追加、URLログ出力、QRコード生成までを統括。

### createForm
**役割**: 指定フォルダにフォーム作成し、メールアドレス収集・説明設定・フォルダ所属設定。

### linkResponseSpreadsheet
**役割**: 作成したフォームを当該スプレッドシートに紐付け、回答結果を書き込み可能な状態に。

### addQuestionsおよび関連関数
**役割**: 各種セクション(貸出、回収、修理、社外持ち出し、確認)の質問項目を追加。

### setSectionNavigation
**役割**: ステータスや回収方法などの選択肢に応じたフォームページ遷移を設定。

### logToSheet
**役割**: ログシート(`createForm_script_log`)がなければ作成し、実行時メッセージを時刻付きで記録。

### generateQrCode
**役割**: QuickChart APIを用いてフォームURLのQRコードを生成し、`QRコード`シートに画像挿入・管理番号と日付を記録。

## 想定される実行方法

- Google スプレッドシート上のボタンに `onButtonClick` 関数を割り当てることで、ワンクリックでフォーム作成処理を開始できます。  
- ボタンがない場合でも、GASエディタから手動で `onButtonClick()` を実行可能。

## エラーハンドリング

- 「フォーム作成」シートが存在しない、または必要セルが空の場合、アラート表示やログ出力でユーザーに通知。  
- 同名のファイルがフォルダ内に存在する場合は処理中断し、ログに記録。  
- 予期せぬエラーが発生する場合、`logToSheet()`でログを残し、状況把握に役立ちます。  
- 必要に応じて、エラーメール通知機能などを追加可能です。

---



