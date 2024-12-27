# 代替管理シート機能 スクリプトドキュメント

---

# 1. 転記スクリプト

## スクリプト名・概要
**転記スクリプト**  
スクリプトプロパティに設定された「SPREADSHEET_ID_SOURCE（転記元）」「SPREADSHEET_ID_DESTINATION（転記先）」「LOCATION_MAPPING（拠点）」などの情報をもとに、転記元シート(main)から指定拠点のシートへ必要なデータを転記・ログ出力・エラーメール送信を行う処理をまとめたスクリプトです。  

---

## 使用する主なスクリプトプロパティ

| プロパティ名                  | 内容                                                 |
|------------------------------|------------------------------------------------------|
| SPREADSHEET_ID_SOURCE        | 転記元のスプレッドシートID                           |
| SPREADSHEET_ID_DESTINATION   | 転記先のスプレッドシートID                           |
| LOCATION_MAPPING             | 転記先で使用するシート名（拠点名などを想定）         |
| ERROR_NOTIFICATION_EMAIL     | エラー発生時に通知するメールアドレス                 |

---

## 関数一覧

### 1. `transferDataMain()`
転記のメイン処理を行う関数です。以下のような流れで処理します。

1. **転記元シートのデータを取得**  
   `listColumnData()`を呼び出して、B列が「AS」から始まる行(資産管理番号が「AS～」)のみを抽出。

2. **データがない場合の処理**  
   取得データが空であれば、ログを出力して終了。

3. **転記を実行**  
   `transferToSpreadsheetDestination()`で取得したデータを転記先へ書き込み。未処理行（資産管理番号が転記先に存在しない行）があればログを残す。

4. **差分ログの出力**  
   `logDifferences()`を用いて、転記元にはあるが転記先シートにはない資産管理番号、またその逆の情報をログ出力。

5. **結果をログシートに出力**  
   `writeLogToSheet()`で、処理内容や未処理件数等をログシート(`transferDataMain_log`)へ追記。

6. **エラー発生時の処理**  
   例外が発生した場合は、ログを出力しつつ `sendErrorNotification()` でメール通知を行う。

---

### 2. `listColumnData()`
転記元スプレッドシートの「main」シートからデータを取得して、以下の条件でフィルタリング・整形します。

- B列（資産管理番号）が「AS」で始まる行のみ抽出
- A列を除外し、**B列以降**のデータを取得  
- 日付型のセルがあれば `formatValue()` で `YYYY/MM/DD`形式に変換

さらに処理中に `checkHeader()` を呼び出し、ヘッダー行に想定するラベルがあるかチェックを行います。

---

### 3. `checkHeader(headerRow)`
転記元シートのヘッダーが想定通りかを検証します。  
想定ヘッダー：`['', '資産管理番号', '', 'ステータス', '顧客名', '', '日付', '担当者', '備考', 'お預かり証No.']`  
もし一致しない列があればエラーを投げ、処理を中断します。

---

### 4. `formatValue(value)`
- 日付型を判定し、`YYYY/MM/DD` の文字列に整形
- 文字列や数値などは文字列化して返す

---

### 5. `transferToSpreadsheetDestination(rows)`
転記先のスプレッドシートIDを取得後、拠点(`LOCATION_MAPPING`)で指定されたシートに対してデータを書き込みます。

1. **転記先のシート名を取得**  
   - スクリプトプロパティ `LOCATION_MAPPING` からシート名を判定
   - 存在しない場合はエラー

2. **転記先のB列をMap化**  
   - 4行目から最終行までの資産管理番号を読み込み  
   - `Map` を使って `資産管理番号 -> 行番号` に紐づけ、高速検索を可能にする

3. **転記**  
   - `rows` (転記元データ) の各資産管理番号が Map に存在するか確認  
   - 転記に該当する行が見つかれば、`destinationSheet.getRange(….)` で必要箇所に書き込み  
     - 代替貸出かどうかで列の書き込み／消去制御
   - 見つからない場合は「未処理行」として配列に格納し、ログ出力・戻り値として返す

---

### 6. `logDifferences(rows)`
- 転記元にあるが、転記先に存在しない資産管理番号  
- 転記先にあるが、転記元に存在しない資産管理番号  
をそれぞれログ出力する。

---

### 7. `writeLogToSheet(logMessage)`
- 転記元の `SPREADSHEET_ID_SOURCE` を開き、`transferDataMain_log` シートにログを残す。  
- シートが無ければ作成し、 `[日時, ログメッセージ]` 形式でAppend。  
- 同時に `rotateLog(logSheet)` を呼び出し、3年より前のログを削除する。

---

### 8. `rotateLog(sheet)`
- ログの世代管理を行う関数。  
- 現在の日付から `maxYears = 3` 年より古いログは削除する。

---

### 9. `sendErrorNotification(error)`
- スクリプトプロパティ `ERROR_NOTIFICATION_EMAIL` 宛にメール通知を送信。  
- 件名に拠点(`LOCATION_MAPPING`)を含める

---

## 想定される実行タイミング
- 手動実行または時間主導型トリガーで定期的に呼び出される想定

---

## 注意点
- 転記元/転記先のシート構成が変更された場合、ヘッダー名や列の位置、Map化のロジックなどに影響が出る可能性があるため要注意です。  
- スクリプトプロパティが正しく設定されていないと正常に動作しないので事前設定が必須です。

---

# 2. データシートチェック

## スクリプト名・概要
**データシートチェック**  
ユーザーが開いているスプレッドシート内の「特定のシート」(PREFIXで指定)ごとに、カラム順を特定の順番に整列し直す処理を行います。  

---

## 関数一覧

### 1. `onOpen()`
- スプレッドシートを開いた際に、メニューに「【メニュー】→ 実行する」が追加されます。  
- ここから `main()` または `dataSheetColumnCheckMain()` が実行できるようにしている形です。

### 2. `dataSheetColumnCheckMain()`
1. **PREFIXを取得**  
   - スクリプトプロパティから `PREFIX` を取得し、カンマ区切り配列に変換

2. **対象シートの判定**  
   - 開いているスプレッドシート内にあるすべてのシートに対して `PREFIX` のいずれかを含むシート名なら処理を実施

3. **カラムの順序揃え**  
   - シートの1行目（ヘッダ）を読み込み、`["タイムスタンプ", "メールアドレス", "ステータス", "顧客名", "預かり機の製造番号", "備考", "お預かり証No."]` が左から順に来るように `sheet.moveColumns()` を実行

4. **結果**  
   - 指定したヘッダー項目の並びを自動でソートし、指定どおりの順番に整列する

---

## 想定される実行タイミング
- メニューから選択して手動実行。

---

## 注意点
- `PREFIX` に指定されるシート名の一部を変更した場合、処理対象が変わる可能性がある  
- カラムの移動を伴うため、想定通りの列見出し名・挿入位置が存在しないと正常に動作しない

---

# 3. フォーム作成スクリプト

## スクリプト名・概要
**フォーム作成スクリプト**  
スプレッドシート上でボタンを押して実行し、Google フォームを自動作成して指定シートにリンクしたり、フォームURLのQRコードを自動生成するものです。

---

## 関数一覧

### 1. `onButtonClick()`
- スプレッドシートの「フォーム作成」という名前のシートを取得  
- シート上の `C2` セルをフォームタイトル、`C3` セルをフォルダIDとして読み込む  
- ポップアップでユーザーに確認のうえ、`createGoogleFormWithEmailCollection()` を呼び出してフォームを作成

---

### 2. `createGoogleFormWithEmailCollection(formTitle, folderId, spreadsheetId)`
1. **必須情報のチェック**  
   - `formTitle`, `folderId`, `spreadsheetId` のいずれかが無いとログを残して終了

2. **重複ファイル名チェック**  
   - 指定フォルダ(`folderId`)に同名のファイルが無いか確認

3. **フォーム作成**  
   - `createForm()` でフォルダ直下にフォームを生成（デフォルトはマイドライブに作成されるので、移動し、元ファイルをマイドライブから削除）

4. **スプレッドシートリンク**  
   - `linkResponseSpreadsheet()` で既存シートにフォーム回答をリンク

5. **フォームへの質問追加**  
   - `addQuestions()` を呼び出し、複数のセクション(ステータス、貸出、回収、修理、社外持ち出し、確認)を追加

6. **セクションナビゲーションの設定**  
   - `setSectionNavigation()` で、ステータス選択肢ごとにフォームページブレークの遷移先を指定

7. **QRコード生成**  
   - フォームの送信URLをもとに `generateQrCode()` を呼び出してQRコード画像を生成し、シート「QRコード」に登録

---

### 3. `createForm(formTitle, folderId)`
- 新しいGoogleフォームを作成し、指定フォルダに配置
- メールアドレス収集を有効化
- フォーム概要文を設定

---

### 4. `linkResponseSpreadsheet(form, spreadsheetId)`
- フォーム回答先として既存スプレッドシートを紐付け

---

### 5. `renameResponseSheet(spreadsheetId, formTitle)`
- 現状ではコメントアウトされている機能です。  
- フォームの回答先となるシート名を `フォームの回答...` → `formTitle` に変更しようとする処理。

---

### 6. `addQuestions(form)`
- ステータスや貸出、回収、修理、社外持ち出し等、複数のセクションをまとめて呼び出す。  
- *内訳*：
  1. `addStatusQuestion(form)`
  2. `addLendSectionQuestions(form)`
  3. `addReturnSectionQuestions(form)`
  4. `addRepairSectionQuestions(form)`
  5. `addBorrowSectionQuestions(form)`
  6. `addConfirmationSectionQuestions(form)`

---

### 7. `setSectionNavigation(form)`
- 「ステータス」質問の回答に応じたセクション移動を設定  
- 「回収方法」質問に応じて、回答を再スタート or フォーム送信といった処理を設定

---

### 8. `logFormUrls(form)`
- フォームの編集リンク、送信リンクをスプレッドシートにログ出力

---

### 9. `logToSheet(spreadsheet, message)`
- `createForm_script_log` シートに、`[日付, メッセージ]` で追記ログを残す

---

### 10. `generateQrCode(url, formTitle)`
- QuickChart APIを使用してQRコード画像を生成
- 生成画像をシート「QRコード」に貼り付けし、枠線や日時なども付与

---

## 想定される実行タイミング
- ボタン押下などの手動実行

---

## 注意点
- 同名のフォームがフォルダ内にあると作成失敗となる  
- フォームの詳細編集はフォーム本体か、ここに関数を追加して行う

---

# 4. mainシート関数チェック

## スクリプト名・概要
**mainシート関数チェック**  
「main」シートの特定列（D列～J列）に設定されている数式を確認・修正するスクリプトです。  
行ごとに数式が正しいかを比較し、異なる場合はスクリプトが自動的に修正します。

---

## 関数一覧

### 1. `checkAndFixColumnsFormulas()`
1. **対象シートの存在確認**  
   - スプレッドシート内に「main」シートが無い場合は終了

2. **ログシート作成または既存シート取得**  
   - `check_formula_of_mainSheet_log` というシートを探し、無ければ作成  
   - ローテート機能で、3ヶ月より古いログを削除

3. **基準数式の取得**  
   - D列(4列目)の4行目に記載されている数式を「基準」として取得  
   - これを `baseFormula` として他の行にも当てはめて比較

4. **各行に対して数式比較 → 修正**  
   - 1行目〜最終行までループし、基準数式を行数に応じて書き換えたものと現在の数式を比較  
   - 違っていれば自動修正し、ログシートに変更履歴として書き込む

5. **結果ログ出力**  
   - 修正状況をLogger出力およびログシートへ appendRow

---

## 想定される実行タイミング
- 手動または時間ベースのトリガーで定期実行

---

## 注意点
- **前提**：4行目に数式が必ず存在し、それが「正しい数式の基準」であること  
- 対象列は (D〜J列) に固定されているため、変更の際はスクリプトのcolumns指定を編集する必要があります。  
- 大量行数の場合、修正に伴う処理負荷に注意が必要です。

---

# 全体の補足
- いずれのスクリプトも、スクリプトプロパティ(あるいはシート上の特定セル)から設定を取得し、それに基づいて処理を動的に分岐させる形が多く採用されています。  
- デバッグ時は `Logger.log` や独自ログシートの出力をよく確認すると、動作状況やエラー原因を把握しやすくなります。  
- 実運用時には、定期実行する場合のトリガー設定や、想定外のシチュエーション（シート構造変更、権限変更など）が発生した際の対応に注意が必要です。
