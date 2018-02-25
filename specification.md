# 設計仕様（第１版）
## 1　初めに
ＺＵＫＩシステムの設計に必要な内容を以下に記す。
## 2　共通関数、サブルーチン
### 2.1 システム制御

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Sub |SetOfcn|||||
|Sub |StartOfProcess||||アプリ開始時の処理|
|Sub |EndOfProcess||||アプリ終了時の処理|
|Sub |GetStartData||||アプリ起動時の初期値設定：アカウント情報、消費税率<br>StartOfProcessから呼ばれる|
|Sub |GetConsumptionTaxRate||||消費税率の設定|
|Sub |SetShippingCondition||||出荷時の条件（[システム定数](#システム定数)）設定処理|
|Sub |SearchFolder||||「フォルダを選択する」ダイアログボックスを開いてフォルダパスを取得する<br>(参照ダイアログボックス使用)  2010/07/17 pPoy|
|Func |SearchFile||||「ファイルを開く」ダイアログボックスで、ファイルのフルパスを取得する<br>(参照ダイアログボックス使用) 2010/07/17 pPoy|
|Func |SearchPictureFile|stInitialFileName As String|||「ファイルを開く」ダイアログボックスで、ファイルのフルパスを取得する<br>(参照ダイアログボックス使用) 2010/07/17 pPoy|
|Func |ExGetExt|sfina As String||||
|Func |ExportToExcel|frm As Form|||'現在、未使用　20170114|
|Func |ExportTableToExcel|WorkTableName, strSheetName|||　|
|Func |ControlOption|lngID||||

### 2.2 メッセージ制御

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |MsgOfDebug|comMsg||||
|Sub |TransactionMsg|comMsg, AddContent|||　|
### 2.3 デバッグ機能
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Sub |GC|||| MsgBox "ガベージコレクションが完了しました"|
|Sub |SetLog|intLevel As Integer, Description As String||||
|Sub |MsgboxAndSetErrLog|strDescription As String|||
|Sub |LogClear|||||
|Sub |SQLMsg|strSQLA As String||||
|Sub |DataTableClear||||「tbl0002データ版数管理」の「初期設定対象」がYseに対応する実テーブルを初期設定する。|
|Func |ExportTables||||全ワークkテーブルを既定のフォルダ"\テーブル\ZUKIテーブル_v3_00_00.accdb"にエクスポート|

### 2.4 データチェック

| 種類 | 　　名　称　| 引　数　| 戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |CountRecords_Table|TableName||×||
|Func |CountRecords_WorkTable|WorkTableName||×||
|Func |CountRecords_WorkTable_Wherecond|WorkTableName, strWhereCond||×||
|Func |CountRecords_Table_Wherecond|TableName, strWhereCond||×|結果：レコード数、"0"または"-1”|
|Func |CheckOfRecordExistance|TableName, KeyFiled1, KeyValue1, KeyFiled2, KeyValue2||×|条件と合致するレコードの有無チェック（重複レコードチェックの目的）|
|Func |SyncCheckState_Sub|SubTableName, KeyField, KeyValue, StateID|True/False|×|KeyFieldの値KeyValueが同じすべての明細レコードが同一の状況IDの場合、trueを返す。トランザクション処理化が必要|
|Func |CheckOfExistance_Null_Wherecond|WorkTableName, strField, strWhereCond|True/False|×|FieldにNullのレコードが存在する場合はTrueを返す。|
|Func |CheckOfExistance_Zero_Wherecond|WorkTableName, strField, strWhereCond|True/False|×| FieldにZeroのレコードが存在する場合はTrueを返す。|
|Func |CheckOfExistance_Null|WorkTableName, strField|True/False|×|FieldにNullのレコードが存在する場合はTrueを返す。|
|Func |ConvertOfNullParam|strField||×|　|

### 2.5 データ計算

| 種類 | 　　名　称　| 引　数　| 戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Sub |Summation|WorkTableName, QueryName As String, ObjectItem As String, MainID As Long||×||
|Sub |Multiplication|WorkTableName, strField As Variant, KeyField1 As Variant, KeyField2 As Variant||×||

### 2.7 データ設定

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |SetData_Wherecond|TableName, strField, varValue, strWhereCond|True/False|×|strWhereCondで指定する全レコードに、varValueで指定された値をstrFieldで指定されたフィールドに設定する|
|Sub |SetData_Work_Wherecond|WorkTableName, strField, varValue, strWhereCond||×|ワークテーブルのstrWhereCondで指定する全レコードに、varValueで指定された値をFieldで指定されたフィールドに設定する|
|Sub |SetData_Work_All|WorkTableName, strField, varValue||×|ワークテーブルの全レコードに対し、指定フィールドに指定された値を設定する|
|Func |SetData_MainAndSub|MainTableName, SubTableName, strField, varValue, KeyField, KeyValue||○|　|
|Func |SetLineNo_Work_Plural|MainWorkTableName, SubWorkTableName, strKeyField, strLineNoField|||　|

### 2.8 データ更新

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |UpdateData_Plural_Conditioned|WorkTableName, TableName, strField, KeyField||○|両テーブルのKeyField（重複なし）の値が一致しているレコードに対し、ワークテーブルのField値をテーブルのFieldに設定する。

### 2.9 データ抽出

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GetData_Wherecond|TableName, strField, strWhereCond|strFieldの値|×|条件式（strWhereCond）の最初のレコードのstrFieldの値を出力する。該当値がNullの場合または対象レコードがない場合は""を返す|
|Func |GetData_Work_Wherecond|WorkTableName, strField, strWhereCond|strFieldの値|×|条件式（strWhereCond）の最初のレコードのstrFieldの値を出力する。該当値がNullの場合または対象レコードがない場合は""を返す|
|Func |GetMaxValue_Work_Wherecond|WorkTableName, strField, strWhereCond|strFieldの値|×|条件式（strWhereCond）のレコードのうち、strFieldの最大値を出力する。該当値がNullの場合または対象レコードがない場合は""を返す|
### 2.10 テーブル複写

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |CopyToWorkTable|TableName, WorkTableName|True/False<br>RecordNumberIsZero|○||
|Func |CopyToTable|WorkTableName, TableName|True/False|○|　|
|Func |CopyWorkTable|WorkTableNameIn, WorkTableNameOut|True/False|×|　|

### 2.11 レコード更新

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |UpdateRecords_Conditioned|WorkTableName, TableName, KeyFieldMain, KeyField, KeyValue||○|条件(KeyFieldの値＝KeyValue)に合致したすべてのレコードに対し、KeyFieldMainが同じレコードを更新する。【注意】処理FLGはリセットされる。|
|Func |UpdateRecord_MainAndSub|MainWorkTableName, MainTableName, SubWorkTableName, SubTableName, MainKeyField As String, SubmainKeyField As String, SubKeyField As String, MainKeyValue As Variant||○|条件に合致した親子レコードを更新する。各KeyFieldは管理ID、管理明細IDであり、SubMainKeyFieldはMain/Subの連携IDである。|

### 2.12 レコード削除

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Sub |DeleteRecords_WhereCond|TableName, strWherecond||×|WhereCond条件で指定されるレコード消去|
|Sub |DeleteRecords_Work_WhereCond|WorkTableName, strWherecond||×|ワークテーブルWorkTableNameのWhereCond条件で指定されるレコード消去|
|Sub |DeleteRecords_Work_All|WorkTableName||×|ワークテーブルWorkTableNameの全レコード消去|

### 2.13 レコード追加

| 種類 | 　　名　称　| 引　数　| 戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |AddRecordAndSetMaxMainAndSub|TableName, KeyFieldMain, KeyFieldSub|True/False<br>lngMaxMain,lngMaxSub|○|KeyFieldMainおよびKeyFieldSubの最大値を抽出し、追加した新しいレコードにカウントアップ値（＋１）を設定する。KeyFieldはオートナンバリング型、非オートナンバリングとも可|
|Func |AddRecordAndSetMax|TableName, KeyField|True/False<br>lngMax|○|KeyFieldの最大値を抽出し、追加した新しいレコードにカウントアップ値（＋１）を設定する。KeyFieldはオートナンバリング型、非オートナンバリングとも可|
|Func |AddRecordAndSetMax_plusWork_AndSetData|TableName, WorkTableName, KeyField, strField, varValue|True/False<br>lngMax|○|TableNameのKeyFieldの最大値を抽出し、追加した新しいレコードにカウントアップ値（＋１）を設定する（AddRecordAndSetMax）と共に、WorkTableNameにはレコード追加、キー値（lngMax）設定、指定値varValueのフィールドstrFieldへの設定をおこなう|
|Func |AddRecords|WorkTableName, TableName, KeyField|True/False<br>lngMax|○|WorkTableNameのすべてのレコードを追加登録する。KeyFieldで指定する主キーはカウントアップされる。|
|Func |AddRecord_Work_AndSetData|WorkTableName, strField, varValue|True/False|×|WorkTableNameに1レコードを追加し、指定値varValueをフィールドstrFieldに設定する|

### 2.14 レコード複写

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |CopyRecords_ConditionedByNvarchar|TableName, WorkTableName, KeyField, KeyValue As String|RecordNumberIsZero|○|KeyValueは全角可変文字列、ワールドカード使用|
|Func |CopyRecords_WhereCond|TableName, WorkTableName, strWhereCond|RecordNumberIsZero|×||
|Func |CopyRecords_Work_WhereCond|WorkTableNameIn, WorkTableNameOut, strWhereCond|RecordNumberIsZero|×|コピー先（WorkTableNameOut）をクリアした後、コピー元（WorkTableNameIn）から、条件に合うレコードのみをコピー。|

### 2.15 SQL文作成

| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |CreateSQLParamPart_Dot|KeyField, varKeyValue|||　|

## 3 個別関数、サブルーチン
### 3.1 受注処理
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GenerateSingle_ReceivedOrderFromQuotation|CustomerOrderID|||【５a】|
|Func |GenerateSingle_Work_PurchaseOrderFromReceivedOrder||||【５b】|
|Func |ImportEDIData_ReceivedOrder||||【】|

### 3.2 計画生産
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GeneratePlural_ReceivedOrderFromProductPlanning||||【２a】画面から入力されている子（tbl0201w受注明細）から親（tbl0200w受注）を生成|
|Func |Generate_ProductPlanningLine||||【２aa】画面から入力されているtbl0251w計画生産明細からtbl0251計画生産明細を生成|

### 3.3 部材調達
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GenerateSingle_Work_PurchaseOrderFromBOM||||【５c】<br>Implicit引数：部品調達明細番号:lngSubID、受注明細番号:lngOrderSubID|
|Func |GenerateSingle_Work_ManufacturingOrderFromBOM||||【６】<br>Implicit引数：lngOrderSubID（受注明細番号）、lngSubID（部品調達明細番号）|
|Func |Generate_ReceivedOrderFromMaterialLine||||【】画面から入力されている部材製造手配用tbl0201w受注明細からtbl0201受注明細を生成|

### 3.4 作業指示
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func|Generate_MfgInstAndSubConOrder_FromLine||True/False|○|【２b】指定された受注明細番号（lngOrederSubID）のtbl0301製造依頼工程（内作および外作）に対し、新しい製造指示番号および外注依頼番号を付与し設定する。<br>Implicit引数：lngOrderSubID（受注明細番号）|
|Func |CopyRecords_ManufactureIndicationAndManufactureProcess_plusRelatedReceivedOrderLine|StateID||| 状況IDの製造指示および製造指示明細を抽出コピー。さらに対応する受注明細をコピー。|
|Sub|Create_tbl0405_Calender|strDate|||SetCalenderForManufactureProcess_TimeSerial|
|Sub| Create_tbl0404_TimeSeries||||SetTimeSerialManufactureProcess|
|Sub| Integrate_tbl0405||||SetTimeSerialManufactureProcessBy??????|
|Sub| Create_tbl0401_ProcessTitle||||　|

### 3.5 外注依頼
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
### 3.6 購入依頼
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GeneratePlural_PurchaseOrderFromLine||||【２d】未処理のtbl0601購入依頼明細から親tbl0600購入依頼を作成　（購入先、送付先、納期日でグループ化）|
|Func |ExportEDIData_PurchaseOrder||||【】|
|Sub |SetValue_PurchaseOrderWork_NumberToName|||||
|Sub |SetValue_PurchaseOrderLineWork_NumberToName||||　|

### 3.7 在庫管理
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Sub |CalOfInventoryUpToDate|||||
|Sub |CalOfInventory0630_TimeSeries|||||
|Sub |Create_tbl0635_Calender|strDate As String|||　|

### 3.8 支払書
| 種類 | 　　名　称　| 引　数　| 戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GenerateSingle_PaymentLineFromPurchaseOrderLine|SubID As Long|||【３c】SubIDで指定された購入品明細レコードを支払明細テーブルにコピーし、（購入品および）購入品明細レコードの支払状況を２に設定する|
|Func |GeneratePlural_PaymentFromLine||||【２f】

### 3.9 出荷処理
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GeneratePlural_ShippingFromReceivedOrder||||【５b】状況ID＝5の受注レコードを出荷テーブルにコピーし、かつ受注明細も出荷明細にコピーし、受注および受注明細の状況IDを6に設定する|

### 3.10 請求書
| 種類 | 　　名　称　| 引　数　|  戻り値、出力データ　| Tran性|　説　明 |
|:-----------:|:-----------|:-----------|:-----------|:-----------:|:-----------|
|Func |GeneratePlural_InvoiceLineFromShippingLine|MainID As Long|||【３b】MainID（出荷番号）で指定された出荷明細レコードを請求明細テーブルにコピーする|
|Func |GeneratePlural_InvoiceFromLine||||【２e】|

## 4　共通定義
### 4.1 Global変数
| 種類 | 　　　内　容　| 　説明 |
|:-----------:|:-----------|:-----------|
|Global| strConnectionString| |
|Global| strFullPathDataSource| |
|Global| strDBType||
|Global| strProvider||
|Global| strDriver||
|Global| strServer||
|Global| strPath| As String
|Global |strDataSource| As String
|Global |strPort| As String
|Global| strUserID| As String
|Global| strPassword| As String
|Global| strImagePath| As String
|Global| strDataVersion| As String
|Global| strDataVersion_Old| As String
|Global| strDataVersion_New| As String
|Global| strDataPath| As String
|Global| strDataFile |As String
|Global| strPictureFullPath| As String
|Global| strPicturePath| As String
|Global| strPictureFile| As String|
|Global| lngCnOpenError| As Long|
|Global| KeyF| |
|Global| lngMax|
|Global| lngMaxMain|
|Global| lngMaxSub|
|Global| lngCount|
|Global| idx| As Long|
|Global| RecordNumberIsZero||
|Global| strOption1|
|Global| strWhereCond|
|Global| strmsg|
|Global| TableName| As String|
|Global| MainTableName| As String|
|Global| SubTableName| As String|
|Global| WorkTableName| As String|
|Global| MainWorkTableName| As String|
|Global| SubWorkTableName| As String|
|Global| WorkTableNameIn| As String|
|Global| WorkTableNameOut| As String|
|Global| WorkTableNameInD |As String|
|Global| WorkTableNameOutD |As String|
|Global| QueryName| As String|
|Global| strMainIDName|    
|Global| strField|
|Global| KeyField| As String|
|Global| KeyFieldMain| As String|
|Global| KeyFieldSub| As String|
|Global| KeyFieldIn| As String|
|Global| KeyFieldOut| As String|
|Global| lngValue |
|Global| varValue |
|Global| KeyValue| As Long|
|Global| varKeyValue|
|Global| strSheetName|
|Global| strTitle|
|Global| strSQL|
|Global| strSQLEnable|
|Global| strSQLEnableIn |
|Global| strSQLEnableOut|
|Global| strSQLIn|
|Global| strSQLOut|
|Global| strSQLInD |
|Global| strSQLOutD|
|Global| lngLogInNo |
|Global| strLogInName|
|Global| strStaffName |           担当者名|
|Global| lngMainID                 |管理ID
|Global| lngSubID                 |管理明細ID
|Global| lngSpecifiedMainID ||
|Global| lngQuotationID             |見積番号|
|Global| lngQuotationSubID             |見積明細番号|
|Global| lngOrderID             |受注番号
|Global| lngOrderSubID           |受注明細番号
|Global| strCustomerQuotationID    |顧客見積番号|
|Global| strCustomerOrderID    |顧客注文番号|
|Global| lngShippingID             |出荷番号
|Global| lngShippingSubID|        出荷明細番号|
|Global| lngBillID|               請求番号|
|Global| lngBillSubID|            請求明細番号|
|Global| lngProcessID              |工程ID|
|Global| lngProductCategoryID    |製品部材区分
|Global| lngProductID             |品目番号
|Global| lngPurchaseID            |購入品番号
|Global| lngManufacturingID      |現在未使用20160106
|Global| lngMaterialID             |材料番号
|Global| lngWarehouseID           |倉庫番号
|Global| strWarehouseName       |倉庫名
|Global| lngStockID As Long               |在庫管理番号
|Global| lngPurchaseCategory    |購入品区分
|Global| lngTreatmentID           |手配区分
|Global| strPurchaseOrderID     |購入費用負担区分
|Global| lngStateID            |現在未使用20160106|
|Global| lngQuantity           |数量|
|Global| varUnit              |単位
|Global| strUnit                |単位
|Global| lngConsumptionTaxRate As Single  |消費税率|
|Global| lngCustomerNo             |顧客番号
|Global| strCustomerCode        |顧客コード
|Global| strCustomerName         |顧客名
|Global| strCustomerStaffName   |顧客担当者名|
|Global| lngSubconID|依頼先番号、購入先番号
|Global| lngDestinationNo          |出荷先番号
|Global| strDestinationCode      |出荷先コード
|Global| strDestinationName      |出荷先名
|Global| strDestinationStaffName|     出荷先担当者名|
|Global| lngHinmokuID          |品目番号
|Global| strHINMEI             |品目名、品目名
|Global| strHINBAN             |品目品番
|Global| strSQLSubItems        |現在未使用20160106
|Global| strSQLSubItemsWithTableName      |現在未使用20160106|
|Global| lngOptionID             |オプション名称ID
|Global| dateToday||
|Global| dateIssueDate |発行日
|Global| dateOrderDate            |受注日
|Global| dateCutoffDate          |締日
|Global| dateCustomerDueDate     |顧客要求納入日
|Global| dateDueDate              |納期日
|Global| dateShippingDate         |納品日、出荷日|
|Global| strDeliveryCond|納入条件
|Global| strPaymentCond|支払条件
|Global| strDeliveryPalce|受渡場所
|Global| strValidityPeriod|見積有効期限
|Global| strRemark                        |記事|
|Global| strOfficeCode         |事業所コード
|Global| strOfficeName        |事業所名|
|Global| strProcessParam    |処理パラメータ
|Global| strProcessParam_StockManagement|        在庫管理　処理パラメータ
|Global| strProcessParam_Master|        マスター管理　処理パラメータ|
|Global| varEDIProcess As Variant|         EDI処理
|Global| lngID||
|Global| strFileName||
|Global| lngParam                 |汎用パラメータ
|Global| lngParam1               |汎用パラメータ
|Global| lngParam2               |汎用パラメータ
|Global| strParam3              |汎用パラメータ
|Global| strParam             |汎用パラメータ
|Global| strParam1           |汎用パラメータ
|Global| strParam2           |汎用パラメータ
|Global| dateParam               |汎用パラメータ
|Global| varParam             |汎用パラメータ
|Global| varNull           |汎用パラメータ|
|Global| lngLeftPosition||
|Global| lngTopPosition||
|Global| lngRed||
|Global| lngBlack||
|Global| lngYellow||
|Global| lngWhite||

### 4.2 Public定数
#### 4.2.1 メッセージ用
| 種類 | 　　　内　容　| 　説明 |
|:-----------:|:-----------|:-----------|
|Public Const |comMsg0001 |登録しました。|
|Public Const |comMsg0002 |削除しました。|
|Public Const |comMsg0003 |内容が変更されています。閉じる前に変更しますか？
|Public Const |comMsg0004 |マスタを削除するとあらゆる処理で障害が発生しますがよろしいですか？|
|Public Const |comMsg0005|削除してよろしいですか？|
|Public Const |comMsg0006 |該当するデータはありません & vbCrLf & vbCrLf|
|Public Const |comMsg0007 |これ以前のコードはありません
|Public Const |comMsg0008 |これ以降のコードはありません|
|Public Const |comMsg0009 |内容が変更されています。印刷前に登録しますか？
|Public Const |comMsg0010 |内容が変更されています。新規データに移動する前に登録しますか？|
|Public Const |comMsg0011 |内容が変更されています。前データへ移動する前に登録しますか？
|Public Const |comMsg0012 |内容が変更されています。後データに移動する前に登録しますか？|
|Public Const |comMsg0013 |処理を中止しました。同時操作で競合している可能性があります。
|Public Const |comMsg0014 |エラー発生処理は完了していません
|Public Const |comMsg0015 |障害発生データは登録できませんでした
|Public Const |comMsg0016 |障害発生データは削除できませんでした|
|Public Const |comMsg0050 |データが重複しています。
|Public Const |comMsg0051 |更新不可のデータです。
|Public Const |comMsg0052 |削除不可のデータです。
|Public Const |comMsg0053 |データ入力が必要です。|
|Public Const |comMsg0054 |入力データが異常です。|
|Public Const |comMsg0055 |処理を中止しました。
|Public Const |comMsg0056 |日付の順序が異常です。正しい値を入力してください。
|Public Const |comMsg0057 |内容が登録されていません。終了してよろしいですか？
|Public Const |comMsg0058 |条件設定が必要です。
|Public Const |comMsg0059 |処理を中止しました。運用者に連絡してください。|
|Public Const |comMsg0060 |処理しました。
|Public Const |comMsg0061 |登録可能数を超えています。追加はできません。
|Public Const |comMsg0062 |追加不可のデータです。|
|Public Const |comMsg0999 |　
|Public Const |comMsg1000 |デバッグ
|Public Const |comMsgTittle1 | 確認|
|Public Const |comMsgTittle2 |エラー|

<a name="システム定数"></a>
#### 4.2.2 システム定数
| 種類 | 　　　内　容　| 　説明 |
|:-----------:|:-----------|:-----------|
|Global Const |Admin_FLG| 0=一般モード<br> 1=管理者モード（状況IDに依らず各種伝票の変更が可能）
|Global Const |Com_FLG| 0=基本システム開発（全機能）|
|Global Const |CS_FLG |  0=運用 <br>  1=試験|
|Global Const |ERRLog_FLG| 0=エラーログ収集停止<br> 1=エラーログ収集中|
|Global Const |GCMode_FLG|   0=終了時ガベージコレクション非適用<BR>  1=終了時ガベージコレクション適用|
|Global Const |lngOfficeNo| 1|
|Global Const |Log_FLG | 0=ログ収集停止<br> 1=レベル１のログ収集中<br>  2=レベル２のログ収集中<br>   3=レベル３のログ収集中|
|Global Const |MDE_FLG| 0=MDBファイル <br> 1=MDEファイル|
|Global Const |SQL_FLG| 0=SQLメッセージの表示停止<br> 1=SQLメッセージの表示|

# Appendix

<a name=システム権限></a>
## システム権限
| 設定値 | 　　権　限　内　容　| 　適　用　例 |
|:-----------:|:-----------|:-----------|
| １|全機能使用可能 | システム管理者 |
| ２| 業務管理者機能および一般機能が使用可能| 業務管理者 |
| ３| 一般機能のみ使用可能 | 業務担当者 |

## 機能リスト
### システム管理者機能
| 機能群 | 　　　　機　能　内　容　|
|:-----------:|:-----------|
| オプション| 事業所管理<br>担当者管理<br>工程管理<br>単位管理<br>システム管理 |
### 業務管理者機能
| 機能群 | 　　　　機　能　内　容　|
|:-----------:|:-----------|
| マスタ管理| 取引先管理（販売先、出荷先、購入先、加工先）<br>品目管理（商品、製造品、購入品、在庫品）<br>倉庫管理 |
### 一般機能
| 機能群 | 　　　　機　能　内　容　|
|:-----------:|:-----------|
| 計画生産| 計画生産作成<br>計画生産一覧　製造手配 |
| 見積受注| 見積作成<br>見積一覧<br>見積結果入力<br>受注作成<br>受注一覧<br>受注品手配 |
| 製造管理| 製造手配作成<br>製造手配一覧<br>作業指示作成<br>作業指示一覧<br>外注依頼作成<br>外注依頼一覧<br>部材調達作成<br>部材調達一覧<br>製造完了 |
| 出荷管理| 出荷伝票作成<br>出荷一覧<br>出荷処理 |
| 請求管理| 請求書作成<br>請求書一覧<br>売掛管理 |
| 在庫管理| 在庫一覧<br>棚卸処理 |
| 購買管理| 購入依頼明細　新規登録<br>購入依頼作成／購入依頼明細更新<br>購入依頼書一覧<br>受入処理<br>検品処理 |
| 支払管理| 支払書作成<br>支払書一覧、発行・更新<br>買掛管理 |
| 集計処理| 受注集計<br>売上集計 |
