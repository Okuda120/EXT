Imports Common
Imports System.Text
Imports Npgsql
Imports CommonEXT
Imports System.Configuration

Public Class LogicEXTY0101

    'インスタンス生成
    Private sqlEXTY0101 As New SqlEXTY0101

    'Private定数宣言
    'EXAS請求依頼情報（シアター）行番号
    Private Const COL_THEATER_CHARGE As Integer = 7          '請求金額

    'EXAS請求依頼情報（スタジオ）行番号
    Private Const COL_STUDIO_CHARGE As Integer = 8           '請求金額

    'Public定数宣言
    Public Const COL_SELECT As Integer = 0                   '選択


    'CSVファイルヘッダ行文字列
    'Private Const CSV_HEADER As String = """請求日"",""入金予定日"",""当社担当者コード"",""当社担当者所属部署コード"",""当社部署""," &
    '                                     """担当者TEL"",""請求先コード"",""請求先郵便番号"",""請求先住所１"",""請求先住所２""," &
    '                                     """請求先名称"",""G請求先部署コード"",""請求先部署名"",""G請求書担当者コード"",""請求書担当者名""," &
    '                                     """経理連絡欄"",""タイトル１"",""タイトル２"",""納品書備考"",""処理区分""," &
    '                                     """番組コード"",""番組シーケンス"",""プロジェクトコード"",""プロジェクトシーケンス"",""目的コード""," &
    '                                     """預り先コード"",""計上部署コード"",""コンテンツ識別区分"",""コンテンツコード"",""コンテンツ内訳コード""," &
    '                                     """予算外案件コード"",""勘定科目コード"",""細目コード"",""内訳コード"",""詳細コード""," &
    '                                     """借方勘定科目コード"",""借方細目コード"",""借方内訳コード"",""借方詳細コード"",""発生月自""," &
    '                                     """発生月至"",""セグメントコード"",""入力摘要１"",""入力摘要２"",""単価""," &
    '                                     """数量"",""消費税額"",""消費税区分"",""消費税率"",""外税内税区分""," &
    '                                     """G請求内容コード"",""G_セグメントコード"",""Gコンテンツ識別区分"",""Gコンテンツコード"",""Gコンテンツ内訳コード"""
    Private Const CSV_HEADER As String = "請求日,入金予定日,当社担当者コード,当社担当者所属部署コード,当社部署," &
                                         "担当者TEL,請求先コード,請求先郵便番号,請求先住所１,請求先住所２," &
                                         "請求先名称,G請求先部署コード,請求先部署名,G請求書担当者コード,請求書担当者名," &
                                         "経理連絡欄,タイトル１,タイトル２,納品書備考,処理区分," &
                                         "番組コード,番組シーケンス,プロジェクトコード,プロジェクトシーケンス,目的コード," &
                                         "預り先コード,計上部署コード,識別区分,コンテンツコード,コンテンツ内訳コード," &
                                         "予算外案件コード,勘定科目コード,細目コード,内訳コード,詳細コード," &
                                         "借方勘定科目コード,借方細目コード,借方内訳コード,借方詳細コード,発生月自," &
                                         "発生月至,セグメントコード,入力摘要１,入力摘要２,単価," &
                                         "数量,消費税額,消費税区分,消費税率,外税内税区分," &
                                         "G請求内容コード,Gセグメントコード,Gコンテンツ識別区分,Gコンテンツコード,Gコンテンツ内訳コード"


    ''' <summary>
    ''' EXAS請求依頼データ作成画面表示メイン処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼データ作成画面の表示を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitDisplayMain(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'EXAS請求依頼情報を取得する
        If GetExasRequestData(dataEXTY0101) = False Then
            Return False
        End If

        'スプレッドシートの表示設定
        If SetDisplayControl(dataEXTY0101) = False Then
            Return False
        End If

        ' 2016.01.22 MOD START↓ y.morooka グループ請求対応　チェックボックス選択不可対応
        'グループ請求チェックボックス非活性処理
        If CheckAitesaki(dataEXTY0101) = False Then
            Return False
        End If
        '2016.01.22 MOD END ↑ y.morooka グループ請求対応　　チェックボックス選択不可対応

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常処理終了
        Return True

    End Function

    ''' <summary>
    ''' EXAS請求依頼データCSVファイル作成メイン処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼データCSVファイルの作成を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function OutputCsvMain(ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'チェックボックスが選択されているかのチェック
        If InputCheck(dataEXTY0101) = False Then
            Return False
        End If

        '検索条件取得
        If GetSearchConition(dataEXTY0101) = False Then
            Return False
        End If

        'CSV出力ファイル用データ取得
        'If GetOutputCsvData(dataEXTY0101) = False Then
        '    Return False
        'End If

        'CSV出力するか確認メッセージ
        If MsgBox(Y0101_C0009, MsgBoxStyle.OkCancel, TITLE_INFO) = vbOK Then

            'CSVファイル出力
            If OutputCsvData(dataEXTY0101) = False Then
                Return False
            End If

            'CSVが出力された場合
            If dataEXTY0101.PropBlnCsvOutputFlg Then
                'EXAS請求依頼フラグ更新
                If UpdateFlg(dataEXTY0101) = False Then
                    Return False
                End If

                '出力処理正常終了時、メッセージを出力
                MsgBox(String.Format(Y0101_I0002, "CSVファイル出力"))

                '表示設定
                If InitDisplayMain(dataEXTY0101) = False Then
                    Return False
                End If
            End If

        End If

        '正常終了
        Return True

    End Function

    ''' <summary>
    ''' EXAS請求依頼情報取得処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼情報を取得する
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetExasRequestData(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'データテーブルを初期化
            If dataEXTY0101.PropDtExasRequestData IsNot Nothing Then
                dataEXTY0101.PropDtExasRequestData.Dispose()
            End If

            'ラジオボタン押下チェック
            If dataEXTY0101.PropRdoTheater.Checked Then
                'シアターの場合
                If GetExasRequestTheater(dataEXTY0101) = False Then
                    Return False
                End If
            ElseIf dataEXTY0101.PropRdoStudio.Checked Then
                'スタジオの場合
                If GetExasRequestStudio(dataEXTY0101) = False Then
                    Return False
                End If
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' スプレッドシート表示設定処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>スプレッドシートの表示設定を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetDisplayControl(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            '通貨型セルの設定
            Dim currcell As New FarPoint.Win.Spread.CellType.CurrencyCellType()
            currcell.CurrencySymbol = "\"
            currcell.ShowSeparator = True

            With dataEXTY0101

                'ラジオボタン押下チェック
                If .PropRdoTheater.Checked Then
                    'シアターがチェックされている場合
                    .PropVwBillpay.Sheets(THEATER).DataSource = .PropDtExasRequestData
                    'シアターのスプレッドシートを表示、スタジオのスプレッドシートを非表示
                    .PropVwBillpay.Sheets(THEATER).Visible = True
                    .PropVwBillpay.Sheets(STUDIO).Visible = False
                    'スプレッドシートのサイズを変更
                    .PropVwBillpay.Width = 1428
                    'スプレッドシートのセル型を設定
                    .PropVwBillpay.Sheets(THEATER).Columns(COL_THEATER_CHARGE).CellType = currcell


                ElseIf .PropRdoStudio.Checked Then
                    'スタジオがチェックされている場合
                    .PropVwBillpay.Sheets(STUDIO).DataSource = .PropDtExasRequestData
                    'スタジオのスプレッドシートを表示、シアターのスプレッドシートを非表示
                    .PropVwBillpay.Sheets(THEATER).Visible = False
                    .PropVwBillpay.Sheets(STUDIO).Visible = True
                    'スプレッドシートのサイズを変更
                    .PropVwBillpay.Width = 1518
                    'スプレッドシートのセル型を設定
                    .PropVwBillpay.Sheets(STUDIO).Columns(COL_STUDIO_CHARGE).CellType = currcell

                End If

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 入力チェック処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>CSV作成前の入力チェック処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function InputCheck(ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim intDataRow As Integer = 0
        Dim intCountCheck As Integer = 0
        Dim vwActiveSheet As FarPoint.Win.Spread.SheetView

        Try
            'スプレッドシートのアクティブシートを取得
            vwActiveSheet = dataEXTY0101.PropVwBillpay.ActiveSheet
            'スプレッドシートのデータ行数取得
            intDataRow = vwActiveSheet.Rows.Count
            '選択されたチェックボックスの存在チェック
            For i As Integer = 0 To intDataRow - 1
                If vwActiveSheet.Cells(i, 0).Value Then
                    intCountCheck = intCountCheck + 1
                End If
            Next

            '選択されたチェックボックスが存在しない場合、エラー
            If intCountCheck = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, String.Format(Y0101_E0014, "出力する請求データ"), Nothing, Nothing)
                puErrMsg = String.Format(Y0101_E0014, "出力する請求データ")
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 検索条件取得処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>スプレッドシートから検索条件となる項目を取得する
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetSearchConition(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim vwActiveSheet As FarPoint.Win.Spread.SheetView
        Dim sbSearchCondition As New StringBuilder()
        Dim intCountRow As Integer = 0
        Dim arySearchConditionList As New ArrayList

        Try
            'スプレッドシートのアクティブシートを取得
            vwActiveSheet = dataEXTY0101.PropVwBillpay.ActiveSheet
            'スプレッドシートのデータ行数を取得
            intCountRow = vwActiveSheet.Rows.Count
            'データ行をループし、チェックボックスが選択されている行の請求依頼番号を取得
            For i As Integer = 0 To intCountRow - 1
                If vwActiveSheet.Cells(i, 0).Value Then
                    '検索条件格納用ハッシュテーブルを宣言
                    Dim htSearchCondition As New Hashtable
                    '表示されているシートの判別
                    ' 2016.01.22 MOD START↓ y.morooka グループ請求対応　スプレッドに１列追加したため
                    'If vwActiveSheet.ColumnCount = 12 Then
                    If vwActiveSheet.ColumnCount = 13 Then
                        '2016.01.22 MOD END ↑ y.morooka グループ請求対応　　スプレッドに１列追加したため
                        'シアターの場合
                        htSearchCondition.Add("seikyu_irai_no", vwActiveSheet.Cells(i, 1).Value)
                        htSearchCondition.Add("yoyaku_no", vwActiveSheet.Cells(i, 10).Value)
                        htSearchCondition.Add("seq", vwActiveSheet.Cells(i, 11).Value)
                    Else
                        'スタジオの場合
                        htSearchCondition.Add("seikyu_irai_no", vwActiveSheet.Cells(i, 1).Value)
                        htSearchCondition.Add("yoyaku_no", vwActiveSheet.Cells(i, 11).Value)
                        htSearchCondition.Add("seq", vwActiveSheet.Cells(i, 12).Value)
                    End If

                    'ハッシュテーブルをリストへ格納
                    arySearchConditionList.Add(htSearchCondition)
                End If
            Next

            '検索条件をデータクラスにセット
            dataEXTY0101.PropArySearchConditionList = arySearchConditionList

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' CSV出力処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼データのCSV出力を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function OutputCsvData(ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strFileName As String = ""                                                  '出力ファイル名
        Dim strFilePath As String = ""                                                  '出力ファイルパス
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") '出力ファイルの文字コード
        Dim dtCsvData As DataTable                                                      'CSVファイル用データテーブル
        Dim arySearchCondition As ArrayList = dataEXTY0101.PropArySearchConditionList   '検索条件格納用リスト
        Dim intArrayCount As Integer = arySearchCondition.Count                         '検索条件数
        Dim intDtCsvDataCount As Integer = 0                                            '請求依頼詳細カウンタ       '2016.10.25 e.watanabe add

        Try
            '出力ファイルパス設定
            strFileName = "EXTY0101_" & DateTime.Now.ToString("yyyyMMddHHmmss")

            ' ファイル保存ダイアログ起動
            Using sfd As New SaveFileDialog()
                sfd.Filter = "EXAS請求依頼データ(*.csv)|*.csv"
                sfd.FileName = strFileName

                If sfd.ShowDialog() = DialogResult.OK Then
                    ' 2016.08.12 ADD START↓ m.hayabuchi 代行処理対応
                    '代行フラグを確認
                    If CommonDeclareEXT.PropStrDaikoFlg.Equals("1") Then
                        '代行処理の場合
                        dataEXTY0101.PropStrTantoNm = CommonDeclareEXT.PropStrDaikoTanto
                        dataEXTY0101.PropStrTantoBusho = CommonDeclareEXT.PropStrDaikoBusho
                    Else
                        dataEXTY0101.PropStrTantoNm = ""
                        dataEXTY0101.PropStrTantoBusho = ""
                    End If
                    ' 2016.08.12 ADD END↑ m.hayabuchi 代行処理対応

                    'OKが押下された場合
                    'ファイル名を設定
                    strFilePath = sfd.FileName

                    '書き込むファイルを開く
                    Dim swCsvFile As New System.IO.StreamWriter(strFilePath, False, enc)

                    'ヘッダを書き込む
                    swCsvFile.Write(CSV_HEADER)
                    '改行する
                    swCsvFile.Write(vbCrLf)

                    '検索条件の件数分ループ
                    For i As Integer = 0 To intArrayCount - 1
                        'HashTable取得
                        Dim htSearchCondition As Hashtable = arySearchCondition(i)
                        '検索条件取得
                        '請求依頼番号
                        dataEXTY0101.PropStrBillNo_Search = htSearchCondition("seikyu_irai_no")
                        '予約NO
                        dataEXTY0101.PropstrReserveNo_Search = htSearchCondition("yoyaku_no")
                        '予約NOシーケンス
                        dataEXTY0101.PropStrSeq_Search = htSearchCondition("seq")
                        '検索条件を基にデータを取得
                        If GetOutputCsvData(dataEXTY0101) = False Then
                            Return False
                        End If
                        '取得したデータを格納
                        dtCsvData = dataEXTY0101.PropDtOutputCsvData

                        '請求依頼データを書き込む
                        For Each Row As DataRow In dtCsvData.Rows

                            '単価＝"0"以外を書き込む
                            If Not dtCsvData.Rows(intDtCsvDataCount).Item("keijo_kin").ToString.Equals("0") Then    '2016.10.25 e.watanabe add 不具合対応
                                '請求依頼データ詳細書き込み
                                If WriteCsvFile(swCsvFile, Row, dataEXTY0101) = False Then
                                    Return False
                                End If
                                '改行する
                                swCsvFile.Write(vbCrLf)
                            End If
                            intDtCsvDataCount = intDtCsvDataCount + 1                      '2016.10.25 e.watanabe add 不具合対応
                        Next
                        intDtCsvDataCount = 0                                              '2016.10.25 e.watanabe add 不具合対応

                    Next

                    'CSVファイルを閉じる
                    swCsvFile.Close()
                    'CSV出力フラグを設定
                    dataEXTY0101.PropBlnCsvOutputFlg = True
                Else
                    'キャンセルが押下された場合
                    'CSV出力フラグを設定
                    dataEXTY0101.PropBlnCsvOutputFlg = False
                End If

            End Using

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' EXAS請求依頼フラグ更新処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼フラグ更新処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function UpdateFlg(ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter()    'アダプタ
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション
        Dim arySearchCondition As ArrayList = dataEXTY0101.PropArySearchConditionList   '検索条件格納用リスト
        Dim intArrayCount As Integer = arySearchCondition.Count                         '検索条件数

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '検索条件の件数分ループ
            For i As Integer = 0 To intArrayCount - 1
                'HashTable取得
                Dim htSearchCondition As Hashtable = arySearchCondition(i)
                '検索条件取得
                '請求依頼番号
                dataEXTY0101.PropStrBillNo_Search = htSearchCondition("seikyu_irai_no")
                '予約NO
                dataEXTY0101.PropstrReserveNo_Search = htSearchCondition("yoyaku_no")
                '予約NOシーケンス
                dataEXTY0101.PropStrSeq_Search = htSearchCondition("seq")

                'EXAS請求依頼フラグ更新(UPDATE)
                If UpdateSeikyuIraiFlg(Cn, dataEXTY0101) = False Then
                    'ロールバック
                    If Tsx IsNot Nothing Then
                        Tsx.Rollback()
                    End If
                    Return False
                End If
            Next

            'コミット
            Tsx.Commit()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' EXAS請求依頼情報（シアター）取得処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼情報（シアター）取得処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetExasRequestTheater(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtExasRequestInfo As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            'EXAS請求依頼情報（シアター）取得SQLを設定
            If sqlEXTY0101.SetExasRequestTheaterSql(Adapter, Cn, dataEXTY0101) = False Then
                Return False
            End If

            'EXAS請求依頼情報（シアター）を取得
            Adapter.Fill(dtExasRequestInfo)
            dataEXTY0101.PropDtExasRequestData = dtExasRequestInfo

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtExasRequestInfo.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' EXAS請求依頼情報（スタジオ）取得処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼情報（スタジオ）取得処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetExasRequestStudio(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtExasRequestInfo As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            'EXAS請求依頼情報（スタジオ）取得SQLを設定
            If sqlEXTY0101.SetExasRequestStudioSql(Adapter, Cn, dataEXTY0101) = False Then
                Return False
            End If

            'EXAS請求依頼情報（スタジオ）を取得
            Adapter.Fill(dtExasRequestInfo)
            dataEXTY0101.PropDtExasRequestData = dtExasRequestInfo

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtExasRequestInfo.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' CSV出力用EXAS請求依頼情報取得処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>CSV出力用EXAS請求依頼情報取得処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetOutputCsvData(ByRef dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtOutputCsvData As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            'CSV出力用EXAS請求依頼データを取得する
            If sqlEXTY0101.SetOutputCsvDataSql(Adapter, Cn, dataEXTY0101) = False Then
                Return False
            End If

            'CSV出力用EXAS請求依頼データを取得
            Adapter.Fill(dtOutputCsvData)
            dataEXTY0101.PropDtOutputCsvData = dtOutputCsvData

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtOutputCsvData.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' EXAS請求依頼フラグ更新処理
    ''' </summary>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼フラグ更新処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function UpdateSeikyuIraiFlg(ByVal Cn As NpgsqlConnection, _
                                         ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand

        Try
            'EXAS請求依頼済フラグ更新(UPDATE)用SQLを作成
            If sqlEXTY0101.SetUpdateSeikyuIraiFlgSql(Cmd, Cn, dataEXTY0101) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXAS請求依頼フラグ更新", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' CSVファイル書込み処理
    ''' </summary>
    ''' <param name="sw">[IN/OUT]CSV出力用StreamWriterクラス</param>
    ''' <param name="dr">[IN]EXAS請求依頼情報DataRowクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>CSVファイルへの書込み処理を行う
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function WriteCsvFile(ByRef sw As System.IO.StreamWriter, ByVal dr As DataRow, ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '変数宣言
        Dim strDouble As String = """"      'ダブルクォーテーション
        Dim strEmpty As String = """"""     'CSV出力時の空白項目
        Dim comma As String = ","           'カンマ
        Dim strContHead As String = "PA0"

        Try
            'sw.Write(strDouble & dr(0) & strDouble & comma)     '請求日
            'sw.Write(strDouble & dr(1) & strDouble & comma)     '入金予定日
            'sw.Write(strEmpty & comma)                          '当社担当者コード
            'sw.Write(strEmpty & comma)                          '当社担当者所属部署コード
            'sw.Write(strEmpty & comma)                          '当社部署
            'sw.Write(strEmpty & comma)                          '担当者Ｔｅｌ
            'sw.Write(strDouble & dr(2) & strDouble & comma)     '請求先コード
            'sw.Write(strDouble & dr(3) & strDouble & comma)     '請求書郵便番号
            'sw.Write(strDouble & dr(4) & strDouble & comma)     '請求書住所１
            'sw.Write(strDouble & dr(5) & strDouble & comma)     '請求書住所２
            'sw.Write(strDouble & dr(6) & strDouble & comma)     '請求先名称
            'sw.Write(strEmpty & comma)                          'G請求先部署コード
            'sw.Write(strEmpty & comma)                          '請求書部署名
            'sw.Write(strEmpty & comma)                          'G請求書担当者コード
            'sw.Write(strEmpty & comma)                          '請求書担当者名
            'sw.Write(strEmpty & comma)                          '経理連絡欄
            'sw.Write(strDouble & dr(7) & strDouble & comma)     'タイトル１
            'sw.Write(strDouble & dr(8) & strDouble & comma)     'タイトル２
            'sw.Write(strEmpty & comma)                          '納品書備考
            'sw.Write(strDouble & "2" & strDouble & comma)       '処理区分(2)
            'sw.Write(strEmpty & comma)                          '番組コード
            'sw.Write(strEmpty & comma)                          '番組シーケンス
            'sw.Write(strDouble & dr(9) & strDouble & comma)     'プロジェクトコード
            'sw.Write(strDouble & dr(10) & strDouble & comma)    'プロジェクトシーケンス
            'sw.Write(strEmpty & comma)                          '目的コード
            'sw.Write(strEmpty & comma)                          '預り先コード
            'sw.Write(strEmpty & comma)                          '計上部署コード
            'sw.Write(strEmpty & comma)                          'コンテンツ識別区分
            'sw.Write(strEmpty & comma)                          'コンテンツコード
            'sw.Write(strEmpty & comma)                          'コンテンツ内訳コード
            'sw.Write(strEmpty & comma)                          '予算外案件コード
            'sw.Write(strDouble & dr(11) & strDouble & comma)    '勘定科目コード
            'sw.Write(strDouble & dr(12) & strDouble & comma)    '細目コード
            'sw.Write(strDouble & dr(13) & strDouble & comma)    '内訳コード
            'sw.Write(strDouble & dr(14) & strDouble & comma)    '詳細コード
            'sw.Write(strDouble & dr(15) & strDouble & comma)    '借方勘定科目コード
            'sw.Write(strDouble & dr(16) & strDouble & comma)    '借方細目コード
            'sw.Write(strDouble & dr(17) & strDouble & comma)    '借方内訳コード
            'sw.Write(strDouble & dr(18) & strDouble & comma)    '借方詳細コード
            'sw.Write(strDouble & dr(19) & strDouble & comma)    '発生月自
            'sw.Write(strEmpty & comma)                          '発生月至
            'sw.Write(strEmpty & comma)                          'セグメントコード
            'sw.Write(strDouble & dr(20) & strDouble & comma)    '入力摘要１
            'sw.Write(strDouble & dr(21) & "[" & dr(22) & "]" & strDouble & comma)    '入力摘要２
            'sw.Write(strDouble & dr(23) & strDouble & comma)    '単価
            'sw.Write(strDouble & "1" & strDouble & comma)       '数量(1)
            'sw.Write(strDouble & dr(24) & strDouble & comma)    '消費税額
            'sw.Write(strEmpty & comma)                          '消費税区分
            'sw.Write(strDouble & dr(25) & strDouble & comma)    '消費税率
            'sw.Write(strEmpty & comma)                          '外税内税区分
            'sw.Write(strEmpty & comma)                          'G請求内容コード
            'sw.Write(strEmpty & comma)                          'G_セグメントコード
            'sw.Write(strEmpty & comma)                          'Gコンテンツ識別区分
            'sw.Write(strEmpty & comma)                          'Gコンテンツコード
            'sw.Write(strEmpty & comma)                          'Gコンテンツ内訳コード
            sw.Write(dr(0))                  '請求日
            sw.Write(comma & dr(1))                  '入金予定日
            ' 2016.08.12 UPD START↓ m.hayabuchi 代行処理時は担当者コード・担当者所属部署コードを設定
            'sw.Write(comma)                          '当社担当者コード
            'sw.Write(comma)                          '当社担当者所属部署コード
            sw.Write(comma & dataEXTY0101.PropStrTantoNm)       '当社担当者コード
            sw.Write(comma & dataEXTY0101.PropStrTantoBusho)    '当社担当者所属部署コード
            ' 2016.08.12 UPD END↑ m.hayabuchi
            sw.Write(comma)                          '当社部署
            sw.Write(comma)                          '担当者Ｔｅｌ
            sw.Write(comma & dr(2))                  '請求先コード
            sw.Write(comma & dr(3))                  '請求書郵便番号
            sw.Write(comma & dr(4))                  '請求書住所１
            sw.Write(comma & dr(5))                  '請求書住所２
            sw.Write(comma & dr(6))                  '請求先名称
            sw.Write(comma)                          'G請求先部署コード
            sw.Write(comma)                          '請求書部署名
            sw.Write(comma)                          'G請求書担当者コード
            sw.Write(comma)                          '請求書担当者名
            sw.Write(comma)                          '経理連絡欄
            sw.Write(comma & dr(7))                  'タイトル１
            sw.Write(comma & dr(8))                  'タイトル２
            sw.Write(comma)                          '納品書備考
            sw.Write(comma & "2")                    '処理区分(2)
            sw.Write(comma)                          '番組コード
            sw.Write(comma)                          '番組シーケンス
            sw.Write(comma & strContHead & dr(9))    'プロジェクトコード
            sw.Write(comma & dr(10))                 'プロジェクトシーケンス
            sw.Write(comma)                          '目的コード
            sw.Write(comma)                          '預り先コード
            sw.Write(comma)                          '計上部署コード
            sw.Write(comma)                          'コンテンツ識別区分
            sw.Write(comma)                          'コンテンツコード
            sw.Write(comma)                          'コンテンツ内訳コード
            sw.Write(comma)                          '予算外案件コード
            sw.Write(comma & dr(11))                 '勘定科目コード
            sw.Write(comma & dr(12))                 '細目コード
            sw.Write(comma & dr(13))                 '内訳コード
            sw.Write(comma & dr(14))                 '詳細コード
            sw.Write(comma & dr(15))                 '借方勘定科目コード
            sw.Write(comma & dr(16))                 '借方細目コード
            sw.Write(comma & dr(17))                 '借方内訳コード
            sw.Write(comma & dr(18))                 '借方詳細コード
            sw.Write(comma & dr(19))                 '発生月自
            sw.Write(comma)                          '発生月至
            sw.Write(comma)                          'セグメントコード
            sw.Write(comma & dr(20))                 '入力摘要１
            sw.Write(comma & dr(21) & "[" & dr(22) & "]")    '入力摘要２
            sw.Write(comma & dr(23))                 '単価
            sw.Write(comma & "1")                    '数量(1)
            ' 2016.03.17 UPD START↓ h.hagiwara 利用料は課税
            'If dr(26) = "1" Then
            '    sw.Write(comma & "0")                    '消費税額
            '    sw.Write(comma & "10")                   '消費税区分
            '    sw.Write(comma)                          '消費税率
            '    sw.Write(comma & "2")                    '外税内税区分
            'Else
            '    sw.Write(comma & dr(24))                 '消費税額
            '    sw.Write(comma)                          '消費税区分
            '    sw.Write(comma & dr(25))                 '消費税率
            '    sw.Write(comma)                          '外税内税区分
            'End If
            If dr(27) = "1" Then
                sw.Write(comma & dr(24))                 '消費税額
                sw.Write(comma)                          '消費税区分
                sw.Write(comma & dr(25))                 '消費税率
                sw.Write(comma)                          '外税内税区分
            Else
                If dr(26) = "1" Then
                    sw.Write(comma & "0")                '消費税額
                    sw.Write(comma & "10")               '消費税区分
                    sw.Write(comma)                      '消費税率
                    sw.Write(comma & "2")                '外税内税区分
                Else
                    sw.Write(comma & dr(24))             '消費税額
                    sw.Write(comma)                      '消費税区分
                    sw.Write(comma & dr(25))             '消費税率
                    sw.Write(comma)                      '外税内税区分
                End If
            End If
            ' 2016.03.17 UPD END↑ h.hagiwara 利用料は課税
            sw.Write(comma)                          'G請求内容コード
            sw.Write(comma)                          'G_セグメントコード
            sw.Write(comma)                          'Gコンテンツ識別区分
            sw.Write(comma)                          'Gコンテンツコード
            sw.Write(comma)                          'Gコンテンツ内訳コード

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickVwCellMain(ByRef dataEXTY0101 As DataEXTY0101) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスをラジオボタンのように制御する
            If SetCheckAsRadio(dataEXTY0101) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' チェックボックス疑似ラジオボタン制御処理
    ''' <paramref name="dataEXTY0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckAsRadio(ByRef dataEXTY0101 As DataEXTY0101) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTY0101
                'ON状態のチェックボックスを選択した場合、自分自身のチェックを外す
                'OFF状態のチェックボックスを選択した場合、選択行にはチェックをつけ、それ以外にはチェックを外す
                For i As Integer = 0 To .PropVwBillpay.ActiveSheet.RowCount - 1

                    'ON状態のチェックボックスを選択した場合
                    If .PropVwBillpay.ActiveSheet.GetValue(i, COL_SELECT) = True Then
                        .PropVwBillpay.ActiveSheet.SetValue(i, COL_SELECT, False)
                    Else
                        'チェックを外す
                        .PropVwBillpay.ActiveSheet.SetValue(i, COL_SELECT, False)

                        '選択行にチェックをつける
                        If i = .PropIntCheckRow Then
                            .PropVwBillpay.ActiveSheet.SetValue(i, COL_SELECT, True)
                        End If

                    End If

                Next

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function


    ''' <summary>
    ''' グループ請求チェックボックス非活性処理
    ''' </summary>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>グループ請求チェックボックス非活性処理を行う
    ''' <para>作成情報：2016/01/22 y.morooka 
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function CheckAitesaki(ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim intDataRow As Integer = 0
        Dim vwActiveSheet As FarPoint.Win.Spread.SheetView

        Try
            'スプレッドシートのアクティブシートを取得
            vwActiveSheet = dataEXTY0101.PropVwBillpay.ActiveSheet
            'スプレッドシートのデータ行数取得
            intDataRow = vwActiveSheet.Rows.Count
            'チェックボックスの非活性チェック
            For i As Integer = 0 To intDataRow - 1
                If vwActiveSheet.ColumnCount = 13 Then
                    'シアターの場合
                    If PropGSeikyusaki Like "*" & vwActiveSheet.Cells(i, 12).Value & "*" Then
                        'チェックボックスを非活性
                        vwActiveSheet.Cells(i, 0).Locked = True   '選択の入力を不可とする
                    End If
                Else
                    'スタジオの場合
                    If PropGSeikyusaki Like "*" & vwActiveSheet.Cells(i, 13).Value & "*" Then
                        'チェックボックスを非活性
                        vwActiveSheet.Cells(i, 0).Locked = True   '選択の入力を不可とする
                    End If
                End If

            Next

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

End Class
