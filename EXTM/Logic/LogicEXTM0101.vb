Imports Npgsql
Imports Common
Imports CommonEXT
Imports FarPoint.Win.Spread

''' <summary>
''' EXシアターユーザーマスタメンテ画面Logicクラス
''' </summary>
''' <remarks>EXシアターユーザー検索画面のロジックを定義する
''' <para>作成情報：2015/08/11 hayabuchi
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTM0101

    'インスタンス作成
    Public SqlEXTM0101 As New SqlEXTM0101
    Public CommonLogic As New CommonLogic
    Public CommonLogicEXT As New CommonLogicEXT
    Public CommonValidation As New CommonValidation

    ''' <summary>
    ''' フォーム情報の初期化
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>フォームの情報を初期化する
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitForm(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'フォームの内容を空にする / チェックボックスを外す
            With dataEXTM0101
                .PropTxtSearchUserID.Text = System.String.Empty
                .PropTxtSearchKanjiName.Text = System.String.Empty
                .PropTxtSearchMail.Text = System.String.Empty
                .PropCmbSearchBushoName.SelectedIndex = -1
                .PropChkShoninFlg.Checked = False
                .PropChkMstFlg.Checked = False
                .PropChkStsFlg.Checked = True
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 部署名情報取得
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>部署名情報の取得を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetComboBusho(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim dtBushoNM As New DataTable              '部署名情報データテーブル

        Try

            'コネクションを開く
            Cn.Open()

            '検索用SQLの作成・設定
            If SqlEXTM0101.SetSelectBushoSql(Adapter, Cn, dataEXTM0101) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "初期表示時コンボボックス取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtBushoNM)
            'データテーブルに取得データをセット
            dataEXTM0101.PropDtExtBushoMasta = dtBushoNM

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            dtBushoNM.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 部署名コンボボックスをフォームオブジェクトに設定
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>コンボボックスの内容を取得する
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ComboBushoSet(ByRef dataEXTM0101 As DataEXTM0101)
        ' 開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            'データテーブルをコンボボックスに設定
            If CommonLogic.SetCmbBox(dataEXTM0101.PropDtExtBushoMasta, dataEXTM0101.PropCmbSearchBushoName, True, "", "") = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False

        End Try
    End Function
    ''' <summary>
    ''' ユーザーの取得
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>初期表示時検索を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitSearch(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)                    'コネクション
        Dim Adapter As New NpgsqlDataAdapter                        'アダプタ
        Dim Table As New DataTable()                                'テーブル

        Try
            'コネクションを開く
            Cn.Open()

            '検索用SQLの作成・設定
            If SqlEXTM0101.SetSelectInitEXTUserSearchSql(Adapter, Cn, dataEXTM0101) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXシアターユーザー初期表示時検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            'データテーブルに取得データをセット
            dataEXTM0101.PropDtExtUsrMasta = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 初期表示時検索（コンボボックス取得）
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>初期表示時検索（コンボボックス取得）を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CreateBushoCmbBox(ByVal dataEXTM0101 As DataEXTM0101, _
                                      ByRef ctCmbBusho As CellType.ComboBoxCellType, _
                                      ByVal dtBushoNM As DataTable) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim ArryVal As String()                      'コンボボックス内実値
        Dim ArryTxt As String()                      'コンボボックス表示項目
        Dim intIndex As Integer                      '配列番号

        Try

            'コンボボックス（部署名）
            ReDim ArryVal(dtBushoNM.Rows.Count - 1)
            ReDim ArryTxt(dtBushoNM.Rows.Count - 1)

            intIndex = 0
            For Each Row As DataRow In dtBushoNM.Rows
                'リスト設定

                ArryVal(intIndex) = Row.Item(0)
                ArryTxt(intIndex) = Row.Item(1)

                intIndex = intIndex + 1

            Next

            'リストに表示される項目
            ctCmbBusho.Items = ArryTxt
            'リストの表示項目に対する実値     
            ctCmbBusho.ItemData = ArryVal
            '参照するセルの値の種類を指定
            ctCmbBusho.EditorValue = CellType.EditorValue.ItemData
            ctCmbBusho.Editable = False

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 画面表示
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>シートに情報をセットする
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSheet(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim ctCmbBusho As New CellType.ComboBoxCellType         'コンボボックスオブジェクト初期化

        Try
            With dataEXTM0101
                '一覧シートにデータをセット
                .PropVwList.DataSource = .PropDtExtUsrMasta
            End With

            With dataEXTM0101.PropVwList.Sheets(0)
                '部署名列の設定
                .Columns(5).Width = 172
                .Columns(5).Resizable = True
                .Columns(5).CellType = ctCmbBusho

                'コンボボックス内データ取得
                If CreateBushoCmbBox(dataEXTM0101, ctCmbBusho, dataEXTM0101.PropDtExtBushoMasta) = False Then
                    Return False
                End If

                '部署名コンボボックスリスト設定
                Dim intCmbIndex As Integer            'コンボボックスリストの行番号
                intCmbIndex = 0

                For Each RowDetail As DataRow In dataEXTM0101.PropDtExtUsrMasta.Rows

                    'セル型にコンボオブジェクト設定
                    .Cells(intCmbIndex, 5).CellType = ctCmbBusho

                    'コンボボックスの初期表示
                    .Cells(intCmbIndex, 5).Value = RowDetail.Item("BUSHO_CD")
                    .Cells(intCmbIndex, 5).Locked = False

                    '行番号をカウントアップ
                    intCmbIndex = intCmbIndex + 1

                Next

            End With

            '不変項目の設定（ユーザーIDは変更不可）
            For i As Integer = 0 To dataEXTM0101.PropVwList.Sheets(0).RowCount - 1
                dataEXTM0101.PropVwList.Sheets(0).Cells(i, 1).Locked = True
            Next


            '登録件数+5行の登録枠を表示
            Dim row As Integer = dataEXTM0101.PropVwList.Sheets(0).RowCount
            dataEXTM0101.PropVwList.Sheets(0).RowCount = dataEXTM0101.PropVwList.Sheets(0).RowCount + 5

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 検索
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>フォームから取得した値をもとに検索を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function Search(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)                'コネクション
        Dim Adapter As New NpgsqlDataAdapter                    'アダプタ
        Dim Table As New DataTable()                            'テーブル

        Try
            'コネクションを開く
            Cn.Open()

            '検索用SQLの作成・設定
            If SqlEXTM0101.SetSelectEXTUserSearchSql(Adapter, Cn, dataEXTM0101) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXシアターユーザー検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            'データテーブルに取得データをセット
            dataEXTM0101.PropDtExtUsrMasta = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 入力チェック
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の入力チェックを行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckInputSheet(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '変数宣言
        Dim j As Integer = 0  '変更フラグカウント

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try

            '一覧シートの行数分チェックする
            For i As Integer = 0 To dataEXTM0101.PropVwList.Sheets(0).RowCount - 1

                '変更区分がチェックされている場合
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 0).Value = True Then
                    'フラグが立っていたらカウント
                    j = j + 1

                    '入力チェックを行う 
                    With dataEXTM0101.PropVwList.Sheets(0)

                        'ユーザーID
                        With .Cells(i, 1)
                            '未入力の場合、エラー
                            If .Text.Trim() = "" Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0001, "ユーザーID")
                                'エラーを返す
                                Return False
                            End If
                            '桁数チェック
                            If .Text.Length > 10 Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0009, "ユーザーID", "1", "10")
                                'エラーを返す
                                Return False
                            End If

                        End With

                        'パスワード
                        With .Cells(i, 2)
                            '未入力の場合、エラー
                            If .Text.Trim() = "" Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0001, "パスワード")
                                'エラーを返す
                                Return False
                            End If
                            '桁数チェック
                            If .Text.Length > 20 Or .Text.Length < 6 Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0009, "パスワード", "6", "20")
                                'エラーを返す
                                Return False
                            End If
                            '半角英数字チェック
                            If CommonValidation.IsHalfChar(.Text) = False Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0005, "パスワード")
                                'エラーを返す
                                Return False
                            End If
                        End With

                        '漢字氏名
                        With .Cells(i, 3)
                            '未入力の場合、エラー
                            If .Text.Trim() = "" Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0001, "漢字氏名")
                                'エラーを返す
                                Return False
                            End If
                            '桁数チェック
                            If .Text.Length > 20 Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0009, "漢字氏名", "1", "20")
                                'エラーを返す
                                Return False
                            End If
                        End With

                        'メールアドレス
                        With .Cells(i, 4)
                            '桁数チェック
                            If .Text.Length > 256 Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0012, "メールアドレス")
                                'エラーを返す
                                Return False
                            End If

                            'メールアドレス書式チェック
                            If CommonLogicEXT.IsMailAddress(.Text) = False Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0101_E0008, "メールアドレス")
                                'エラーを返す
                                Return False
                            End If

                        End With

                        '部署名
                        With .Cells(i, 5)
                            '桁数チェック
                            If .Value <> "" Then
                                If .Value > 10 Then
                                    'エラーメッセージ設定
                                    puErrMsg = String.Format(M0101_E0009, "部署コード", "10")
                                    'エラーを返す
                                    Return False
                                End If
                                '半角英数字チェック
                                If CommonValidation.IsHalfChar(.Value) = False Then
                                    'エラーメッセージ設定
                                    puErrMsg = String.Format(M0101_E0005, "部署名")
                                    'エラーを返す
                                    Return False
                                End If
                            End If
                        End With

                    End With

                End If

            Next

            If j = 0 Then
                '変更区分がチェックされていない場合
                'エラーメッセージを設定
                puErrMsg = String.Format(M0101_E0014, "変更区分")
                Return False
            End If

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False

        End Try
        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        '確認メッセージ設定
        puErrMsg = String.Format(M0101_C0011, "ユーザーマスタ")
        Return True

    End Function
    ''' <summary>
    ''' ユーザー名の重複チェック
    ''' </summary>
    ''' <param name="dataEXTM0101">[IN/OUT]ユーザーマスタメンテDataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>スプレッドシート登録項目が重複していないかチェックする
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SelectCountSameUserData(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)                'コネクション
        Dim Adapter As New NpgsqlDataAdapter                    'アダプタ
        Dim dtResult As New DataTable                           'データテーブル

        Try
            'コネクションを開く
            Cn.Open()

            For j As Integer = 0 To dataEXTM0101.PropVwList.Sheets(0).RowCount - 1
                '変更区分がチェックされている場合
                If dataEXTM0101.PropVwList.Sheets(0).Cells(j, 0).Value = True Then
                    '隠し列（＝更新区分）に１が立っていない場合（＝新規登録時のみ）
                    If dataEXTM0101.PropVwList.Sheets(0).Cells(j, 11).Value = False Then

                        '同じユーザー名のデータ有無取得（SELECT）用SQLを作成
                        If SqlEXTM0101.SetSelectCountSameUserSql(Adapter, Cn, j, dataEXTM0101) = False Then
                            Return False
                        End If

                        'ログ出力
                        CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "同じユーザー名のデータ有無取得", Nothing, Adapter.SelectCommand)

                        'SQL実行
                        Adapter.Fill(dtResult)


                        '重複データがある場合、エラー
                        If dtResult.Rows.Count > 0 Then

                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0101_E1002, "ユーザーID")

                            'エラーを返す
                            Return False
                        End If
                    End If
                End If

            Next

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 登録処理
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の登録処理を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function RegUsermstData(ByRef dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数を宣言
        Dim sqlEXTM0101 As New SqlEXTM0101        'Sqlクラスをインスタンス化
        Dim Cn As New NpgsqlConnection(DbString)  'コネクション
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション
        Dim Table As DataTable                    'データテーブル

        Table = DirectCast(dataEXTM0101.PropVwList.DataSource, DataTable)

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            For j As Integer = 0 To dataEXTM0101.PropVwList.Sheets(0).RowCount - 1
                '変更区分がチェックされている場合
                If dataEXTM0101.PropVwList.Sheets(0).Cells(j, 0).Value = True Then
                    '隠し列（＝更新区分）に１が立っている場合
                    If dataEXTM0101.PropVwList.Sheets(0).Cells(j, 11).Value = 1 Then
                        'ユーザーIDが存在する = 更新(Update)
                        Update(dataEXTM0101, j, Tsx, Cn)
                    Else
                        'ユーザーIDが存在しない = 新規登録(Insert)
                        Insert(dataEXTM0101, j, Tsx, Cn)
                    End If
                End If
            Next

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_I0002, "登録")

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0101_E0000)
            Return False

        Finally
            
            'トランザクションを破棄
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            'コネクションを破棄
            Cn.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 更新
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の登録処理を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function Update(ByRef dataEXTM0101 As DataEXTM0101, _
                           ByRef j As Integer, _
                           ByRef Tsx As NpgsqlTransaction, _
                           ByRef Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            'ユーザーマスタ新規登録（INSERT）用SQLを作成
            If SqlEXTM0101.SetUpdateUsermstSql(Cn, Cmd, j, dataEXTM0101) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ユーザーマスタメンテ更新", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

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
            puErrMsg = String.Format(M0101_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function
    ''' <summary>
    ''' 新規登録
    ''' <paramref name="dataEXTM0101">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の登録処理を行う
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function Insert(ByRef dataEXTM0101 As DataEXTM0101, _
                           ByRef j As Integer, _
                           ByRef Tsx As NpgsqlTransaction, _
                           ByRef Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            'ユーザーマスタ新規登録（INSERT）用SQLを作成
            If SqlEXTM0101.SetInsertUsermstSql(Cn, Cmd, j, dataEXTM0101) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ユーザーマスタメンテ新規登録", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

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
            puErrMsg = String.Format(M0101_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function
End Class
