Imports Common
Imports Npgsql
Imports CommonEXT

Public Class LogicEXTM0203

    'インスタンス生成
    Private sqlEXTM0203 As New SqlEXTM0203

    '定数宣言
    Public Const COL_SELECT As Integer = 0  '選択

    ''' <summary>
    ''' EXAS相手先情報検索メイン処理
    ''' </summary>
    ''' <param name="dataEXTM0203">[IN/OUT]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS相手先情報検索処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ExasSearchMain(ByRef dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '検索条件取得
        With dataEXTM0203
            .PropStrExasCode_Search = .PropTxtExasCode.Text     'EXAS相手先コード
            .PropStrExasName_Search = .PropTxtExasName.Text     'EXAS相手先名
        End With

        'EXAS相手先データを取得
        If GetExasTheOtherData(dataEXTM0203) = False Then
            Return False
        End If

        'スプレッドシート表示設定
        If SetDisplay(dataEXTM0203) = False Then
            Return False
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常処理終了
        Return True

    End Function

    ''' <summary>
    ''' スプレッドシート表示設定
    ''' </summary>
    ''' <param name="dataEXTM0203">[IN/OUT]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>スプレッドシート表示設定を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetDisplay(ByRef dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTM0203

                'スプレッドシートにデータを設定
                .PropVwResult.Sheets(0).DataSource = .PropDtResult

                '折り返して表示するセルを設定
                Dim WrapCell As New FarPoint.Win.Spread.CellType.TextCellType
                WrapCell.WordWrap = True

                'スプレッドシートの行設定
                .PropVwResult.Sheets(0).Columns(0).Width = 39 '選択チェックボックス
                .PropVwResult.Sheets(0).Columns(1).Width = 84 '相手先コード
                .PropVwResult.Sheets(0).Columns(2).Width = 148 '相手先名
                .PropVwResult.Sheets(0).Columns(3).Width = 148 '相手先名カナ
                .PropVwResult.Sheets(0).Columns(4).Width = 60 '郵便番号
                .PropVwResult.Sheets(0).Columns(5).Width = 229 '住所１＋２＋３
                .PropVwResult.Sheets(0).Columns(6).Width = 80 '電話番号
                .PropVwResult.Sheets(0).Columns(7).Width = 80 'FAX
                .PropVwResult.Sheets(0).Columns(8).Width = 68 '使用停止

                'セル型設定
                .PropVwResult.Sheets(0).Columns(2).CellType = WrapCell '相手先名(折り返し表示可に設定)
                .PropVwResult.Sheets(0).Columns(3).CellType = WrapCell '相手先名カナ(折り返し表示可に設定
                .PropVwResult.Sheets(0).Columns(5).CellType = WrapCell '住所１＋２＋３(折り返し表示可に設定)
                .PropVwResult.Sheets(0).Columns(6).CellType = WrapCell '電話番号(折り返し表示可に設定)
                .PropVwResult.Sheets(0).Columns(7).CellType = WrapCell 'FAX(折り返し表示可に設定)
                .PropVwResult.Sheets(0).Columns(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center '使用停止(中央表示)


                '使用停止の相手先は選択不可にする
                'データ行数取得
                Dim intCountRow As Integer = .PropVwResult.Sheets(0).RowCount
                For i As Integer = 0 To intCountRow - 1
                    '住所が複数行でも表示されるように、行の高さを調整する
                    .PropVwResult.Sheets(0).Rows(i).Height = .PropVwResult.Sheets(0).Rows(i).GetPreferredHeight
                    '使用停止の行のチェックボックスを選択不可にする
                    If .PropVwResult.Sheets(0).Cells(i, 8).Value = "×" Then
                        .PropVwResult.Sheets(0).Cells(i, 0).Locked = True
                    End If
                    'メニューから遷移時は選択区分を選択不可にする
                    If dataEXTM0203.PropStrMenuID = "" Then
                        .PropVwResult.Sheets(0).Cells(i, 0).Locked = True
                    End If
                Next

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
    ''' <param name="dataEXTM0203">[IN/OUT]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>入力チェック処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InputCheck(ByRef dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'チェックされた行数
        Dim intCheckSum As Integer = 0

        Try
            With dataEXTM0203

                'データ行数取得
                Dim intCountRow As Integer = .PropVwResult.Sheets(0).RowCount
                'ループしてチェックされた行数を取得する
                For i As Integer = 0 To intCountRow - 1
                    If .PropVwResult.Sheets(0).Cells(i, 0).Value = True Then
                        intCheckSum = intCheckSum + 1
                        .PropIntCheckRow = i
                    End If
                Next

            End With

            If intCheckSum = 0 Then
                '１件もチェックされてない場合、エラー
                puErrMsg = String.Format(EXTM0203_E0002, "相手先")
                Return False
            ElseIf intCheckSum > 1 Then
                '２件以上チェックされている場合、エラー
                puErrMsg = String.Format(EXTM0203_E2035, "相手先")
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
    ''' EXAS相手先情報取得処理
    ''' </summary>
    ''' <param name="dataEXTM0203">[IN/OUT]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS相手先情報取得処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetExasTheOtherData(ByRef dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtResult As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            'EXAS相手先情報取得SQLを設定する
            If sqlEXTM0203.SetSearchExasTheOther(Adapter, Cn, dataEXTM0203) = False Then
                Return False
            End If

            'EXAS相手先情報を取得
            Adapter.Fill(dtResult)
            dataEXTM0203.PropDtResult = dtResult

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
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtResult.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTM0203">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickVwCellMain(ByRef dataEXTM0203 As DataEXTM0203) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスをラジオボタンのように制御する
            If SetCheckAsRadio(dataEXTM0203) = False Then
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
    ''' <paramref name="dataHEXTM0203">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckAsRadio(ByRef dataEXTM0203 As DataEXTM0203) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTM0203

                'ON状態のチェックボックスを選択した場合、自分自身のチェックを外す
                'OFF状態のチェックボックスを選択した場合、選択行にはチェックをつけ、それ以外にはチェックを外す
                For i As Integer = 0 To .PropVwResult.ActiveSheet.RowCount - 1

                    'ON状態のチェックボックスを選択した場合
                    If .PropVwResult.ActiveSheet.GetValue(i, COL_SELECT) = True Then
                        .PropVwResult.ActiveSheet.SetValue(i, COL_SELECT, False)
                    Else
                        'チェックを外す
                        .PropVwResult.ActiveSheet.SetValue(i, COL_SELECT, False)

                        '選択行にチェックをつける
                        If i = .PropIntCheckRow Then
                            .PropVwResult.ActiveSheet.SetValue(i, COL_SELECT, True)
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
    ''' 入力チェック処理
    ''' </summary>
    ''' <param name="dataEXTM0203">[IN/OUT]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>検索時入力チェック処理を行う
    ''' <para>作成情報： 2015.12.03 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SearchInputCheck(ByRef dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim strSetflg As String = "0"

        Try
            With dataEXTM0203

                ' 検索条件：相手先コードの入力判定 
                If .PropTxtExasCode.Text <> "" Then
                    strSetflg = "1"
                End If
                ' 検索条件：相手先名の入力判定 
                If .PropTxtExasName.Text <> "" Then
                    strSetflg = "1"
                End If

                If strSetflg  = "0" Then
                    puErrMsg = EXTM0203_E2036
                    Return False
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

End Class
