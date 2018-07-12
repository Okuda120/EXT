Imports Common
Imports CommonEXT
Imports Npgsql
Public Class LogicEXTZ0202

    '変数宣言
    Private sqlEXTZ0202 As New SqlEXTZ0202          'sqlクラス

    ''' <summary>
    ''' 予約受付制御登録チェック
    ''' </summary>
    ''' <param name="strTorokuKbn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <param name="lstRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckRiyobi(ByVal strTorokuKbn As String, ByVal dataRiyobi As CommonDataRiyobi, ByRef lstRiyobi As ArrayList) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0202.cntRiyobi(Adapter, Cn, dataRiyobi) = False Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Adapter.SelectCommand)
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得", Nothing, Adapter.SelectCommand)
            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Nothing)
                Return False
            End If

            '検索結果とlstRiyobiの合計件数でチェックする
            Dim cnt As Integer = Table.Rows(0).Item(0)
            For Each riyobi As CommonDataRiyobi In lstRiyobi
                If dataRiyobi.PropStrYoyakuDt = riyobi.PropStrYoyakuDt And dataRiyobi.PropStrStudioKbn = riyobi.PropStrStudioKbn Then
                    '同じ日付は登録させない
                    Return False
                End If
            Next
            Cn.Close()

            'チェック処理
            If strTorokuKbn = "0" Or strTorokuKbn = "1" Then
                If dataRiyobi.PropStrShisetuKbn = "1" And 2 <= cnt Then
                    Return False
                ElseIf dataRiyobi.PropStrShisetuKbn = "2" And 3 <= cnt Then
                    Return False
                End If
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 予約受付制御登録チェック
    ''' </summary>
    ''' <param name="strTorokuKbn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckRiyobiRegister(ByVal strTorokuKbn As String, ByVal dataRiyobi As CommonDataRiyobi) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0202.cntRiyobi(Adapter, Cn, dataRiyobi) = False Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Adapter.SelectCommand)
                Return False
            End If

            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Nothing)
                Return False
            End If

            '検索結果とlstRiyobiの合計件数でチェックする
            Dim cnt As Integer = Table.Rows(0).Item(0)
            Cn.Close()

            '利用形態
            Dim keitai As String = ""
            If dataRiyobi.PropStrRiyoKeitai = Nothing = False Then
                If String.IsNullOrEmpty(dataRiyobi.PropStrRiyoKeitai) = False Then
                    keitai = dataRiyobi.PropStrRiyoKeitai
                End If
            End If

            'チェック処理
            If strTorokuKbn = "0" Or strTorokuKbn = "1" Then
                If dataRiyobi.PropStrShisetuKbn = SHISETU_KBN_THEATER And 2 <= cnt Then
                    Return False
                ElseIf dataRiyobi.PropStrShisetuKbn = SHISETU_KBN_STUDIO And keitai = RIYOKEITAI_LOCKOUT And 1 <= cnt Then
                    Return False
                ElseIf dataRiyobi.PropStrShisetuKbn = SHISETU_KBN_STUDIO And 3 <= cnt Then
                    Return False
                End If
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 予約受付制御削除(施設区分指定なし)
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の全削除
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteYoyakuCtlData() As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約受付制御の削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteYoyakuCtl(Cmd, Cn, SHISETU_KBN_BLANK)
            Cmd.ExecuteNonQuery()

            'エラーがなければコミットする
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約受付制御削除(施設区分指定)
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の全削除
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteYoyakuCtlShisetuData(ByVal strShisetuCd As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約受付制御の削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteYoyakuCtl(Cmd, Cn, strShisetuCd)
            Cmd.ExecuteNonQuery()
            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約受付制御登録
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の削除
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function InsertYoyakuCtlData(ByVal aryRiyobiList As ArrayList) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            For Each dataRiyobi As CommonDataRiyobi In aryRiyobiList
                sqlEXTZ0202.insertYoyakuCtl(Cmd, Cn, dataRiyobi)
                Cmd.ExecuteNonQuery()
            Next

            'エラーがなければコミットする
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約受付制御情報単体削除
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約受付制御情報単体削除
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteYoyakuCtlData(ByRef dataRiyobi As CommonDataRiyobi) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約制御データの削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteLineYoyakuCtl(Cmd, Cn, CommonEXT.PropComStrUserId, dataRiyobi.PropStrYoyakuDt)
            Cmd.ExecuteNonQuery()

            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御登録チェック
    ''' </summary>
    ''' <param name="strTorokuKbn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <param name="lstRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckCancelWait(ByVal strTorokuKbn As String, ByVal dataRiyobi As CommonDataCancel, ByRef lstRiyobi As ArrayList) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0202.cntCancelWait(Adapter, Cn, dataRiyobi) = False Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Adapter.SelectCommand)
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得", Nothing, Adapter.SelectCommand)
            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Nothing)
                Return False
            End If

            '検索結果とlstRiyobiの合計件数でチェックする
            Dim cnt As Integer = Table.Rows(0).Item(0)
            For Each riyobi As CommonDataCancel In lstRiyobi
                If dataRiyobi.PropStrCancelDt = riyobi.PropStrCancelDt And dataRiyobi.PropStrStudioKbn = riyobi.PropStrStudioKbn Then
                    '同じ日付は登録させない
                    Return False
                End If
            Next
            Cn.Close()

            'チェック処理
            If 3 <= cnt Then
                Return False
            End If
            
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御登録チェック
    ''' </summary>
    ''' <param name="strTorokuKbn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckCancelWaitRegister(ByVal strTorokuKbn As String, ByVal dataRiyobi As CommonDataCancel) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0202.cntCancelWait(Adapter, Cn, dataRiyobi) = False Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Adapter.SelectCommand)
                Return False
            End If

            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "予約情報カウント失敗", Nothing, Nothing)
                Return False
            End If

            '検索結果とlstRiyobiの合計件数でチェックする
            Dim cnt As Integer = Table.Rows(0).Item(0)
            Cn.Close()

            'チェック処理
            If 3 <= cnt Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御削除(施設区分指定なし)
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の全削除
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteCancelCtlData() As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'キャンセル待ち受付制御の削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteCancelCtl(Cmd, Cn, SHISETU_KBN_BLANK)
            Cmd.ExecuteNonQuery()

            'エラーがなければコミットする
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御削除(施設区分指定)
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の全削除
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteCancelCtlShisetuData(ByVal strShisetuCd As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'キャンセル待ち受付制御の削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteCancelCtl(Cmd, Cn, strShisetuCd)
            Cmd.ExecuteNonQuery()
            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御登録
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の削除
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function InsertCancelCtlData(ByVal aryRiyobiList As ArrayList) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            For Each dataRiyobi As CommonDataCancel In aryRiyobiList
                If dataRiyobi.PropStrKibobiKbn = "1" Then
                    sqlEXTZ0202.insertCancelCtl(Cmd, Cn, dataRiyobi)
                    Cmd.ExecuteNonQuery()
                End If
            Next

            'エラーがなければコミットする
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御情報単体削除
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>キャンセル待ち受付制御情報単体削除
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteCancelCtlData(ByRef dataRiyobi As CommonDataCancel) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約制御データの削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteLineCancelCtl(Cmd, Cn, CommonEXT.PropComStrUserId, dataRiyobi.PropStrCancelDt)
            Cmd.ExecuteNonQuery()

            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

End Class
