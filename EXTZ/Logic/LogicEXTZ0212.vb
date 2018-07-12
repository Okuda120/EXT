Imports Common
Imports CommonEXT
Imports Npgsql

Public Class LogicEXTZ0212

    Private sqlEXTZ0212 As New SqlEXTZ0212          'sqlクラス

    '定数宣言========================
    '列番号
    Public Const COL_HRSELECT As Integer = 2          '時間貸し料金選択
    Public Const COL_LRSELECT As Integer = 4          '日貸し料金選択
    Public Const COL_SELECT As Integer = 2            '倍率選択

    ''' <summary>
    ''' 分類取得
    ''' </summary>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBunrui(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean
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
            If sqlEXTZ0212.selectBunrui(Adapter, Cn, dataEXTZ0212) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "分類取得", Nothing, Adapter.SelectCommand)
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
    ''' 利用料金取得
    ''' </summary>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRyokin(ByRef dataEXTZ0212 As DataEXTZ0212, ByVal strBunrui As String) As Boolean
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
            If sqlEXTZ0212.selectRyokin(Adapter, Cn, dataEXTZ0212, strBunrui) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用料金取得", Nothing, Adapter.SelectCommand)
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
    ''' 倍率取得
    ''' </summary>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBairitu(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean
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
            If sqlEXTZ0212.selectBairitu(Adapter, Cn, dataEXTZ0212) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "倍率取得", Nothing, Adapter.SelectCommand)
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
    ''' 料金一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0212">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>料金一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickRyokinCellMain(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスを制御する
            If SetCheckRyokin(dataEXTZ0212) = False Then
                Return False
            End If


            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'puErrMsg = HBK_E001 & ex.Message
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 倍率一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0212">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>倍率一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickBairituCellMain(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスを制御する
            If SetCheckBairitu(dataEXTZ0212) = False Then
                Return False
            End If


            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'puErrMsg = HBK_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金一覧チェックボックス制御処理
    ''' <paramref name="dataEXTZ0212">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckRyokin(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTZ0212.PropRyokin.ActiveSheet


                '選択したセルにはチェックをつけ、それ以外はチェックを外す
                For i As Integer = 0 To .RowCount - 1
                    If dataEXTZ0212.PropIntCheckCol = COL_HRSELECT Then
                        If .GetValue(i, COL_HRSELECT) = True Then
                            '選択したセルがすでにチェック済みの場合はチェックを外す
                            .SetValue(i, COL_HRSELECT, False)

                            '選択したセルにチェックをつける
                        Else
                            .SetValue(i, COL_HRSELECT, False)
                            .SetValue(i, COL_LRSELECT, False)
                            If i = dataEXTZ0212.PropIntCheckRow Then
                                .SetValue(i, COL_HRSELECT, True)
                            End If
                        End If

                    ElseIf dataEXTZ0212.PropIntCheckCol = COL_LRSELECT Then

                        '選択したセルにはチェックをつけ、それ以外はチェックを外す
                        If .GetValue(i, COL_LRSELECT) = True Then
                            '選択したセルがすでにチェック済みの場合はチェックを外す
                            .SetValue(i, COL_LRSELECT, False)

                            '選択したセルにチェックをつける
                        Else
                            .SetValue(i, COL_HRSELECT, False)
                            .SetValue(i, COL_LRSELECT, False)
                            If i = dataEXTZ0212.PropIntCheckRow Then
                                .SetValue(i, COL_LRSELECT, True)
                            End If
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
            'puErrMsg = HBK_E001 & ex.Message
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 倍率一覧チェックボックス制御処理
    ''' <paramref name="dataEXTZ0212">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckBairitu(ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTZ0212.PropBairitu.ActiveSheet

                '選択行（セット）にはチェックをつけ、それ以外はチェックを外す
                For i As Integer = 0 To .RowCount - 1
                    If .GetValue(i, COL_SELECT) = True Then
                        '選択したセルがすでにチェック済みの場合はチェックを外す
                        .SetValue(i, COL_SELECT, False)
                    Else
                        .SetValue(i, COL_SELECT, False)
                        If i = dataEXTZ0212.PropIntCheckRow Then
                            .SetValue(i, COL_SELECT, True)
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
            'puErrMsg = HBK_E001 & ex.Message
            Return False
        End Try

    End Function
End Class

