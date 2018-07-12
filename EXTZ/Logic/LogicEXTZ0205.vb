Imports Common
Imports CommonEXT
Imports Npgsql

Public Class LogicEXTZ0205

    Private sqlEXTZ0205 As New SqlEXTZ0205          'sqlクラス

    '定数宣言
    Public Const COL_SELECT As Integer = 0            '選択

    ''' <summary>
    ''' 初期表示時検索
    ''' <paramref name="dataEXTZ0205">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>初期表示時検索を行う
    ''' <para>作成情報：2015/12/11 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitSearch(ByRef dataEXTZ0205 As DataEXTZ0205) As Boolean

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
            If sqlEXTZ0205.SetSelectProjectSql(Adapter, Cn, dataEXTZ0205) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXシアターEXASプロジェクト設定初期表示時検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            'データテーブルに取得データをセット
            dataEXTZ0205.PropDtResult = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' プロジェクト検索
    ''' </summary>
    ''' <param name="dataEXTZ0205"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProject(ByRef dataEXTZ0205 As DataEXTZ0205) As Boolean
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
            If sqlEXTZ0205.searchProject(Adapter, Cn, dataEXTZ0205) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "プロジェクト検索", Nothing, Adapter.SelectCommand)
            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            dataEXTZ0205.PropDtResult = Table
            Cn.Close()
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
    ''' スプレッドシートチェックボックスクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0205">[IN/OUT]プロジェクト一覧画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッドシートのチェックボックス状態を制御する
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickResultCellMain(ByRef dataEXTZ0205 As DataEXTZ0205) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスを制御する
            If SetCheck(dataEXTZ0205) = False Then
                Return False
            End If


            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' チェックボックス制御処理
    ''' <paramref name="dataEXTZ0205">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheck(ByRef dataEXTZ0205 As DataEXTZ0205) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTZ0205.PropResult.ActiveSheet

                '選択行にチェックをつけ、それ以外はチェックを外す
                For i As Integer = 0 To .RowCount - 1
                    If .GetValue(i, COL_SELECT) = True Then
                        '選択行がすでにチェック済みの場合はチェックを外す
                        .SetValue(i, COL_SELECT, False)

                        '選択行をチェックする
                    Else
                        .SetValue(i, COL_SELECT, False)
                        If i = dataEXTZ0205.PropIntCheckRow Then
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
            Return False
        End Try

    End Function
End Class
