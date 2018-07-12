Imports Common
Imports CommonEXT
Imports Npgsql

Public Class LogicEXTZ0203

    Private sqlEXTZ0203 As New SqlEXTZ0203          'sqlクラス

    ''' <summary>
    ''' 入金予定日の曜日判定
    ''' </summary>
    ''' <param name="CehckDay">[IN/OUT]算出元日付</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>請求日から自動算出した入金予定日が土日祝祭日の場合前営業日にする
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function WeekDayCheck(ByRef CehckDay As String) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim dtNyukin As DateTime
        Dim intweek As Integer

        dtNyukin = DateTime.Parse(CehckDay)

        intweek = Weekday(dtNyukin)
        If intweek = 1 Or intweek = 7 Then                      ' 土日の場合判定
            Return False
        Else
            If GetSelectHolidayMst(CehckDay) = True Then        ' 祝祭日マスタに日付が存在するか判定する
                Return False
            End If
        End If

        Return True

    End Function

    ''' <summary>
    ''' 祝祭日マスタの取得準備
    ''' <paramref name="CehckDay">対象日付</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>サーバ日時取得
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetSelectHolidayMst(ByRef CehckDay As String) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる

        Try
            'コネクションを開く
            Cn.Open()

            '新規Inc番号、システム日付取得（SELECT）
            If GetSelectHoliday(Cn, CehckDay) = False Then
                Return False
            End If

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            Return False
        Finally
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 祝祭日マスタの取得
    ''' </summary>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="CehckDay">[IN/OUT]判定対象日付</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>祝祭日を取得する
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetSelectHoliday(ByVal Cn As NpgsqlConnection, _
                               ByRef CehckDay As string) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlDataAdapter
        Dim dtResult As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTZ0203.GetHolidayMst(Adapter, Cn, CehckDay) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "祝祭日", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtResult)

            'データが取得できなかった場合は
            If dtResult.Rows.Count = 0 Then
                Return False
            Else
                If dtResult.Rows(0).Item(0) = 0 Then
                    Return False
                End If
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function

End Class
