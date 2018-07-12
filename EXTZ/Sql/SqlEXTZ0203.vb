Imports Common
Imports Npgsql
Imports CommonEXT

Public Class SqlEXTZ0203

    Private commonLogicEXT As New CommonLogicEXT

    'SQL文(ログイン)
    '祝祭日マスタ存在チェック
    Private strHolidayCntSql As String = "SELECT  COUNT(*) " & vbCrLf &
                                         "FROM HOLIDAY_MST " & vbCrLf &
                                         "WHERE HOLIDAY_DT = :HOLIDAY_DT "



    ''' <summary>
    ''' 祝祭日マスタを取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="CehckDay">[IN]取得対象日付</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日マスタ取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetHolidayMst(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal CehckDay As string) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strHolidayCntSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("HOLIDAY_DT").Value = CehckDay

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

End Class
