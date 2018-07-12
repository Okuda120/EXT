Imports Common
Imports Npgsql
Imports System.Text

Public Class SqlEXTA0101

    'SQL文(ログイン)
    Private strEX01S001 As String = _
                           "select " & vbCrLf & _
                               "USER_NM, " & vbCrLf & _
                               "MAIL, " & vbCrLf & _
                               "BUSHO_CD, " & vbCrLf & _
                               "SHONIN_FLG, " & vbCrLf & _
                               "MST_FLG " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "USER_MST " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "USER_CD = :UserId " & vbCrLf & _
                                "and PW = :Password " & vbCrLf & _
                                "and STS = '0' "

    'SQL文(ログイン)
    '2016.09.20 m.hayabuchi MOD START↓ 請求依頼データ代行処理対応
    '' 2016.01.18 ADD START↓ y.morooka グループ請求対応
    ''Private strEXSYS_PROP As String = _
    ''                       "select " & vbCrLf & _
    ''                           "TEIIN_A, " & vbCrLf & _
    ''                           "TEIIN_B, " & vbCrLf & _
    ''                           "KARI_JIZEN_TUTI " & vbCrLf & _
    ''                       "FROM " & vbCrLf & _
    ''                           "SYSTEM_MST "
    'Private strEXSYS_PROP As String = _
    '                   "select " & vbCrLf & _
    '                       "TEIIN_A, " & vbCrLf & _
    '                       "TEIIN_B, " & vbCrLf & _
    '                       "KARI_JIZEN_TUTI, " & vbCrLf & _
    '                       "GROUP_SEIKYU " & vbCrLf & _
    '                   "FROM " & vbCrLf & _
    '                       "SYSTEM_MST "
    ' 2016.01.18 ADD END↑ y.morooka グループ請求対応
    Private strEXSYS_PROP As String = _
                   "select " & vbCrLf & _
                       "TEIIN_A, " & vbCrLf & _
                       "TEIIN_B, " & vbCrLf & _
                       "KARI_JIZEN_TUTI, " & vbCrLf & _
                       "GROUP_SEIKYU, " & vbCrLf & _
                       "DAIKO_TANTO_CD, " & vbCrLf & _
                       "DAIKO_BUSHO_CD " & vbCrLf & _
                   "FROM " & vbCrLf & _
                       "SYSTEM_MST "
    ' 2016.09.20 m.hayabuchi MOD END↑ 請求依頼データ代行処理対応

    ''' <summary>
    ''' ＳＱＬ文作成(SELECT)
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTA0101"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectLoginData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTA0101 As DataEXTA0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX01S001

            'Dim cmd = New NpgsqlCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Password", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":UserId").Value = dataEXTA0101.PropTxtUserId.Text                   'ユーザID
                .Parameters(":Password").Value = dataEXTA0101.PropTxtPassword.Text                 'パスワード
            End With

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)

            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' システムプロパティ取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>システムプロパティ
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectSystemProperty(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEXSYS_PROP

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)

            '例外処理
            Return False
        End Try
    End Function

End Class
