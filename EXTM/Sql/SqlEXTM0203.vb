Imports Npgsql
Imports Common
Imports CommonEXT

Public Class SqlEXTM0203

    Private strGetExasTheOther As String = "SELECT FALSE, " & vbCrLf &
                                           "aite_cd, " & vbCrLf &
                                           "aite_nm, " & vbCrLf &
                                           "aite_nm_kana, " & vbCrLf &
                                           "post_bango, " & vbCrLf &
                                           "add1 || add2 || add3 AS add, " & vbCrLf &
                                           "tel_bango, " & vbCrLf &
                                           "fax_bango, " & vbCrLf &
                                           "CASE kosin_kbn " & vbCrLf &
                                           "WHEN '3' THEN '×' " & vbCrLf &
                                           "ELSE '' " & vbCrLf &
                                           "END stop_use " & vbCrLf &
                                           "FROM aitesaki_mst " & vbCrLf

    ''' <summary>
    ''' EXAS相手先情報取得SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTM0203">[IN]EXAS相手先一覧画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS相手先情報取得SQLの作成・設定処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSearchExasTheOther(ByRef Adapter As NpgsqlDataAdapter, _
                                          ByVal Cn As NpgsqlConnection, _
                                          ByVal dataEXTM0203 As DataEXTM0203) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""
        Dim strWhere As String = ""

        Try
            '検索条件設定
            '相手先コードに入力がある場合
            If dataEXTM0203.PropStrExasCode_Search IsNot Nothing And dataEXTM0203.PropStrExasCode_Search IsNot "" Then
                strWhere = "WHERE aite_cd = :AiteCd"
                '相手先名にも入力がある場合
                If dataEXTM0203.PropStrExasName_Search IsNot Nothing And dataEXTM0203.PropStrExasName_Search IsNot "" Then
                    strWhere &= vbCrLf & "AND aite_nm LIKE :AiteNm"
                End If
                '相手先コードに入力がなく、相手先名に入力がある場合
            ElseIf dataEXTM0203.PropStrExasName_Search IsNot Nothing And dataEXTM0203.PropStrExasName_Search IsNot "" Then
                strWhere = "WHERE aite_nm LIKE :AiteNm"
            End If
            'order by句を記述
            strWhere &= vbCrLf & "ORDER BY aite_cd"

            'SQL文(SELECT)
            strSQL = strGetExasTheOther & strWhere

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型を設定
            Adapter.SelectCommand.Parameters.Add("AiteCd", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters.Add("AiteNm", NpgsqlTypes.NpgsqlDbType.Varchar)

            'バインド変数に値を設定
            Adapter.SelectCommand.Parameters("AiteCd").Value = dataEXTM0203.PropStrExasCode_Search
            Adapter.SelectCommand.Parameters("AiteNm").Value = "%" & dataEXTM0203.PropStrExasName_Search & "%"

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

End Class
