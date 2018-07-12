Imports Common
Imports Npgsql
Imports CommonEXT

Public Class SqlEXTZ0205
    Private commonLogicEXT As New CommonLogicEXT

    'SQL文(ログイン)
    Private strEX39S001 As String = _
                           "SELECT " & vbCrLf & _
                           "    m1.event_cd, " & vbCrLf & _
                           "    m1.event_nm, " & vbCrLf & _
                           "    m2.content_uchi_cd, " & vbCrLf & _
                           "    m2.content_uchi_nm " & vbCrLf & _
                           "FROM CONTENT_MST m1 " & vbCrLf & _
                           "LEFT JOIN CONTENT_UCHI_MST m2 " & vbCrLf & _
                           "    ON m1.event_cd = m2.content_cd " & vbCrLf & _
                           "WHERE (m1.del_flg <> '1' OR m1.del_flg is null ) " & vbCrLf & _
                           "    AND (m2.del_flg <> '1' OR m2.del_flg is null ) " & vbCrLf & _
                           "    AND :Riyobi BETWEEN m1.strat_dt AND m1.end_dt "
    ''' <summary>
    ''' プロジェクト情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0205"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/12/11 m.hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetSelectProjectSql(ByRef Adapter As NpgsqlDataAdapter, _
                                        ByRef Cn As NpgsqlConnection, _
                                        ByVal dataEXTZ0205 As DataEXTZ0205) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            strSQL = strEX39S001 & "AND m1.event_bunrui_cd = '03'"

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":Riyobi").Value = dataEXTZ0205.PropStrRiyobi
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
    ''' プロジェクト情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0205"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function searchProject(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0205 As DataEXTZ0205) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX39S001
            If String.IsNullOrEmpty(dataEXTZ0205.PropStrPrjCd) = False Then
                strSQL = strSQL + "    AND m1.event_cd = :PrjCd "
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropStrPrjNm) = False Then
                strSQL = strSQL + "    AND m1.event_nm like :PrjNm "
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropUchiCd) = False Then
                strSQL = strSQL + "    AND m2.content_uchi_cd = :UchiCd "
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropUchiNm) = False Then
                strSQL = strSQL + "    AND m2.content_uchi_nm like :UchiNm "
            End If
            strSQL = strSQL + "ORDER BY m1.event_cd, m2.content_uchi_cd "

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":Riyobi").Value = dataEXTZ0205.PropStrRiyobi
            End With

            If String.IsNullOrEmpty(dataEXTZ0205.PropStrPrjCd) = False Then
                Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":PrjCd", NpgsqlTypes.NpgsqlDbType.Varchar))
                Adapter.SelectCommand.Parameters(":PrjCd").Value = dataEXTZ0205.PropStrPrjCd
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropStrPrjNm) = False Then
                Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":PrjNm", NpgsqlTypes.NpgsqlDbType.Varchar))
                Adapter.SelectCommand.Parameters(":PrjNm").Value = "%" + dataEXTZ0205.PropStrPrjNm + "%"
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropUchiCd) = False Then
                Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":UchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
                Adapter.SelectCommand.Parameters(":UchiCd").Value = dataEXTZ0205.PropUchiCd
            End If
            If String.IsNullOrEmpty(dataEXTZ0205.PropUchiNm) = False Then
                Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":UchiNm", NpgsqlTypes.NpgsqlDbType.Varchar))
                Adapter.SelectCommand.Parameters(":UchiNm").Value = "%" + dataEXTZ0205.PropUchiNm + "%"
            End If


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
