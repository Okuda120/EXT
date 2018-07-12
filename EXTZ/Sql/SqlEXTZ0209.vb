Imports Common
Imports CommonEXT
Imports Npgsql

Public Class sqlEXTZ0209

    'SQL文(付帯分類)
    Private strEX34S001 As String = _
                           "SELECT " & vbCrLf & _
                           "    shisetu_kbn, " & vbCrLf & _
                           "    bunrui_cd, " & vbCrLf & _
                           "    bunrui_nm, " & vbCrLf & _
                           "    notax_flg, " & vbCrLf & _
                           "    kamoku_cd, " & vbCrLf & _
                           "    saimoku_cd, " & vbCrLf & _
                           "    uchi_cd, " & vbCrLf & _
                           "    shosai_cd, " & vbCrLf & _
                           "    karikamoku_cd, " & vbCrLf & _
                           "    kari_saimoku_cd, " & vbCrLf & _
                           "    kari_uchi_cd, " & vbCrLf & _
                           "    kari_shosai_cd, " & vbCrLf & _
                           "    sts, " & vbCrLf & _
                           "    sort " & vbCrLf & _
                           "FROM " & vbCrLf & _
                           "    fbunrui_mst " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                           "    shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "AND :Riyobi BETWEEN kikan_from AND kikan_to " & vbCrLf & _
                           "AND sts = '0' " & vbCrLf & _
                           "ORDER BY " & vbCrLf & _
                           "    sort "

    Private strEX34S002 As String = _
                            "SELECT " & vbCrLf & _
                            "    m1.bunrui_cd as bunrui_cd, " & vbCrLf & _
                            "    m1.futai_cd, " & vbCrLf & _
                            "    m1.futai_nm, " & vbCrLf & _
                            "    m1.tanka as tanka, " & vbCrLf & _
                            "    m1.tani as tani, " & vbCrLf & _
                            "    m1.sort as sort, " & vbCrLf & _
                            "    m2.bunrui_nm as bunrui_nm, " & vbCrLf & _
                            "    m2.shukei_grp, " & vbCrLf & _
                            "    m2.kamoku_cd, " & vbCrLf & _
                            "    m2.saimoku_cd, " & vbCrLf & _
                            "    m2.uchi_cd, " & vbCrLf & _
                            "    m2.shosai_cd, " & vbCrLf & _
                            "    m2.karikamoku_cd, " & vbCrLf & _
                            "    m2.kari_saimoku_cd, " & vbCrLf & _
                            "    m2.kari_uchi_cd, " & vbCrLf & _
                            "    m2.kari_shosai_cd, " & vbCrLf & _
                            "    m2.notax_flg, " & vbCrLf & _
                            "    m3.kamoku_nm, " & vbCrLf & _
                            "    m3.saimoku_nm, " & vbCrLf & _
                            "    m3.uchi_nm, " & vbCrLf & _
                            "    m3.shosai_nm " & vbCrLf & _
                            "    ,0 " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    futai_mst m1 " & vbCrLf & _
                            "LEFT JOIN fbunrui_mst m2 " & vbCrLf & _
                            "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf & _
                            "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                            "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf & _
                            "    AND m2.sts = '0' " & vbCrLf & _
                            "LEFT JOIN kamoku_mst m3 " & vbCrLf & _
                            "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf & _
                            "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf & _
                            "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf & _
                            "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf & _
                            "    AND m3.sts = '0' " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "    m1.sts = '0' " & vbCrLf & _
                            "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf & _
                            "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                            "ORDER BY " & vbCrLf & _
                            "    m1.sort "

    '消費税の取得 2015.11.13 ADD h.hagiwara
    Private strEX34S003 As String = _
                            "SELECT " & vbCrLf & _
                            "    tax_ritu " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    tax_mst " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            " :Riyobi BETWEEN taxs_dt AND taxe_dt "

    ''' <summary>
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0209"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectBunrui(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection,
                                       ByRef dataEXTZ0209 As DataEXTZ0209) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Dim Table As New DataTable
            Cmd.Connection = Cn
            Cmd.CommandText = strEX34S001
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTZ0209.PropStrShisetu
                .Parameters(1).Value = dataEXTZ0209.PropStrRiyobi
            End With

            Adapter.SelectCommand = Cmd

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
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0209"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectFutaiMst(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection,
                                       ByRef dataEXTZ0209 As DataEXTZ0209) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Dim Table As New DataTable
            Cmd.Connection = Cn
            Cmd.CommandText = strEX34S002
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTZ0209.PropStrShisetu
                .Parameters(1).Value = dataEXTZ0209.PropStrRiyobi
            End With

            Adapter.SelectCommand = Cmd

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
    ''' 消費税取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>システムプロパティ
    ''' <para>作成情報：2015.11.13 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectTax(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal Riyobi As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX34S003
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            Adapter.SelectCommand.Parameters(0).Value = Riyobi
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
