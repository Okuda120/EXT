Imports Common
Imports CommonEXT
Imports Npgsql

Public Class sqlEXTZ0209

    'SQL文(付帯分類)
    ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
    Private strEX34S001 As String =
                           "SELECT " & vbCrLf &
                           "    m1.shisetu_kbn, " & vbCrLf &
                           "    m1.bunrui_cd, " & vbCrLf &
                           "    m1.bunrui_nm, " & vbCrLf &
                           "    m1.notax_flg, " & vbCrLf &
                           "    m1.tax_kbn, " & vbCrLf &
                           "    m2.tax_rate, " & vbCrLf &
                           "    m1.kamoku_cd, " & vbCrLf &
                           "    m1.saimoku_cd, " & vbCrLf &
                           "    m1.uchi_cd, " & vbCrLf &
                           "    m1.shosai_cd, " & vbCrLf &
                           "    m1.karikamoku_cd, " & vbCrLf &
                           "    m1.kari_saimoku_cd, " & vbCrLf &
                           "    m1.kari_uchi_cd, " & vbCrLf &
                           "    m1.kari_shosai_cd, " & vbCrLf &
                           "    m1.sts, " & vbCrLf &
                           "    m1.sort " & vbCrLf &
                           "FROM " & vbCrLf &
                           "    fbunrui_mst m1" & vbCrLf &
                           "LEFT JOIN (" & vbCrLf &
                           "    SELECT " & vbCrLf &
                           "        tax_kbn, " & vbCrLf &
                           "        CASE " & vbCrLf &
                           "            WHEN tax_kbn = 1 THEN tax_ritu " & vbCrLf &
                           "            WHEN tax_kbn = 2 THEN reduced_rate " & vbCrLf &
                           "	        WHEN tax_kbn = 3 THEN untaxed_rate " & vbCrLf &
                           "	        WHEN tax_kbn = 4 THEN tax_free " & vbCrLf &
                           "            WHEN tax_kbn = 5 THEN tax_exemption " & vbCrLf &
                           "	        WHEN tax_kbn = 6 THEN tax_old1 " & vbCrLf &
                           "            WHEN tax_kbn = 7 THEN tax_old2 " & vbCrLf &
                           "            WHEN tax_kbn = 8 THEN tax_spare1 " & vbCrLf &
                           "            WHEN tax_kbn = 9 THEN tax_spare2 " & vbCrLf &
                           "            WHEN tax_kbn = 10 THEN tax_spare3 " & vbCrLf &
                           "	        WHEN tax_kbn = 11 THEN tax_spare4 " & vbCrLf &
                           "            WHEN tax_kbn = 12 THEN tax_spare5 " & vbCrLf &
                           "        END as tax_rate " & vbCrLf &
                           "    FROM " & vbCrLf &
                           "        tax_mst " & vbCrLf &
                           "        CROSS JOIN generate_series(1,12) As s(tax_kbn) " & vbCrLf &
                           "    WHERE " & vbCrLf &
                           "        :Riyobi BETWEEN taxs_dt AND taxe_dt" & vbCrLf &
                           "    ) m2 " & vbCrLf &
                           "    ON m1.tax_kbn = m2.tax_kbn " & vbCrLf &
                           "WHERE " & vbCrLf &
                           "    m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                           "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
                           "AND m1.sts = '0' " & vbCrLf &
                           "ORDER BY " & vbCrLf &
                           "    m1.sort "

    ' 2019/06/11 軽減税率対応 変更 Start E.Okuda@Compass
    ' 税率追加
    'Private strEX34S001 As String =
    '                       "SELECT " & vbCrLf &
    '                       "    m1.shisetu_kbn, " & vbCrLf &
    '                       "    m1.bunrui_cd, " & vbCrLf &
    '                       "    m1.bunrui_nm, " & vbCrLf &
    '                       "    m1.notax_flg, " & vbCrLf &
    '                       "    m1.zeiritsu, " & vbCrLf &
    '                       "    m1.kamoku_cd, " & vbCrLf &
    '                       "    m1.saimoku_cd, " & vbCrLf &
    '                       "    m1.uchi_cd, " & vbCrLf &
    '                       "    m1.shosai_cd, " & vbCrLf &
    '                       "    m1.karikamoku_cd, " & vbCrLf &
    '                       "    m1.kari_saimoku_cd, " & vbCrLf &
    '                       "    m1.kari_uchi_cd, " & vbCrLf &
    '                       "    m1.kari_shosai_cd, " & vbCrLf &
    '                       "    m1.sts, " & vbCrLf &
    '                       "    m1.sort " & vbCrLf &
    '                       "FROM " & vbCrLf &
    '                       "    fbunrui_mst m1" & vbCrLf &
    '                        "LEFT JOIN tax_mst m2 " & vbCrLf &
    '                        "    ON :Riyobi BETWEEN m2.taxs_dt AND m2.taxe_dt " & vbCrLf &
    '                       "WHERE " & vbCrLf &
    '                       "    m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                       "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
    '                       "AND m1.sts = '0' " & vbCrLf &
    '                       "ORDER BY " & vbCrLf &
    '                       "    m1.sort "

    'Private strEX34S001 As String =
    '                       "SELECT " & vbCrLf &
    '                       "    shisetu_kbn, " & vbCrLf &
    '                       "    bunrui_cd, " & vbCrLf &
    '                       "    bunrui_nm, " & vbCrLf &
    '                       "    notax_flg, " & vbCrLf &
    '                       "    kamoku_cd, " & vbCrLf &
    '                       "    saimoku_cd, " & vbCrLf &
    '                       "    uchi_cd, " & vbCrLf &
    '                       "    shosai_cd, " & vbCrLf &
    '                       "    karikamoku_cd, " & vbCrLf &
    '                       "    kari_saimoku_cd, " & vbCrLf &
    '                       "    kari_uchi_cd, " & vbCrLf &
    '                       "    kari_shosai_cd, " & vbCrLf &
    '                       "    sts, " & vbCrLf &
    '                       "    sort " & vbCrLf &
    '                       "FROM " & vbCrLf &
    '                       "    fbunrui_mst " & vbCrLf &
    '                       "WHERE " & vbCrLf &
    '                       "    shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                       "AND :Riyobi BETWEEN kikan_from AND kikan_to " & vbCrLf &
    '                       "AND sts = '0' " & vbCrLf &
    '                       "ORDER BY " & vbCrLf &
    '                       "    sort "
    ' 2019/06/11 軽減税率対応 変更 Start E.Okuda@Compass

    ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

    ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
    Private strEX34S002 As String =
                            "SELECT " & vbCrLf &
                            "    m1.bunrui_cd as bunrui_cd, " & vbCrLf &
                            "    m1.futai_cd, " & vbCrLf &
                            "    m1.futai_nm, " & vbCrLf &
                            "    m1.tanka as tanka, " & vbCrLf &
                            "    m1.tani as tani, " & vbCrLf &
                            "    m1.sort as sort, " & vbCrLf &
                            "    m2.bunrui_nm as bunrui_nm, " & vbCrLf &
                            "    m2.shukei_grp, " & vbCrLf &
                            "    m2.kamoku_cd, " & vbCrLf &
                            "    m2.saimoku_cd, " & vbCrLf &
                            "    m2.uchi_cd, " & vbCrLf &
                            "    m2.shosai_cd, " & vbCrLf &
                            "    m2.karikamoku_cd, " & vbCrLf &
                            "    m2.kari_saimoku_cd, " & vbCrLf &
                            "    m2.kari_uchi_cd, " & vbCrLf &
                            "    m2.kari_shosai_cd, " & vbCrLf &
                            "    m2.notax_flg, " & vbCrLf &
                            "    m2.tax_kbn, " & vbCrLf &
                            "    m4.tax_rate, " & vbCrLf &
                            "    m3.kamoku_nm, " & vbCrLf &
                            "    m3.saimoku_nm, " & vbCrLf &
                            "    m3.uchi_nm, " & vbCrLf &
                            "    m3.shosai_nm " & vbCrLf &
                            "    ,0 " & vbCrLf &
                            "FROM " & vbCrLf &
                            "    futai_mst m1 " & vbCrLf &
                            "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
                            "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf &
                            "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
                            "    AND m2.sts = '0' " & vbCrLf &
                            "LEFT JOIN kamoku_mst m3 " & vbCrLf &
                            "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
                            "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
                            "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
                            "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
                            "    AND m3.sts = '0' " & vbCrLf &
                            "LEFT JOIN ( " & vbCrLf &
                            "    SELECT " & vbCrLf &
                            "        tax_kbn, " & vbCrLf &
                            "        CASE " & vbCrLf &
                            "            WHEN tax_kbn = 1 THEN tax_ritu " & vbCrLf &
                            "            WHEN tax_kbn = 2 THEN reduced_rate " & vbCrLf &
                            "	         WHEN tax_kbn = 3 THEN untaxed_rate " & vbCrLf &
                            "	         WHEN tax_kbn = 4 THEN tax_free " & vbCrLf &
                            "            WHEN tax_kbn = 5 THEN tax_exemption " & vbCrLf &
                            "	         WHEN tax_kbn = 6 THEN tax_old1 " & vbCrLf &
                            "            WHEN tax_kbn = 7 THEN tax_old2 " & vbCrLf &
                            "            WHEN tax_kbn = 8 THEN tax_spare1 " & vbCrLf &
                            "            WHEN tax_kbn = 9 THEN tax_spare2 " & vbCrLf &
                            "            WHEN tax_kbn = 10 THEN tax_spare3 " & vbCrLf &
                            "	         WHEN tax_kbn = 11 THEN tax_spare4 " & vbCrLf &
                            "            WHEN tax_kbn = 12 THEN tax_spare5 " & vbCrLf &
                            "        END as tax_rate " & vbCrLf &
                            "    FROM " & vbCrLf &
                            "        tax_mst " & vbCrLf &
                            "        CROSS JOIN generate_series(1,12) As s(tax_kbn) " & vbCrLf &
                            "    WHERE " & vbCrLf &
                            "        :Riyobi BETWEEN taxs_dt AND taxe_dt" & vbCrLf &
                            "    ) m4 " & vbCrLf &
                            "    ON m2.tax_kbn = m4.tax_kbn " & vbCrLf &
                            "WHERE " & vbCrLf &
                            "    m1.sts = '0' " & vbCrLf &
                            "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
                            "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "ORDER BY " & vbCrLf &
                            "    m1.sort "

    '' 2019/05/31 軽減税率対応 変更 Start E.Okuda@Compass
    'Private strEX34S002 As String =
    '                        "SELECT " & vbCrLf &
    '                        "    m1.bunrui_cd as bunrui_cd, " & vbCrLf &
    '                        "    m1.futai_cd, " & vbCrLf &
    '                        "    m1.futai_nm, " & vbCrLf &
    '                        "    m1.tanka as tanka, " & vbCrLf &
    '                        "    m1.tani as tani, " & vbCrLf &
    '                        "    m1.sort as sort, " & vbCrLf &
    '                        "    m2.bunrui_nm as bunrui_nm, " & vbCrLf &
    '                        "    m2.shukei_grp, " & vbCrLf &
    '                        "    m2.kamoku_cd, " & vbCrLf &
    '                        "    m2.saimoku_cd, " & vbCrLf &
    '                        "    m2.uchi_cd, " & vbCrLf &
    '                        "    m2.shosai_cd, " & vbCrLf &
    '                        "    m2.karikamoku_cd, " & vbCrLf &
    '                        "    m2.kari_saimoku_cd, " & vbCrLf &
    '                        "    m2.kari_uchi_cd, " & vbCrLf &
    '                        "    m2.kari_shosai_cd, " & vbCrLf &
    '                        "    m2.notax_flg, " & vbCrLf &
    '                        "    (CASE WHEN m2.zeiritsu IS NULL THEN m4.tax_ritu ELSE m2.zeiritsu END) as zeiritsu, " & vbCrLf &
    '                        "    m2.zeiritsu, " & vbCrLf &
    '                        "    m3.kamoku_nm, " & vbCrLf &
    '                        "    m3.saimoku_nm, " & vbCrLf &
    '                        "    m3.uchi_nm, " & vbCrLf &
    '                        "    m3.shosai_nm " & vbCrLf &
    '                        "    ,0 " & vbCrLf &
    '                        "FROM " & vbCrLf &
    '                        "    futai_mst m1 " & vbCrLf &
    '                        "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
    '                        "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf &
    '                        "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
    '                        "    AND m2.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN kamoku_mst m3 " & vbCrLf &
    '                        "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
    '                        "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
    '                        "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
    '                        "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
    '                        "    AND m3.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN tax_mst m4 " & vbCrLf &
    '                        "    ON :Riyobi BETWEEN m4.taxs_dt AND m4.taxe_dt " & vbCrLf &
    '                        "WHERE " & vbCrLf &
    '                        "    m1.sts = '0' " & vbCrLf &
    '                        "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
    '                        "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "ORDER BY " & vbCrLf &
    '                        "    m1.sort "

    ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

    'Private strEX34S002 As String =
    '                        "SELECT " & vbCrLf &
    '                        "    m1.bunrui_cd as bunrui_cd, " & vbCrLf &
    '                        "    m1.futai_cd, " & vbCrLf &
    '                        "    m1.futai_nm, " & vbCrLf &
    '                        "    m1.tanka as tanka, " & vbCrLf &
    '                        "    m1.tani as tani, " & vbCrLf &
    '                        "    m1.sort as sort, " & vbCrLf &
    '                        "    m2.bunrui_nm as bunrui_nm, " & vbCrLf &
    '                        "    m2.shukei_grp, " & vbCrLf &
    '                        "    m2.kamoku_cd, " & vbCrLf &
    '                        "    m2.saimoku_cd, " & vbCrLf &
    '                        "    m2.uchi_cd, " & vbCrLf &
    '                        "    m2.shosai_cd, " & vbCrLf &
    '                        "    m2.karikamoku_cd, " & vbCrLf &
    '                        "    m2.kari_saimoku_cd, " & vbCrLf &
    '                        "    m2.kari_uchi_cd, " & vbCrLf &
    '                        "    m2.kari_shosai_cd, " & vbCrLf &
    '                        "    m2.notax_flg, " & vbCrLf &
    '                        "    m3.kamoku_nm, " & vbCrLf &
    '                        "    m3.saimoku_nm, " & vbCrLf &
    '                        "    m3.uchi_nm, " & vbCrLf &
    '                        "    m3.shosai_nm " & vbCrLf &
    '                        "    ,0 " & vbCrLf &
    '                        "FROM " & vbCrLf &
    '                        "    futai_mst m1 " & vbCrLf &
    '                        "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
    '                        "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf &
    '                        "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
    '                        "    AND m2.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN kamoku_mst m3 " & vbCrLf &
    '                        "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
    '                        "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
    '                        "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
    '                        "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
    '                        "    AND m3.sts = '0' " & vbCrLf &
    '                        "WHERE " & vbCrLf &
    '                        "    m1.sts = '0' " & vbCrLf &
    '                        "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
    '                        "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "ORDER BY " & vbCrLf &
    '                        "    m1.sort "
    ' 2019/05/31 軽減税率対応 変更 End E.Okuda@Compass

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
