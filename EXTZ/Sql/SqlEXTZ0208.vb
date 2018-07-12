Imports Common
Imports Npgsql
Imports CommonEXT
Imports System.Text

Public Class SqlEXTZ0208
    Private commonLogicEXT As New CommonLogicEXT

    'SQL文(EXAS入金検索)
    ' 2016.02.12 UPD START↓ h.hagiwara
    'Private strEX41S002 As String = _
    '                        "SELECT " & vbCrLf & _
    '                        "    t1.aite_cd, " & vbCrLf & _
    '                        "    t2.yoyaku_start_dt, " & vbCrLf & _
    '                        "    t2.yoyaku_end_dt, " & vbCrLf & _
    '                        "    t1.aite_nm, " & vbCrLf & _
    '                        "    t1.yoyaku_no, " & vbCrLf & _
    '                        "    CASE t3.shisetu_kbn " & vbCrLf & _
    '                        "        WHEN '1' THEN t3.saiji_nm " & vbCrLf & _
    '                        "        WHEN '2' THEN t3.shutsuen_nm " & vbCrLf & _
    '                        "    End as saiji_nm, " & vbCrLf & _
    '                        "    t3.shisetu_kbn, " & vbCrLf & _
    '                        "    to_char(to_date(t1.nyukin_yotei_dt, 'yyyyMMdd'),'yyyy/MM/dd') as nyukin_yotei_dt, " & vbCrLf & _
    '                        "    to_char(to_date(t1.nyukin_dt, 'yyyyMMdd'),'yyyy/MM/dd') as nyukin_dt, " & vbCrLf & _
    '                        "    t1.seikyu_kin, " & vbCrLf & _
    '                        "    CASE t4.nyukin_input_flg " & vbCrLf & _
    '                        "        WHEN '0' THEN '' " & vbCrLf & _
    '                        "        WHEN '1' THEN '○' " & vbCrLf & _
    '                        "    END as nyukin_input_flg, " & vbCrLf & _
    '                        "    to_char(to_date(t1.seikyu_dt, 'yyyyMMdd'),'yyyy/MM/dd') as seikyu_dt, " & vbCrLf & _
    '                        "    to_char(to_date(t1.input_dt, 'yyyyMMdd'),'yyyy/MM/dd') as input_dt, " & vbCrLf & _
    '                        "    t1.sekikyu_no, " & vbCrLf & _
    '                        "    t1.seikyu_irai_no, " & vbCrLf & _
    '                        "    t1.monitor_no, " & vbCrLf & _
    '                        "    t1.nyukin_link_no " & vbCrLf & _
    '                        "FROM " & vbCrLf & _
    '                        "    exas_nyukin_tbl t1 " & vbCrLf & _
    '                        "    LEFT JOIN " & vbCrLf & _
    '                        "        ( " & vbCrLf & _
    '                        "            SELECT " & vbCrLf & _
    '                        "                MIN(yoyaku_dt) as yoyaku_start_dt, " & vbCrLf & _
    '                        "                MAX(yoyaku_dt) as yoyaku_end_dt, " & vbCrLf & _
    '                        "                yoyaku_no " & vbCrLf & _
    '                        "            FROM " & vbCrLf & _
    '                        "                ydt_tbl " & vbCrLf & _
    '                        "            GROUP BY yoyaku_no " & vbCrLf & _
    '                        "        ) t2 " & vbCrLf & _
    '                        "    ON  t1.yoyaku_no = t2.yoyaku_no " & vbCrLf & _
    '                        "    LEFT JOIN " & vbCrLf & _
    '                        "        yoyaku_tbl t3 " & vbCrLf & _
    '                        "    ON  t1.yoyaku_no = t3.yoyaku_no " & vbCrLf & _
    '                        "    LEFT JOIN " & vbCrLf & _
    '                        "        billpay_tbl t4 " & vbCrLf & _
    '                        "    ON  t1.seikyu_irai_no = t4.seikyu_irai_no " & vbCrLf & _
    '                        "WHERE 1 = 1 "
    Private strEX41S002 As String = _
                        "SELECT " & vbCrLf & _
                        "    t1.aite_cd, " & vbCrLf & _
                        "    t2.yoyaku_start_dt, " & vbCrLf & _
                        "    t2.yoyaku_end_dt, " & vbCrLf & _
                        "    t1.aite_nm, " & vbCrLf & _
                        "    t1.yoyaku_no, " & vbCrLf & _
                        "    CASE t3.shisetu_kbn " & vbCrLf & _
                        "        WHEN '1' THEN t3.saiji_nm " & vbCrLf & _
                        "        WHEN '2' THEN t3.shutsuen_nm " & vbCrLf & _
                        "    End as saiji_nm, " & vbCrLf & _
                        "    t3.shisetu_kbn, " & vbCrLf & _
                        "    to_char(to_date(t1.nyukin_yotei_dt, 'yyyyMMdd'),'yyyy/MM/dd') as nyukin_yotei_dt, " & vbCrLf & _
                        "    to_char(to_date(t1.nyukin_dt, 'yyyyMMdd'),'yyyy/MM/dd') as nyukin_dt, " & vbCrLf & _
                        "    t1.seikyu_kin, " & vbCrLf & _
                        "    CASE WHEN TRIM(t1.seikyu_irai_no) = '' " & vbCrLf & _
                        "           THEN  CASE t5.nyukin_input_flg " & vbCrLf & _
                        "                    WHEN '0' THEN '' " & vbCrLf & _
                        "                    WHEN '1' THEN '○' " & vbCrLf & _
                        "                 END" & vbCrLf & _
                        "           ELSE " & vbCrLf & _
                        "                 CASE t4.nyukin_input_flg " & vbCrLf & _
                        "                     WHEN '0' THEN '' " & vbCrLf & _
                        "                     WHEN '1' THEN '○' " & vbCrLf & _
                        "                 END" & vbCrLf & _
                        "    END as nyukin_input_flg, " & vbCrLf & _
                        "    to_char(to_date(t1.seikyu_dt, 'yyyyMMdd'),'yyyy/MM/dd') as seikyu_dt, " & vbCrLf & _
                        "    to_char(to_date(t1.input_dt, 'yyyyMMdd'),'yyyy/MM/dd') as input_dt, " & vbCrLf & _
                        "    t1.sekikyu_no, " & vbCrLf & _
                        "    t1.seikyu_irai_no, " & vbCrLf & _
                        "    t1.monitor_no, " & vbCrLf & _
                        "    t1.nyukin_link_no " & vbCrLf & _
                        "FROM " & vbCrLf & _
                        "    exas_nyukin_tbl t1 " & vbCrLf & _
                        "    LEFT JOIN " & vbCrLf & _
                        "        ( " & vbCrLf & _
                        "            SELECT " & vbCrLf & _
                        "                MIN(yoyaku_dt) as yoyaku_start_dt, " & vbCrLf & _
                        "                MAX(yoyaku_dt) as yoyaku_end_dt, " & vbCrLf & _
                        "                yoyaku_no " & vbCrLf & _
                        "            FROM " & vbCrLf & _
                        "                ydt_tbl " & vbCrLf & _
                        "            GROUP BY yoyaku_no " & vbCrLf & _
                        "        ) t2 " & vbCrLf & _
                        "    ON  t1.yoyaku_no = t2.yoyaku_no " & vbCrLf & _
                        "    LEFT JOIN " & vbCrLf & _
                        "        yoyaku_tbl t3 " & vbCrLf & _
                        "    ON  t1.yoyaku_no = t3.yoyaku_no " & vbCrLf & _
                        "    LEFT JOIN " & vbCrLf & _
                        "        billpay_tbl t4 " & vbCrLf & _
                        "    ON  t1.seikyu_irai_no = t4.seikyu_irai_no " & vbCrLf & _
                        "    LEFT JOIN " & vbCrLf & _
                        "        billpay_tbl t5 " & vbCrLf & _
                        "    ON  t1.NYUKIN_LINK_NO = t5.NYUKIN_LINK_NO " & vbCrLf & _
                        "WHERE 1 = 1 "
    ' 2016.02.12 UPD END↑ h.hagiwara

    ' 2016.09.12 ADD START m.hayabuchi
    '   Private strDup As String = _
    '              "SELECT 1" & vbCrLf & _
    '              "FROM billpay_tbl" & vbCrLf & _
    '              "WHERE nyukin_link_no = :nyukinlinkno"

    Private strDup As String = _
                              "select B.YOYAKU_NO," & vbCrLf & _
                              "Y.SAIJI_NM," & vbCrLf & _
                              "Y.RIYO_NM," & vbCrLf & _
                              "MIN(YD.YOYAKU_DT) || '～' || MAX(YD.YOYAKU_DT) YOYAKU_DAY," & vbCrLf & _
                              "B.SEIKYU_IRAI_NO," & vbCrLf & _
                              "B.SEIKYU_NAIYO," & vbCrLf & _
                              "B.SEIKYU_KIN" & vbCrLf & _
                              "from billpay_tbl B" & vbCrLf & _
                              "left join yoyaku_tbl Y" & vbCrLf & _
                              "on B.yoyaku_no = Y.yoyaku_no" & vbCrLf & _
                              "left join ydt_tbl YD" & vbCrLf & _
                              "on B.yoyaku_no = YD.yoyaku_no" & vbCrLf & _
                              "where B.nyukin_link_no = :nyukinlinkno" & vbCrLf & _
                              "group by B.YOYAKU_NO,Y.SAIJI_NM,Y.RIYO_NM,B.SEIKYU_IRAI_NO,B.SEIKYU_NAIYO,B.SEIKYU_KIN"

    ''' <summary>
    ''' EXAS入金情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0208"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function searchEXASNyukin(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0208 As DataEXTZ0208) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'SQL文(SELECT)
            Dim strSQL As New StringBuilder
            strSQL.Append(strEX41S002)
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrAiteCd) = False Then
                strSQL.AppendLine(String.Format("AND t1.aite_cd = '{0}' ", dataEXTZ0208.PropStrAiteCd))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrAiteNm) = False Then
                strSQL.AppendLine(String.Format("AND t1.aite_nm LIKE '{0}{1}{2}' ", "%", dataEXTZ0208.PropStrAiteNm, "%"))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrNyukinYoteiFrom) = False Then
                strSQL.AppendLine(String.Format("AND TO_DATE(t1.nyukin_yotei_dt, 'YYYYMMDD') >= TO_DATE('{0}', 'YYYY/MM/DD') ", dataEXTZ0208.PropStrNyukinYoteiFrom))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrNyukinYoteiTo) = False Then
                strSQL.AppendLine(String.Format("AND TO_DATE(t1.nyukin_yotei_dt, 'YYYYMMDD') <= TO_DATE('{0}', 'YYYY/MM/DD') ", dataEXTZ0208.PropStrNyukinYoteiTo))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrNyukinFrom) = False Then
                strSQL.AppendLine(String.Format("AND TO_DATE(t1.nyukin_dt, 'YYYYMMDD') >= TO_DATE('{0}', 'YYYY/MM/DD') ", dataEXTZ0208.PropStrNyukinFrom))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrNyukinTo) = False Then
                strSQL.AppendLine(String.Format("AND TO_DATE(t1.nyukin_dt, 'YYYYMMDD') <= TO_DATE('{0}', 'YYYY/MM/DD') ", dataEXTZ0208.PropStrNyukinTo))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrSeikyuNo) = False Then
                strSQL.AppendLine(String.Format("AND t1.sekikyu_no = '{0}' ", dataEXTZ0208.PropStrSeikyuNo))
            End If
            If String.IsNullOrEmpty(dataEXTZ0208.PropStrSeikyuIraiNo) = False Then
                strSQL.AppendLine(String.Format("AND t1.seikyu_irai_no = '{0}' ", dataEXTZ0208.PropStrSeikyuIraiNo))
            End If

            ' 2015.12.04 ADD START↓ h.hagiwara ソート順設定
            strSQL.AppendLine(" ORDER BY nyukin_yotei_dt DESC ,  YOYAKU_NO ")
            ' 2015.12.04 ADD END↑ h.hagiwara ソート順設定

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL.ToString, Cn)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "END", Nothing, Adapter.SelectCommand)
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
    ''' 請求入金情報表取得（重複チェック）
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0208"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2016.08.16 m.hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function searchDup(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0208 As DataEXTZ0208) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strDup

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("nyukinlinkno", NpgsqlTypes.NpgsqlDbType.Integer))        ' 入金リンクNO
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("nyukinlinkno").Value = dataEXTZ0208.PropIntResNyukinLink        ' 入金リンクNO
            End With

            '終了ログ出力
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
