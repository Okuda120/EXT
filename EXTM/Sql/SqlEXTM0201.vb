Imports Npgsql
Imports Common
Imports CommonEXT
Imports System.Text


''' <summary>利用者一覧画面Sqlクラス
''' </summary>
''' <remarks>利用者一覧画面のSql作成・設定を行う
''' <para>作成情報：2015/08/10 ozawa
''' <p>改訂情報</p>
''' </para></remarks>

Public Class SqlEXTM0201

    '利用者一覧取得用Sql(電話番号2削除後)
    Private strSelectRiyosha As String = "SELECT " & _
                                         "  FALSE AS 選択, " & _
                                         "  t1.RIYOSHA_CD AS 利用名コード," & _
                                         "  CASE t1.RIYO_LVL WHEN '1' THEN '一般'  WHEN '2' THEN '要注意' WHEN '3' THEN '利用不可' END AS 利用者レベル," & _
                                         "  t1.RIYO_NM AS 利用者名," & _
                                         "  t1.RIYO_KANA AS 利用者名カナ," & _
                                         "  t3.LAST_USE_DAY AS 最終利用日, " & _
                                         "  t2.AITE_CD AS 相手先コード, " & _
                                         "  t2.AITE_NM AS 相手先名, " & _
                                         "  t1.RIYO_TEL11 || t1.RIYO_TEL12 || t1.RIYO_TEL13 AS 電話番号," & _
                                         "  null AS 携帯," & _
                                         "  t1.RIYO_FAX11 || t1.RIYO_FAX12 || t1.RIYO_FAX13 AS FAX," & _
                                         "  t1.RIYO_YUBIN1 || RIYO_YUBIN2 AS 郵便番号," & _
                                         "  t1.RIYO_TODO || t1.RIYO_SHIKU || t1.RIYO_BAN || COALESCE(t1.RIYO_BUILD,'') AS 住所, " & _
                                         "  TRUE AS 詳細 " & _
                                         " ,t1.RIYO_LVL     AS RIYO_LVL " & _
                                         " ,t1.DAIHYO_NM    AS DAIHYO_NM " & _
                                         " ,t1.RIYO_TEL11   AS RIYO_TEL11 " & _
                                         " ,t1.RIYO_TEL12   AS RIYO_TEL12 " & _
                                         " ,t1.RIYO_TEL13   AS RIYO_TEL13 " & _
                                         " ,t1.RIYO_NAISEN  AS RIYO_NAISEN " & _
                                         " ,t1.RIYO_FAX11   AS RIYO_FAX11 " & _
                                         " ,t1.RIYO_FAX12   AS RIYO_FAX12 " & _
                                         " ,t1.RIYO_FAX13   AS RIYO_FAX13 " & _
                                         " ,t1.RIYO_YUBIN1  AS RIYO_YUBIN1 " & _
                                         " ,t1.RIYO_YUBIN2  AS RIYO_YUBIN2 " & _
                                         " ,t1.RIYO_TODO    AS RIYO_TODO " & _
                                         " ,t1.RIYO_SHIKU   AS RIYO_SHIKU " & _
                                         " ,t1.RIYO_BAN     AS RIYO_BAN " & _
                                         " ,t1.RIYO_BUILD   AS RIYO_BUILD " & _
                                         "FROM RIYOSHA_MST t1 " & _
                                         " LEFT JOIN AITESAKI_MST t2 " & _
                                         "        ON  t1.AITE_CD = t2.AITE_CD " & _
                                         " LEFT JOIN " & _
                                         "     (SELECT " & _
                                         "          MAX(YOYAKU_DT) LAST_USE_DAY," & _
                                         "          t3a.RIYOSHA_CD " & _
                                         "      FROM YOYAKU_TBL t3a " & _
                                         "       LEFT JOIN YDT_TBL t3b " & _
                                         "              ON t3a.YOYAKU_NO = t3b.YOYAKU_NO " & _
                                         "      GROUP BY " & _
                                         "         t3a.RIYOSHA_CD " & _
                                         "     ) t3 " & _
                                         "       ON  t1.RIYOSHA_CD = t3.RIYOSHA_CD "

    ''' <summary>
    ''' 初期表示用SQL作成
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0201"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetSelectInit(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim strSQL As String
        Dim strWhere As String

        Try
            'SQL文(SELECT)
            strSQL = strSelectRiyosha
            strWhere = ""
            'Where句作成
            If dataEXTM0201.PropParamValue = "0" Then
                strWhere = "WHERE t1.RIYO_LVL = '1' "

            ElseIf dataEXTM0201.PropParamValue = "1" Then
                strWhere = "WHERE t1.RIYO_LVL = '2' OR  t1.RIYO_LVL = '3' "
            End If
            If dataEXTM0201.PropTxtRiyo_kana.Text <> "" Then
                strWhere += " AND t1.RIYO_KANA LIKE('" & dataEXTM0201.PropTxtRiyo_kana.Text & "%')"
            End If

            '並べ替える
            strWhere &= "ORDER BY t1.RIYOSHA_CD "

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 利用者一覧取得SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0201">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/11 ozawa
    ''' </para></remarks>
    ''' 
    Public Function SetSelectRiyosha(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0201 As DataEXTM0201) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim param As Npgsql.NpgsqlParameter() = New Npgsql.NpgsqlParameter(10) {}

        Try
            'SQLに渡すデータの設定
            'ラジオボタン
            param(0) = New NpgsqlParameter("setTUJYO", NpgsqlTypes.NpgsqlDbType.Boolean)
            param(0).Value = dataEXTM0201.PropRdoTujyo.Checked
            param(1) = New NpgsqlParameter("setCHUI", NpgsqlTypes.NpgsqlDbType.Boolean)
            param(1).Value = dataEXTM0201.PropRdoChui.Checked
            param(2) = New NpgsqlParameter("setHUKA", NpgsqlTypes.NpgsqlDbType.Boolean)
            param(2).Value = dataEXTM0201.PropRdoHuka.Checked
            '利用者番号
            param(3) = New NpgsqlParameter("setRiyoCd", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(3).Value = dataEXTM0201.PropTxtRiyo_cd.Text
            '利用者名
            param(4) = New NpgsqlParameter("setRiyoNm", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(4).Value = dataEXTM0201.PropTxtRiyo_nm.Text
            '利用者名カナ
            param(5) = New NpgsqlParameter("setRiyoKn", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(5).Value = dataEXTM0201.PropTxtRiyo_kana.Text
            '相手先コード
            param(6) = New NpgsqlParameter("setAiteCd", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(6).Value = dataEXTM0201.PropTxtAite_cd.Text
            '相手先名
            param(7) = New NpgsqlParameter("setAiteNm", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(7).Value = dataEXTM0201.PropTxtAite_nm.Text
            '電話番号①
            param(8) = New NpgsqlParameter("setTel1", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(8).Value = dataEXTM0201.PropTxtTel1.Text
            '電話番号②
            param(9) = New NpgsqlParameter("setTel2", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(9).Value = dataEXTM0201.PropTxtTel2.Text
            '電話番号③
            param(10) = New NpgsqlParameter("setTel3", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(10).Value = dataEXTM0201.PropTxtTel3.Text

            'SQL文(SELECT)
            strSQL = strSelectRiyosha
            'Where句作成
            strWhere = ""

            'ラジオボタン
            If dataEXTM0201.PropRdoTujyo.Checked = True Then
                '「通常」がチェックされている場合
                strWhere &= "WHERE (CASE :setTUJYO " & _
                            "WHEN TRUE THEN t1.RIYO_LVL = '1' END ) "
            ElseIf dataEXTM0201.PropRdoChui.Checked = True Then
                '「要注意」がチェックされている場合
                strWhere &= "WHERE (CASE :setCHUI " & _
                            "WHEN TRUE THEN t1.RIYO_LVL = '2' END ) "
            ElseIf dataEXTM0201.PropRdoHuka.Checked = True Then
                '「使用不可」がチェックされている場合
                strWhere &= "WHERE (CASE :setHUKA " & _
                            "WHEN TRUE THEN t1.RIYO_LVL = '3' END ) "
            End If

            If dataEXTM0201.PropTxtRiyo_cd.Text <> System.String.Empty Then
                '利用者番号が入力されている場合
                strWhere &= "AND t1.RIYOSHA_CD = :setRiyoCd "
            End If

            If dataEXTM0201.PropTxtRiyo_nm.Text <> System.String.Empty Then
                '利用者名が入力されている場合
                strWhere &= "AND t1.RIYO_NM LIKE '%' || :setRiyoNm  || '%' "
            End If

            If dataEXTM0201.PropTxtRiyo_kana.Text <> System.String.Empty Then
                '利用者カナが入力されている場合
                strWhere &= "AND t1.RIYO_KANA LIKE '%' || :setRiyoKn  || '%' "
            End If

            If dataEXTM0201.PropTxtAite_cd.Text <> System.String.Empty Then
                '相手先コードが入力されている場合
                strWhere &= "AND t2.AITE_CD = :setAiteCd "
            End If

            If dataEXTM0201.PropTxtAite_nm.Text <> System.String.Empty Then
                '相手先名が入力されている場合
                strWhere &= "AND t2.AITE_NM LIKE '%' || :setAiteNm  || '%' "
            End If

            If dataEXTM0201.PropTxtTel1.Text <> System.String.Empty And
               dataEXTM0201.PropTxtTel2.Text <> System.String.Empty And
               dataEXTM0201.PropTxtTel3.Text <> System.String.Empty Then
                '電話番号が入力されている場合
                strWhere &= "AND t1.RIYO_TEL11 || t1.RIYO_TEL12 || t1.RIYO_TEL13 = :setTel1 || :setTel2 || :setTel3 "
            End If

            '並べ替え
            strWhere &= "ORDER BY t1.RIYOSHA_CD "

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値をセット
            Adapter.SelectCommand.Parameters.AddRange(param)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function
End Class
