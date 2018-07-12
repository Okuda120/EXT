Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' ALSOK現金入金機データ登録画面Sqlクラス
''' </summary>
''' <remarks>ALSOK現金入金機データ登録画面のSQLの作成・設定を行う
''' <para>作成情報：2015.08.14 h.hagiwara
''' <p>改訂情報:</p>
''' </para></remarks>

Public Class SqlEXTY0102

    'インスタンス作成
    Private commonLogicEXT As New CommonLogicEXT

    'SQL文宣言

    ' サーバ日時取得
    Private EX26S000 As String = "SELECT Now() AS Sysdate "

    '電子マネー入力用コンボ（レジ情報）の取得
    Private EX26S001 As String = "SELECT " & vbCrLf & _
                                 " REGI_CD " & vbCrLf & _
                                 ",REGI_NM " & vbCrLf & _
                                 "FROM  ALSOK_TENPO_MST " & vbCrLf & _
                                 "WHERE COALESCE(DEPOSIT_KBN,'0') = '2' " & vbCrLf & _
                                 "ORDER BY REGI_CD "

    '電子マネー入力用コンボ（店舗情報）の取得
    Private EX26S002 As String = "SELECT " & vbCrLf & _
                                 " TENPO_CD " & vbCrLf & _
                                 ",TENPO_NM " & vbCrLf & _
                                 "FROM  TENPO_MST " & vbCrLf & _
                                 "ORDER BY TENPO_CD "

    ' 入金情報の取得
    ' 2015.12.15 UPD START↓ h.hagiwara
    'Private EX26S003 As String = "SELECT " & vbCrLf & _
    '                             " A.DEPOSIT_MACHINE_CD AS DEPOSIT_MACHINE_CD ," & vbCrLf & _
    '                             " A.SEQ  AS SEQ," & vbCrLf & _
    '                             " A.DEPOSIT_KBN AS DEPOSIT_KBN , " & vbCrLf & _
    '                             " A.DEPOSIT_DT AS DEPOSIT_DT , " & vbCrLf & _
    '                             " A.YOYAKU_NO AS YOYAKU_NO , " & vbCrLf & _
    '                             " A.REGISTER_CD AS REGISTER_CD ," & vbCrLf & _
    '                             " B.REGI_NM AS REGI_NM ," & vbCrLf & _
    '                             " A.TENPO_CD AS TENPO_CD , " & vbCrLf & _
    '                             " C.TENPO_NM AS TENPO_NM ," & vbCrLf & _
    '                             " A.DEPOSIT_AMOUNT AS DEPOSIT_AMOUNT , " & vbCrLf & _
    '                             " A.KEIYAKUSAKI_CD AS KEIYAKUSAKI_CD  " & vbCrLf & _
    '                             "FROM ALSOK_DEPOSIT_TBL A" & vbCrLf & _
    '                             " LEFT OUTER JOIN ALSOK_TENPO_MST B " & vbCrLf & _
    '                             "       ON A.REGISTER_CD = B.REGI_CD " & vbCrLf & _
    '                             " LEFT OUTER JOIN TENPO_MST C " & vbCrLf & _
    '                             "       ON A.TENPO_CD = C.TENPO_CD "
    Private EX26S003 As String = "SELECT " & vbCrLf & _
                                 " A.DEPOSIT_MACHINE_CD AS DEPOSIT_MACHINE_CD ," & vbCrLf & _
                                 " A.SEQ  AS SEQ," & vbCrLf & _
                                 " A.DEPOSIT_KBN AS DEPOSIT_KBN , " & vbCrLf & _
                                 " A.DEPOSIT_DT AS DEPOSIT_DT , " & vbCrLf & _
                                 " A.YOYAKU_NO AS YOYAKU_NO , " & vbCrLf & _
                                 " A.REGISTER_CD AS REGISTER_CD ," & vbCrLf & _
                                 " B.REGI_NM AS REGI_NM ," & vbCrLf & _
                                 " A.TENPO_CD AS TENPO_CD , " & vbCrLf & _
                                 " C.TENPO_NM AS TENPO_NM ," & vbCrLf & _
                                 " A.DEPOSIT_AMOUNT AS DEPOSIT_AMOUNT , " & vbCrLf & _
                                 " A.KEIYAKUSAKI_CD AS KEIYAKUSAKI_CD,  " & vbCrLf & _
                                 " B.SHISETSU_KBN AS SHISETU_KBN " & vbCrLf & _
                                 "FROM ALSOK_DEPOSIT_TBL A" & vbCrLf & _
                                 " LEFT OUTER JOIN ALSOK_TENPO_MST B " & vbCrLf & _
                                 "       ON A.REGISTER_CD = B.REGI_CD " & vbCrLf & _
                                 " LEFT OUTER JOIN TENPO_MST C " & vbCrLf & _
                                 "       ON A.TENPO_CD = C.TENPO_CD "
    ' 2015.12.15 UPD END↑ h.hagiwara

    ' 予約情報の取得
    Private EX26S004 As String = "SELECT " & vbCrLf & _
                                 " A.YOYAKU_NO AS YOYAKU_NO , " & vbCrLf & _
                                 " CASE WHEN A.SHISETU_KBN = '1' THEN B.SAIJI_NM " & vbCrLf & _
                                 "      WHEN A.SHISETU_KBN = '2' THEN B.SHUTSUEN_NM " & vbCrLf & _
                                 " END " & vbCrLf & _
                                 "FROM YDT_TBL A " & vbCrLf & _
                                 " INNER JOIN YOYAKU_TBL B " & vbCrLf & _
                                 "    ON  A.YOYAKU_NO = B.YOYAKU_NO " & vbCrLf & _
                                 "WHERE A.SHISETU_KBN = (SELECT SHISETSU_KBN FROM ALSOK_TENPO_MST WHERE REGI_CD = :REGI_CD ) " & vbCrLf & _
                                 " AND  A.YOYAKU_DT   = :YOYSKU_DT " & vbCrLf
    ' 2015.12.04 ADD START↓ h.hagiwara
    Private EX26S004A As String = "UNION " & vbCrLf & _
                                  "SELECT " & vbCrLf & _
                                  " C.YOYAKU_NO AS YOYAKU_NO , " & vbCrLf & _
                                  " CASE WHEN C.SHISETU_KBN = '1' THEN C.SAIJI_NM " & vbCrLf & _
                                  "      WHEN C.SHISETU_KBN = '2' THEN C.SHUTSUEN_NM " & vbCrLf & _
                                  " END " & vbCrLf & _
                                  "FROM YOYAKU_TBL C " & vbCrLf & _
                                  "WHERE C.YOYAKU_NO = :YOYAKU_NO " & vbCrLf
    ' 2015.12.04 ADD END↑ h.hagiwara
    ' 2015.12.18 ADD START↓ h.hagiwara
    Private EX26S004B As String = "SELECT " & vbCrLf & _
                                 " A.YOYAKU_NO AS YOYAKU_NO , " & vbCrLf & _
                                 " CASE WHEN A.SHISETU_KBN = '1' THEN B.SAIJI_NM " & vbCrLf & _
                                 "      WHEN A.SHISETU_KBN = '2' THEN B.SHUTSUEN_NM " & vbCrLf & _
                                 " END " & vbCrLf & _
                                 "FROM YDT_TBL A " & vbCrLf & _
                                 " INNER JOIN YOYAKU_TBL B " & vbCrLf & _
                                 "    ON  A.YOYAKU_NO = B.YOYAKU_NO " & vbCrLf & _
                                 "WHERE A.SHISETU_KBN = :SHISETU_KBN " & vbCrLf & _
                                 " AND  A.YOYAKU_DT   = :YOYSKU_DT " & vbCrLf

    ' 2015.12.18 ADD END↑ h.hagiwara

    ' 入金機コード・連番の存在チェック用
    Private EX26S005 As String = "SELECT " & vbCrLf & _
                                 " COUNT(*) AS KENSU " & vbCrLf & _
                                 "FROM ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 "WHERE DEPOSIT_MACHINE_CD = :DEPOSIT_MACHINE_CD " & vbCrLf & _
                                 " AND  SEQ = :SEQ "

    ' レジ名取得用
    Private EX26S006 As String = "SELECT " & vbCrLf & _
                                 " REGI_NM  " & vbCrLf & _
                                 "FROM ALSOK_TENPO_MST " & vbCrLf & _
                                 "WHERE REGI_CD = :REGI_CD "

    ' 店舗名取得用
    Private EX26S007 As String = "SELECT " & vbCrLf & _
                                 " TENPO_NM  " & vbCrLf & _
                                 "FROM TENPO_MST " & vbCrLf & _
                                 "WHERE TENPO_CD = :TENPO_CD "

    ' ALSOK入金情報の登録（ＣＳＶ取り込み用）
    Private EX26I001 As String = "INSERT INTO ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 " ( " & vbCrLf & _
                                 "   DEPOSIT_MACHINE_CD ," & vbCrLf & _
                                 "   SEQ ," & vbCrLf & _
                                 "   DEPOSIT_KBN , " & vbCrLf & _
                                 "   DEPOSIT_DT , " & vbCrLf & _
                                 "   YOYAKU_NO , " & vbCrLf & _
                                 "   REGISTER_CD , " & vbCrLf & _
                                 "   TENPO_CD , " & vbCrLf & _
                                 "   DEPOSIT_AMOUNT , " & vbCrLf & _
                                 "   KEIYAKUSAKI_CD , " & vbCrLf & _
                                 "   ADD_DT , " & vbCrLf & _
                                 "   ADD_USER_CD " & vbCrLf & _
                                 " ) VALUES ( " & vbCrLf & _
                                 "   :DEPOSIT_MACHINE_CD ," & vbCrLf & _
                                 "   :SEQ ," & vbCrLf & _
                                 "   :DEPOSIT_KBN , " & vbCrLf & _
                                 "   :DEPOSIT_DT , " & vbCrLf & _
                                 "   :YOYAKU_NO , " & vbCrLf & _
                                 "   :REGISTER_CD , " & vbCrLf & _
                                 "   :TENPO_CD , " & vbCrLf & _
                                 "   :DEPOSIT_AMOUNT , " & vbCrLf & _
                                 "   :KEIYAKUSAKI_CD , " & vbCrLf & _
                                 "   :ADD_DT , " & vbCrLf & _
                                 "   :USER_CD " & vbCrLf & _
                                 " ) "

    ' ALSOK入金情報の登録（電子マネー用）
    Private EX26I002 As String = "INSERT INTO ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 " ( " & vbCrLf & _
                                 "   DEPOSIT_MACHINE_CD ," & vbCrLf & _
                                 "   SEQ ," & vbCrLf & _
                                 "   DEPOSIT_KBN , " & vbCrLf & _
                                 "   DEPOSIT_DT , " & vbCrLf & _
                                 "   YOYAKU_NO , " & vbCrLf & _
                                 "   REGISTER_CD , " & vbCrLf & _
                                 "   TENPO_CD , " & vbCrLf & _
                                 "   DEPOSIT_AMOUNT , " & vbCrLf & _
                                 "   KEIYAKUSAKI_CD , " & vbCrLf & _
                                 "   ADD_DT , " & vbCrLf & _
                                 "   ADD_USER_CD " & vbCrLf & _
                                 " ) VALUES ( " & vbCrLf & _
                                 "   :DEPOSIT_MACHINE_CD ," & vbCrLf & _
                                 "   ( SELECT COALESCE(MAX(SEQ),0) + 1 FROM ALSOK_DEPOSIT_TBL WHERE DEPOSIT_MACHINE_CD = :DEPOSIT_MACHINE_CD1 ) , " & vbCrLf & _
                                 "   :DEPOSIT_KBN , " & vbCrLf & _
                                 "   :DEPOSIT_DT , " & vbCrLf & _
                                 "   :YOYAKU_NO , " & vbCrLf & _
                                 "   :REGISTER_CD , " & vbCrLf & _
                                 "   :TENPO_CD , " & vbCrLf & _
                                 "   :DEPOSIT_AMOUNT , " & vbCrLf & _
                                 "   :KEIYAKUSAKI_CD , " & vbCrLf & _
                                 "   :ADD_DT , " & vbCrLf & _
                                 "   :USER_CD " & vbCrLf & _
                                 " ) "

    ' ALSOK入金情報の更新（ＣＳＶ取り込み用）
    Private EX26U001 As String = "UPDATE ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 " SET  YOYAKU_NO   = :YOYAKU_NO  " & vbCrLf & _
                                 "    , ADD_DT      = :ADD_DT  " & vbCrLf & _
                                 "    , ADD_USER_CD = :USER_ID " & vbCrLf & _
                                 "WHERE DEPOSIT_MACHINE_CD = :DEPOSIT_MACHINE_CD " & vbCrLf & _
                                 " AND  SEQ = :SEQ "

    ' ALSOK入金情報の更新（電子マネー用）
    Private EX26U002 As String = "UPDATE ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 " SET  " & vbCrLf & _
                                 "     DEPOSIT_DT     = :DEPOSIT_DT  " & vbCrLf & _
                                 "   , YOYAKU_NO      = :YOYAKU_NO  " & vbCrLf & _
                                 "   , REGISTER_CD    = :REGISTER_CD  " & vbCrLf & _
                                 "   , TENPO_CD       = :TENPO_CD  " & vbCrLf & _
                                 "   , DEPOSIT_AMOUNT = :DEPOSIT_AMOUNT  " & vbCrLf & _
                                 "   , KEIYAKUSAKI_CD = :KEIYAKUSAKI_CD  " & vbCrLf & _
                                 "   , ADD_DT         = :ADD_DT  " & vbCrLf & _
                                 "   , ADD_USER_CD    = :USER_ID " & vbCrLf & _
                                 "WHERE DEPOSIT_MACHINE_CD = :DEPOSIT_MACHINE_CD " & vbCrLf & _
                                 " AND  SEQ = :SEQ "

    ' ALSOK入金情報の削除
    Private EX26D001 As String = "DELETE FROM ALSOK_DEPOSIT_TBL " & vbCrLf & _
                                 "WHERE DEPOSIT_MACHINE_CD = :DEPOSIT_MACHINE_CD " & vbCrLf & _
                                 " AND  SEQ = :SEQ "

    ' レジ施設区分用
    Private EX26S010 As String = "SELECT " & vbCrLf & _
                                 " SHISETSU_KBN  " & vbCrLf & _
                                 "FROM ALSOK_TENPO_MST " & vbCrLf & _
                                 "WHERE REGI_CD = :REGI_CD "

    ''' <summary>
    ''' マスタデータ取得：ALSOKレジマスタ
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ALSOKレジマスタコンボボックス取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetCmbRejiMstData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S001

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' マスタデータ取得：店舗情報
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>店舗情報コンボボックス取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetCmbTenpMstData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S002

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 入金情報取得：ALSOK入金情報（ＣＳＶ取り込み分）
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶ取り込みで登録した入金情報取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetAlsokDepositInf(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""
        Dim strWhere As String = ""
        Dim strOrderBy As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S003

            ' WHERE条件の設定
            If dataEXTY0102.PropDtpDspFrom.txtDate.Text = Nothing Then
                strWhere = " WHERE DEPOSIT_DT <= :DSP_TO "
            ElseIf dataEXTY0102.PropDtpDspTo.txtDate.Text = Nothing Then
                strWhere = " WHERE DEPOSIT_DT >= :DSP_FROM "
            Else
                strWhere = " WHERE DEPOSIT_DT >= :DSP_FROM " & vbCrLf & _
                           "  AND  DEPOSIT_DT <= :DSP_TO "
            End If
            strWhere += " AND A.DEPOSIT_KBN = '1' "
            strOrderBy = " ORDER BY A.DEPOSIT_DT "
            strSQL += strWhere + strOrderBy

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("DSP_FROM", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 現金機利用日（自）
                .Add(New NpgsqlParameter("DSP_TO", NpgsqlTypes.NpgsqlDbType.Varchar))          ' 現金機利用日（至）
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("DSP_FROM").Value = dataEXTY0102.PropDtpDspFrom.txtDate.Text       ' 現金機利用日（自）
                .Parameters("DSP_TO").Value = dataEXTY0102.PropDtpDspTo.txtDate.Text           ' 現金機利用日（至）
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 入金情報取得：ALSOK入金情報（電子マネー用）
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>電子マネー分で登録した入金情報取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetDegitalCashInf(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""
        Dim strWhere As String = ""
        Dim strOrderBy As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S003

            ' WHERE条件の設定
            If dataEXTY0102.PropDtpDspFrom.txtDate.Text = Nothing Then
                strWhere = " WHERE DEPOSIT_DT <= :DSP_TO "
            ElseIf dataEXTY0102.PropDtpDspTo.txtDate.Text = Nothing Then
                strWhere = " WHERE DEPOSIT_DT >= :DSP_FROM "
            Else
                strWhere = " WHERE DEPOSIT_DT >= :DSP_FROM " & vbCrLf & _
                           "  AND  DEPOSIT_DT <= :DSP_TO "
            End If
            strWhere += " AND A.DEPOSIT_KBN = '2' "
            strOrderBy = " ORDER BY A.DEPOSIT_DT "
            strSQL += strWhere + strOrderBy

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("DSP_FROM", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 現金機利用日（自）
                .Add(New NpgsqlParameter("DSP_TO", NpgsqlTypes.NpgsqlDbType.Varchar))          ' 現金機利用日（至）
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("DSP_FROM").Value = dataEXTY0102.PropDtpDspFrom.txtDate.Text       ' 現金機利用日（自）
                .Parameters("DSP_TO").Value = dataEXTY0102.PropDtpDspTo.txtDate.Text           ' 現金機利用日（至）
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 予約情報取得：コンボ情報の取得
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>コンボに設定する予約情報取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqlYoyakudata(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S004

            ' 2015.12.04 ADD START↓ h.hagiwara
            If dataEXTY0102.PropStrUnionFlg = "1" Then
                strSQL = strSQL & EX26S004A
            End If

            strSQL = strSQL & "ORDER BY YOYAKU_NO "
            ' 2015.12.04 ADD END↑ h.hagiwara

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("REGI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))            ' レジ№
                .Add(New NpgsqlParameter("YOYSKU_DT", NpgsqlTypes.NpgsqlDbType.Varchar))          ' 利用日
                ' 2015.12.04 ADD START↓ h.hagiwara
                If dataEXTY0102.PropStrUnionFlg = "1" Then
                    .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))      ' 予約番号
                End If
                ' 2015.12.04 ADD END↑ h.hagiwara
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("REGI_CD").Value = dataEXTY0102.PropStrRegisterCd                     ' レジ№
                .Parameters("YOYSKU_DT").Value = dataEXTY0102.PropStrDepositDt                    ' 利用日
                ' 2015.12.04 ADD START↓ h.hagiwara
                If dataEXTY0102.PropStrUnionFlg = "1" Then
                    .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropStrYoyakuNo                 ' 予約番号
                End If
                ' 2015.12.04 ADD END↑ h.hagiwara
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表取得：重複チェック用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報が存在するか件数取得を行うSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetDupDataCnt(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S005

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' 連番
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropStrDepositCd                  ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropStrSeq)                        ' 連番
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' レジ名取得：ＣＳＶデータ取り込み時用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報が存在するか件数取得を行うSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSqlRegisterNm(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S006

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("REGI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' レジ№
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("REGI_CD").Value = dataEXTY0102.PropStrRegisterCd                 ' レジ№
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 店舗名取得：ＣＳＶデータ取り込み時用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報の店舗名を取得するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSqlTenpoNm(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S007

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("TENPO_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 店舗№
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("TENPO_CD").Value = dataEXTY0102.PropStrTenpoCd                    '  店舗№
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表登録：ＣＳＶ取り込み用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶ取り込みした情報をALSOK入金情報表に登録するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetInsertAlsokDepositInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26I001

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' SEQ
                .Add(New NpgsqlParameter("DEPOSIT_KBN", NpgsqlTypes.NpgsqlDbType.Char))                  ' 入金区分
                .Add(New NpgsqlParameter("DEPOSIT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))                ' 入金日
                .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 予約NO
                .Add(New NpgsqlParameter("REGISTER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))               ' レジコード
                .Add(New NpgsqlParameter("TENPO_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                  ' 店舗コード
                .Add(New NpgsqlParameter("DEPOSIT_AMOUNT", NpgsqlTypes.NpgsqlDbType.Integer))            ' 入金額
                .Add(New NpgsqlParameter("KEIYAKUSAKI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))            ' 契約先コード
                .Add(New NpgsqlParameter("ADD_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                  ' 登録年月日
                .Add(New NpgsqlParameter("USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                   ' 登録ユーザCD
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("Seq"))                          ' SEQ
                .Parameters("DEPOSIT_KBN").Value = dataEXTY0102.PropRowReg.Item("DepositKbn")                          ' 入金区分
                .Parameters("DEPOSIT_DT").Value = dataEXTY0102.PropRowReg.Item("DepositDt")                            ' 入金日
                .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropRowReg.Item("YoyakuNo")                              ' 予約NO
                .Parameters("REGISTER_CD").Value = dataEXTY0102.PropRowReg.Item("RegisterCd")                          ' レジコード
                .Parameters("TENPO_CD").Value = dataEXTY0102.PropRowReg.Item("TenpoCd")                                ' 店舗コード
                .Parameters("DEPOSIT_AMOUNT").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("DepositAmount"))     ' 入金額
                .Parameters("KEIYAKUSAKI_CD").Value = dataEXTY0102.PropRowReg.Item("KeiyakusakiCd")                    ' 契約先コード
                .Parameters("ADD_DT").Value = dataEXTY0102.PropDtSysDate                                               ' 登録年月日
                .Parameters("USER_CD").Value = dataEXTY0102.PropStrUserId                                              ' 登録ユーザCD
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表変更：ＣＳＶ取り込み用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶ取り込みした情報を変更するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetUpdateAlsokDepositInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26U001

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 予約NO
                .Add(New NpgsqlParameter("ADD_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                  ' 登録年月日
                .Add(New NpgsqlParameter("USER_ID", NpgsqlTypes.NpgsqlDbType.Varchar))                   ' 登録ユーザCD
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' SEQ
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropRowReg.Item("YoyakuNo")                              ' 予約NO
                .Parameters("ADD_DT").Value = dataEXTY0102.PropDtSysDate                                               ' 登録年月日
                .Parameters("USER_ID").Value = dataEXTY0102.PropStrUserId                                              ' 登録ユーザCD
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("Seq"))                          ' SEQ
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表削除：ＣＳＶ取り込み用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ＣＳＶ取り込みした情報を削除するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetDeleteAlsokDepositInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26D001

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' SEQ
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("Seq"))                          ' SEQ
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表登録：電子マネー用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>電子マネー情報をALSOK入金情報表に登録するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetInsertDegitalCashInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26I002

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD1", NpgsqlTypes.NpgsqlDbType.Varchar))       ' 入金機コード
                .Add(New NpgsqlParameter("DEPOSIT_KBN", NpgsqlTypes.NpgsqlDbType.Char))                  ' 入金区分
                .Add(New NpgsqlParameter("DEPOSIT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))                ' 入金日
                .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 予約NO
                .Add(New NpgsqlParameter("REGISTER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))               ' レジコード
                .Add(New NpgsqlParameter("TENPO_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                  ' 店舗コード
                .Add(New NpgsqlParameter("DEPOSIT_AMOUNT", NpgsqlTypes.NpgsqlDbType.Integer))            ' 入金額
                .Add(New NpgsqlParameter("KEIYAKUSAKI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))            ' 契約先コード
                .Add(New NpgsqlParameter("ADD_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                  ' 登録年月日
                .Add(New NpgsqlParameter("USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                   ' 登録ユーザCD
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("DEPOSIT_MACHINE_CD1").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                   ' 入金機コード
                .Parameters("DEPOSIT_KBN").Value = dataEXTY0102.PropRowReg.Item("DepositKbn")                          ' 入金区分
                .Parameters("DEPOSIT_DT").Value = dataEXTY0102.PropRowReg.Item("DepositDt")                            ' 入金日
                .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropRowReg.Item("YoyakuNo")                              ' 予約NO
                .Parameters("REGISTER_CD").Value = dataEXTY0102.PropRowReg.Item("RegisterCd")                          ' レジコード
                .Parameters("TENPO_CD").Value = dataEXTY0102.PropRowReg.Item("TenpoCd")                                ' 店舗コード
                .Parameters("DEPOSIT_AMOUNT").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("DepositAmount"))     ' 入金額
                .Parameters("KEIYAKUSAKI_CD").Value = dataEXTY0102.PropRowReg.Item("KeiyakusakiCd")                    ' 契約先コード
                .Parameters("ADD_DT").Value = dataEXTY0102.PropDtSysDate                                               ' 登録年月日
                .Parameters("USER_CD").Value = dataEXTY0102.PropStrUserId                                              ' 登録ユーザCD
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情情報変更：電子マネー用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>電子マネー情報を変更するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetUpdateDegitalCashInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26U002

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' SEQ
                .Add(New NpgsqlParameter("DEPOSIT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))                ' 入金日
                .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 予約NO
                .Add(New NpgsqlParameter("REGISTER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))               ' レジコード
                .Add(New NpgsqlParameter("TENPO_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                  ' 店舗コード
                .Add(New NpgsqlParameter("DEPOSIT_AMOUNT", NpgsqlTypes.NpgsqlDbType.Integer))            ' 入金額
                .Add(New NpgsqlParameter("KEIYAKUSAKI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))            ' 契約先コード
                .Add(New NpgsqlParameter("ADD_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                  ' 登録年月日
                .Add(New NpgsqlParameter("USER_ID", NpgsqlTypes.NpgsqlDbType.Varchar))                   ' 登録ユーザCD
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("Seq"))                          ' SEQ
                .Parameters("DEPOSIT_DT").Value = dataEXTY0102.PropRowReg.Item("DepositDt")                            ' 入金日
                .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropRowReg.Item("YoyakuNo")                              ' 予約NO
                .Parameters("REGISTER_CD").Value = dataEXTY0102.PropRowReg.Item("RegisterCd")                          ' レジコード
                .Parameters("TENPO_CD").Value = dataEXTY0102.PropRowReg.Item("TenpoCd")                                ' 店舗コード
                .Parameters("DEPOSIT_AMOUNT").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("DepositAmount"))     ' 入金額
                .Parameters("KEIYAKUSAKI_CD").Value = dataEXTY0102.PropRowReg.Item("KeiyakusakiCd")                    ' 契約先コード
                .Parameters("ADD_DT").Value = dataEXTY0102.PropDtSysDate                                               ' 登録年月日
                .Parameters("USER_ID").Value = dataEXTY0102.PropStrUserId                                              ' 登録ユーザCD
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表削除：電子マネー用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>電子マネー入力情報を削除するSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetDeleteDegitalCashInf(ByRef Adapter As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26D001

            'データアダプタに、SQLのSELECT文を設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.Parameters
                .Add(New NpgsqlParameter("DEPOSIT_MACHINE_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 入金機コード
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                       ' SEQ
            End With

            'バインド変数に値をセット
            With Adapter
                .Parameters("DEPOSIT_MACHINE_CD").Value = dataEXTY0102.PropRowReg.Item("DepositCd")                    ' 入金機コード
                .Parameters("SEQ").Value = Integer.Parse(dataEXTY0102.PropRowReg.Item("Seq"))                          ' SEQ
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' サーバ日時取得
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>サーバ日時取得のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSysDate(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S000

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' レジ施設区分取得：電子マネーレジ変更時取得用
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>レジに設定されている施設区分を取得
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqRejiShisetuKbn(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S010

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("REGI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))        ' レジ№
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("REGI_CD").Value = dataEXTY0102.PropStrRegisterCd                 ' レジ№
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 予約情報取得：コンボ情報の取得（施設区分）
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN]ALSOK現金入金機データ登録画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>コンボに設定する予約情報取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.12.18 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqlYoyakudata2(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = EX26S004B

            ' 2015.12.04 ADD START↓ h.hagiwara
            If dataEXTY0102.PropStrUnionFlg = "1" Then
                strSQL = strSQL & EX26S004A
            End If

            strSQL = strSQL & "ORDER BY YOYAKU_NO "
            ' 2015.12.04 ADD END↑ h.hagiwara

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("SHISETU_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))        ' 施設区分
                .Add(New NpgsqlParameter("YOYSKU_DT", NpgsqlTypes.NpgsqlDbType.Varchar))          ' 利用日
                ' 2015.12.04 ADD START↓ h.hagiwara
                If dataEXTY0102.PropStrUnionFlg = "1" Then
                    .Add(New NpgsqlParameter("YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))      ' 予約番号
                End If
                ' 2015.12.04 ADD END↑ h.hagiwara
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("SHISETU_KBN").Value = dataEXTY0102.PropStrShisetuKbn                 ' 施設区分№
                .Parameters("YOYSKU_DT").Value = dataEXTY0102.PropStrDepositDt                    ' 利用日
                ' 2015.12.04 ADD START↓ h.hagiwara
                If dataEXTY0102.PropStrUnionFlg = "1" Then
                    .Parameters("YOYAKU_NO").Value = dataEXTY0102.PropStrYoyakuNo                 ' 予約番号
                End If
                ' 2015.12.04 ADD END↑ h.hagiwara
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

End Class
