Imports Npgsql
Imports Common

''' <summary>
''' EXシアター消費税マスタメンテSqlクラス
''' </summary>
''' <remarks>EXシアター消費税マスタメンテのSQLの作成・設定を行う
''' <para>作成情報：2015/08/26 hayabuchi
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTM0103

    ' --- 2020/03/05 軽減税率対応 Start E.Okuda@Compass ---
    Private Const COL_NAME_SEQ As String = "SEQ"
    Private Const COL_NAME_UPDT_KBN As String = "更新区分"
    Private Const COL_NAME_LATEST_TAXS_DT As String = "修正前開始日"
    Private Const COL_NAME_LATEST_TAXE_DT As String = "修正前終了日"
    Private Const COL_NAME_TAXS_DT As String = "開始日"
    Private Const COL_NAME_TAXE_DT As String = "終了日"

    Private Const COL_SHEET_SEQ As Integer = 0                      ' SEQ
    Private Const COL_SHEET_UPDATE_KBN As Integer = 1               ' 更新区分
    Private Const COL_SHEET_BEFORE_TAXS_DT As Integer = 2           ' 修正前開始日
    Private Const COL_SHEET_BEFORE_TAXE_DT As Integer = 3           ' 修正前終了日
    Private Const COL_SHEET_TAXS_DT As Integer = 4                  ' 開始日
    Private Const COL_SHEET_TAXE_DT As Integer = 5                  ' 終了日
    Private Const COL_SHEET_TAX_RITU As Integer = 6                 ' 課売税率
    Private Const COL_SHEET_BEFORE_TAX_RITU As Integer = 7          ' 修正前課売税率
    Private Const COL_SHEET_REDUCED_RATE As Integer = 8             ' 課売軽減税率
    Private Const COL_SHEET_BEFORE_REDUCED_RATE As Integer = 9      ' 修正前課売軽減税率
    Private Const COL_SHEET_UNTAXED_RATE As Integer = 10            ' 対象外税率
    Private Const COL_SHEET_BEFORE_UNTAXED_RATE As Integer = 11     ' 修正前対象外税率
    Private Const COL_SHEET_TAX_FREE As Integer = 12                ' 非課税税率
    Private Const COL_SHEET_BEFORE_TAX_FREE As Integer = 13         ' 修正前非課税税率
    Private Const COL_SHEET_TAX_EXEMPTION As Integer = 14           ' 免税税率
    Private Const COL_SHEET_BEFORE_TAX_EXEMPTION As Integer = 15    ' 修正前免税税率
    Private Const COL_SHEET_TAX_OLD1 As Integer = 16                ' 旧課税税率1
    Private Const COL_SHEET_BEFORE_TAX_OLD1 As Integer = 17         ' 修正前旧課税税率1
    Private Const COL_SHEET_TAX_OLD2 As Integer = 18                ' 旧課税税率2
    Private Const COL_SHEET_BEFORE_TAX_OLD2 As Integer = 19         ' 修正前旧課税税率2
    Private Const COL_SHEET_TAX_SPARE1 As Integer = 20              ' 予備税率1
    Private Const COL_SHEET_BEFORE_TAX_SPARE1 As Integer = 21       ' 修正前予備税率1
    Private Const COL_SHEET_TAX_SPARE2 As Integer = 22              ' 予備税率2
    Private Const COL_SHEET_BEFORE_TAX_SPARE2 As Integer = 23       ' 修正前予備税率2
    Private Const COL_SHEET_TAX_SPARE3 As Integer = 24              ' 予備税率3
    Private Const COL_SHEET_BEFORE_TAX_SPARE3 As Integer = 25       ' 修正前予備税率3
    Private Const COL_SHEET_TAX_SPARE4 As Integer = 26              ' 予備税率4
    Private Const COL_SHEET_BEFORE_TAX_SPARE4 As Integer = 27       ' 修正前予備税率4
    Private Const COL_SHEET_TAX_SPARE5 As Integer = 28              ' 予備税率5
    Private Const COL_SHEET_BEFORE_TAX_SPARE5 As Integer = 29       ' 修正前予備税率5

    Private Const STR_HYPHEN As String = "-"                        ' ハイフン
    Private Const STR_MINUS1 As String = "-1"

    '' --- 2019/08/13 軽減税率対応 Start E.Okuda@Compass ---
    'Private Const COL_SHEET_TAXS_DT As Integer = 0              ' 開始日
    'Private Const COL_SHEET_TAXE_DT As Integer = 1              ' 終了日
    'Private Const COL_SHEET_TAX_RITU As Integer = 2             ' 消費税率
    'Private Const COL_SHEET_REDUCED_RATE As Integer = 3         ' 軽減税率
    'Private Const COL_SHEET_SEQ As Integer = 4                  ' SEQ
    'Private Const COL_SHEET_UPDATE_KBN As Integer = 5           ' 更新区分
    'Private Const COL_SHEET_BEFORE_TAXS_DTN As Integer = 6      ' 修正前開始日
    'Private Const COL_SHEET_BEFORE_TAXE_DT As Integer = 7       ' 修正前終了日
    'Private Const COL_SHEET_BEFORE_TAX_RITU As Integer = 8      ' 修正前消費税率
    'Private Const COL_SHEET_BEFORE_REDUCED_RATE As Integer = 9  ' 修正前軽減税率

    '' --- 2019/08/13 軽減税率対応 End E.Okuda@Compass ---
    ' --- 2020/03/05 軽減税率対応 End E.Okuda@Compass ---

    'SQL文宣言

    '<sqlid:EX20S001>消費税の初期表示（SELECT）SQL
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
    'Public strSelectTaxmstSearch As String = " SELECT " & vbCrLf &
    '                                         " TAXS_DT " & vbCrLf &
    '                                         ",TAXE_DT " & vbCrLf &
    '                                         ",TAX_RITU " & vbCrLf &
    '                                         ",SEQ " & vbCrLf &
    '                                         ",'1' " & vbCrLf &
    '                                         ",TAXS_DT " & vbCrLf &
    '                                         ",TAXE_DT " & vbCrLf &
    '                                         ",TAX_RITU " & vbCrLf &
    '                                         " FROM TAX_MST " & vbCrLf &
    '                                         " ORDER BY SEQ "
    ' --- 2020/03/04 税区分追加対応 Start E.Okuda@Compass ---
    'Public strSelectTaxmstSearch As String = " SELECT " & vbCrLf &
    '                                         " TAXS_DT " & vbCrLf &
    '                                         ",TAXE_DT " & vbCrLf &
    '                                         ",TAX_RITU " & vbCrLf &
    '                                         ",REDUCED_RATE " & vbCrLf &
    '                                         ",SEQ " & vbCrLf &
    '                                         ",'1' " & vbCrLf &
    '                                         ",TAXS_DT " & vbCrLf &
    '                                         ",TAXE_DT " & vbCrLf &
    '                                         ",TAX_RITU " & vbCrLf &
    '                                         ",REDUCED_RATE " & vbCrLf &
    '                                         " FROM TAX_MST " & vbCrLf &
    '                                         " ORDER BY SEQ "
    ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---
    Public strSelectTaxmstSearch As String = " SELECT " & vbCrLf &
                                             " SEQ " & vbCrLf &
                                             ",'1' as UPDT_KBN" & vbCrLf &
                                             ",TAXS_DT as LATEST_TAXS_DT " & vbCrLf &
                                             ",TAXE_DT as LATEST_TAXE_DT " & vbCrLf &
                                             ",TAXS_DT " & vbCrLf &
                                             ",TAXE_DT " & vbCrLf &
                                             ",TAX_RITU " & vbCrLf &
                                             ",TAX_RITU as LATEST_TAX_RITU " & vbCrLf &
                                             ",CASE WHEN REDUCED_RATE < 0 THEN '-' ELSE to_char(REDUCED_RATE, '999') END as REDUCED_RATE " & vbCrLf &
                                             ",CASE WHEN REDUCED_RATE < 0 THEN '-' ELSE to_char(REDUCED_RATE, '999') END as LATEST_REDUCED_RATE " & vbCrLf &
                                             ",CASE WHEN UNTAXED_RATE < 0 THEN '-' ELSE to_char(UNTAXED_RATE, '999') END as UNTAXED_RATE " & vbCrLf &
                                             ",CASE WHEN UNTAXED_RATE < 0 THEN '-' ELSE to_char(UNTAXED_RATE, '999') END as LATEST_UNTAXED_RATE " & vbCrLf &
                                             ",CASE WHEN TAX_FREE < 0 THEN '-' ELSE to_char(TAX_FREE, '999') END as TAX_FREE " & vbCrLf &
                                             ",CASE WHEN TAX_FREE < 0 THEN '-' ELSE to_char(TAX_FREE, '999') END as LATEST_TAX_FREE " & vbCrLf &
                                             ",CASE WHEN TAX_EXEMPTION < 0 THEN '-' ELSE to_char(TAX_EXEMPTION, '999') END as TAX_EXEMPTION " & vbCrLf &
                                             ",CASE WHEN TAX_EXEMPTION < 0 THEN '-' ELSE to_char(TAX_EXEMPTION, '999') END as LATEST_TAX_EXEMPTION " & vbCrLf &
                                             ",CASE WHEN TAX_OLD1 < 0 THEN '-' ELSE to_char(TAX_OLD1, '999') END as TAX_OLD1 " & vbCrLf &
                                             ",CASE WHEN TAX_OLD1 < 0 THEN '-' ELSE to_char(TAX_OLD1, '999') END as LATEST_TAX_OLD1 " & vbCrLf &
                                             ",CASE WHEN TAX_OLD2 < 0 THEN '-' ELSE to_char(TAX_OLD2, '999') END as TAX_OLD2 " & vbCrLf &
                                             ",CASE WHEN TAX_OLD2 < 0 THEN '-' ELSE to_char(TAX_OLD2, '999') END as LATEST_TAX_OLD2 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE1 < 0 THEN '-' ELSE to_char(TAX_SPARE1, '999') END as TAX_SPARE1 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE1 < 0 THEN '-' ELSE to_char(TAX_SPARE1, '999') END as LATEST_TAX_SPARE1 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE2 < 0 THEN '-' ELSE to_char(TAX_SPARE2, '999') END as TAX_SPARE2 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE2 < 0 THEN '-' ELSE to_char(TAX_SPARE2, '999') END as LATEST_TAX_SPARE2 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE3 < 0 THEN '-' ELSE to_char(TAX_SPARE3, '999') END as TAX_SPARE3 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE3 < 0 THEN '-' ELSE to_char(TAX_SPARE3, '999') END as LATEST_TAX_SPARE3 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE4 < 0 THEN '-' ELSE to_char(TAX_SPARE4, '999') END as TAX_SPARE4 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE4 < 0 THEN '-' ELSE to_char(TAX_SPARE4, '999') END as LATEST_TAX_SPARE4 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE5 < 0 THEN '-' ELSE to_char(TAX_SPARE5, '999') END as TAX_SPARE5 " & vbCrLf &
                                             ",CASE WHEN TAX_SPARE5 < 0 THEN '-' ELSE to_char(TAX_SPARE5, '999') END as LATEST_TAX_SPARE5 " & vbCrLf &
                                             " FROM TAX_MST " & vbCrLf &
                                             " ORDER BY SEQ "
    ' --- 2020/03/04 税区分追加対応 Start E.Okuda@Compass ---

    '<sqlid:EX20I001>消費税の登録（INSERT）SQL
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
    ' 2015.10.09 UPDATE START↓ h.hagiwara 
    'Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf & _
    '                                        " (SEQ, " & vbCrLf & _
    '                                        " TAXS_DT, " & vbCrLf & _
    '                                        " TAXE_DT, " & vbCrLf & _
    '                                        " TAX_RITU, " & vbCrLf & _
    '                                        " ADD_DT, " & vbCrLf & _
    '                                        " ADD_USER_CD, " & vbCrLf & _
    '                                        " UP_DT, " & vbCrLf & _
    '                                        " UP_USER_CD) " & vbCrLf & _
    '                                        " VALUES " & vbCrLf & _
    '                                        " (( SELECT MAX(SEQ) FROM TAX_MST ) + 1 " & vbCrLf & _
    '                                        " , :setTaxsDt " & vbCrLf & _
    '                                        " , :setTaxeDt " & vbCrLf & _
    '                                        " , :setTaxRitu " & vbCrLf & _
    '                                        " , :setAddDt " & vbCrLf & _
    '                                        " , :setAddUserCd " & vbCrLf & _
    '                                        " , :setUpDt " & vbCrLf & _
    '                                        " , :setUpUserCd ); "
    'Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf &
    '                                        " (SEQ, " & vbCrLf &
    '                                        " TAXS_DT, " & vbCrLf &
    '                                        " TAXE_DT, " & vbCrLf &
    '                                        " TAX_RITU, " & vbCrLf &
    '                                        " ADD_DT, " & vbCrLf &
    '                                        " ADD_USER_CD, " & vbCrLf &
    '                                        " UP_DT, " & vbCrLf &
    '                                        " UP_USER_CD) " & vbCrLf &
    '                                        " VALUES " & vbCrLf &
    '                                        " (( SELECT COALESCE(MAX(SEQ),0) + 1 FROM TAX_MST ) " & vbCrLf &
    '                                        " , :setTaxsDt " & vbCrLf &
    '                                        " , :setTaxeDt " & vbCrLf &
    '                                        " , :setTaxRitu " & vbCrLf &
    '                                        " , :setAddDt " & vbCrLf &
    '                                        " , :setAddUserCd " & vbCrLf &
    '                                        " , :setUpDt " & vbCrLf &
    '                                        " , :setUpUserCd ); "
    ' --- 2020/03/04 税区分追加対応 Start E.Okuda@Compass ---
    Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf &
                                            " (SEQ, " & vbCrLf &
                                            " TAXS_DT, " & vbCrLf &
                                            " TAXE_DT, " & vbCrLf &
                                            " TAX_RITU, " & vbCrLf &
                                            " REDUCED_RATE, " & vbCrLf &
                                            " UNTAXED_RATE, " & vbCrLf &
                                            " TAX_FREE, " & vbCrLf &
                                            " TAX_EXEMPTION, " & vbCrLf &
                                            " TAX_OLD1, " & vbCrLf &
                                            " TAX_OLD2, " & vbCrLf &
                                            " TAX_SPARE1, " & vbCrLf &
                                            " TAX_SPARE2, " & vbCrLf &
                                            " TAX_SPARE3, " & vbCrLf &
                                            " TAX_SPARE4, " & vbCrLf &
                                            " TAX_SPARE5, " & vbCrLf &
                                            " ADD_DT, " & vbCrLf &
                                            " ADD_USER_CD, " & vbCrLf &
                                            " UP_DT, " & vbCrLf &
                                            " UP_USER_CD) " & vbCrLf &
                                            " VALUES " & vbCrLf &
                                            " (( SELECT COALESCE(MAX(SEQ),0) + 1 FROM TAX_MST ) " & vbCrLf &
                                            " , :setTaxsDt " & vbCrLf &
                                            " , :setTaxeDt " & vbCrLf &
                                            " , :setTaxRitu " & vbCrLf &
                                            " , :setReducedRate " & vbCrLf &
                                            " , :setUntaxedRate " & vbCrLf &
                                            " , :setTaxFree " & vbCrLf &
                                            " , :setTaxExemption " & vbCrLf &
                                            " , :setTaxOld1 " & vbCrLf &
                                            " , :setTaxOld2 " & vbCrLf &
                                            " , :setTaxSpare1 " & vbCrLf &
                                            " , :setTaxSpare2 " & vbCrLf &
                                            " , :setTaxSpare3 " & vbCrLf &
                                            " , :setTaxSpare4 " & vbCrLf &
                                            " , :setTaxSpare5 " & vbCrLf &
                                            " , :setAddDt " & vbCrLf &
                                            " , :setAddUserCd " & vbCrLf &
                                            " , :setUpDt " & vbCrLf &
                                            " , :setUpUserCd ); "
    'Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf &
    '                                        " (SEQ, " & vbCrLf &
    '                                        " TAXS_DT, " & vbCrLf &
    '                                        " TAXE_DT, " & vbCrLf &
    '                                        " TAX_RITU, " & vbCrLf &
    '                                        " REDUCED_RATE, " & vbCrLf &
    '                                        " ADD_DT, " & vbCrLf &
    '                                        " ADD_USER_CD, " & vbCrLf &
    '                                        " UP_DT, " & vbCrLf &
    '                                        " UP_USER_CD) " & vbCrLf &
    '                                        " VALUES " & vbCrLf &
    '                                        " (( SELECT COALESCE(MAX(SEQ),0) + 1 FROM TAX_MST ) " & vbCrLf &
    '                                        " , :setTaxsDt " & vbCrLf &
    '                                        " , :setTaxeDt " & vbCrLf &
    '                                        " , :setTaxRitu " & vbCrLf &
    '                                        " , :setReducedRate " & vbCrLf &
    '                                        " , :setAddDt " & vbCrLf &
    '                                        " , :setAddUserCd " & vbCrLf &
    '                                        " , :setUpDt " & vbCrLf &
    '                                        " , :setUpUserCd ); "
    ' 2015.10.09 UPDATE END↑ h.hagiwara 
    ' --- 2020/03/04 税区分追加対応 End E.Okuda@Compass ---
    ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

    '<sqlid:EX20U001>消費税マスタの更新（UPDATE）SQL
    ' --- 2020/03/04 税区分追加対応 Start E.Okuda@Compass ---
    Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf &
                                            " TAXS_DT        = :UpdateTaxsDt " & vbCrLf &
                                            ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf &
                                            ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf &
                                            ",REDUCED_RATE   = :UpdateReducedRate " & vbCrLf &
                                            ",UNTAXED_RATE   = :UpdateUntaxedRate " & vbCrLf &
                                            ",TAX_FREE       = :UpdateTaxFree " & vbCrLf &
                                            ",TAX_EXEMPTION  = :UpdateTaxExemption " & vbCrLf &
                                            ",TAX_OLD1       = :UpdateTaxOld1 " & vbCrLf &
                                            ",TAX_OLD2       = :UpdateTaxOld2 " & vbCrLf &
                                            ",TAX_SPARE1     = :UpdateTaxSpare1 " & vbCrLf &
                                            ",TAX_SPARE2     = :UpdateTaxSpare2 " & vbCrLf &
                                            ",TAX_SPARE3     = :UpdateTaxSpare3 " & vbCrLf &
                                            ",TAX_SPARE4     = :UpdateTaxSpare4 " & vbCrLf &
                                            ",TAX_SPARE5     = :UpdateTaxSpare5 " & vbCrLf &
                                            ",UP_DT          = :UpdateUpDt " & vbCrLf &
                                            ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf &
                                            " WHERE SEQ      = :SEQ "
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
    'Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf &
    '                                        " TAXS_DT        = :UpdateTaxsDt " & vbCrLf &
    '                                        ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf &
    '                                        ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf &
    '                                        ",REDUCED_RATE   = :UpdateReducedRate " & vbCrLf &
    '                                        ",UP_DT          = :UpdateUpDt " & vbCrLf &
    '                                        ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf &
    '                                        " WHERE SEQ      = :SEQ "

    'Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf & _
    '                                        " TAXS_DT        = :UpdateTaxsDt " & vbCrLf & _
    '                                        ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf & _
    '                                        ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf & _
    '                                        ",UP_DT          = :UpdateUpDt " & vbCrLf & _
    '                                        ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf & _
    '                                        " WHERE SEQ      = :SEQ "
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---

    ' --- 2020/03/04 税区分追加対応 End E.Okuda@Compass ---

    ' 2015.12.08 ADD START↓ h.hagiwara 
    Private strDeleteTaxmstSql As String = " DELETE FROM TAX_MST " & vbCrLf & _
                                           " WHERE SEQ      = :SEQ "
    ' 2015.12.08 ADD END↑ h.hagiwara 


    ''' <summary>
    ''' 初期表示用検索のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアター消費税マスタから、初期表示時検索を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi 
    ''' </para></remarks>
    Public Function SetSelectInitTaxSearchSql(ByRef Adapter As NpgsqlDataAdapter, _
                                              ByVal Cn As NpgsqlConnection, _
                                              ByVal dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strSelectTaxmstSearch

            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 登録のSQLの設定
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="Cmd">[IN]NpgSqlCommandクラス</param>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="dataEXTM0103">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアター消費税マスタメンテフォームから渡される値をもとに消費税マスタへの登録を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi
    '''    <p>改訂情報：2019/08/09 E.Okuda@Compass 軽減税率対応</p>
    ''' </para></remarks>
    Public Function SetInsertTaxmstSql(ByVal Cn As NpgsqlConnection, _
                                       ByRef Cmd As NpgsqlCommand, _
                                       ByRef j As Integer, _
                                       ByVal dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""                    '実行SQL
        'Dim setAdd_User_CD As String = "u1234567"    'セッション情報（仮）
        'Dim setUp_User_CD As String = "u1234567"     'セッション情報（仮）

        Try
            'SQL文(INSERT)
            strSQL = strInsertTaxmstSql

            'データアダプタに、SQLのINSERT文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("setTaxsDt", NpgsqlTypes.NpgsqlDbType.Varchar))                 '開始日
                .Add(New NpgsqlParameter("setTaxeDt", NpgsqlTypes.NpgsqlDbType.Varchar))                 '終了日
                .Add(New NpgsqlParameter("setTaxRitu", NpgsqlTypes.NpgsqlDbType.Integer))                '消費税割合
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("setReducedRate", NpgsqlTypes.NpgsqlDbType.Integer))             ' 軽減税率
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("setUntaxedRate", NpgsqlTypes.NpgsqlDbType.Integer))               ' 対象外
                .Add(New NpgsqlParameter("setTaxFree", NpgsqlTypes.NpgsqlDbType.Integer))                   ' 非課税
                .Add(New NpgsqlParameter("setTaxExemption", NpgsqlTypes.NpgsqlDbType.Integer))              ' 免税
                .Add(New NpgsqlParameter("setTaxOld1", NpgsqlTypes.NpgsqlDbType.Integer))                   ' 旧課税1
                .Add(New NpgsqlParameter("setTaxOld2", NpgsqlTypes.NpgsqlDbType.Integer))                   ' 旧課税2
                .Add(New NpgsqlParameter("setTaxSpare1", NpgsqlTypes.NpgsqlDbType.Integer))                 ' 予備1
                .Add(New NpgsqlParameter("setTaxSpare2", NpgsqlTypes.NpgsqlDbType.Integer))                 ' 予備2
                .Add(New NpgsqlParameter("setTaxSpare3", NpgsqlTypes.NpgsqlDbType.Integer))                 ' 予備3
                .Add(New NpgsqlParameter("setTaxSpare4", NpgsqlTypes.NpgsqlDbType.Integer))                 ' 予備4
                .Add(New NpgsqlParameter("setTaxSpare5", NpgsqlTypes.NpgsqlDbType.Integer))                 ' 予備5
                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---

                .Add(New NpgsqlParameter("setAddDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                '登録年月日
                .Add(New NpgsqlParameter("setAddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))              '登録ユーザーCD
                .Add(New NpgsqlParameter("setUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("setUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
            End With

            'バインド変数に値をセット
            With Cmd
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                .Parameters("setTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXS_DT).Text      '開始日
                If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXE_DT).Text.Trim() = "" Then
                    .Parameters("setTaxeDt").Value = "2099/12/31"                                        '終了日
                Else : .Parameters("setTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXE_DT).Text
                End If
                .Parameters("setTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_RITU).Text     '終了日更新年月日


                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("setReducedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text = STR_HYPHEN Then
                        .Parameters("setReducedRate").Value = STR_MINUS1
                    Else
                        .Parameters("setReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                    End If
                End If
                'If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                '    .Parameters("setReducedRate").Value = DBNull.Value
                'Else
                '    .Parameters("setReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                'End If

                '.Parameters("setTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text      '開始日
                'If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text.Trim() = "" Then
                '    .Parameters("setTaxeDt").Value = "2099/12/31"                                        '終了日
                'Else : .Parameters("setTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text
                'End If
                '.Parameters("setTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text     '終了日更新年月日
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

                ' 対象外
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text) Then
                    .Parameters("setUntaxedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text = STR_HYPHEN Then
                        .Parameters("setUntaxedRate").Value = STR_MINUS1
                    Else
                        .Parameters("setUntaxedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text
                    End If
                End If

                ' 非課税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text) Then
                    .Parameters("setTaxFree").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text = STR_HYPHEN Then
                        .Parameters("setTaxFree").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxFree").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text
                    End If
                End If

                ' 免税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text) Then
                    .Parameters("setTaxExemption").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text = STR_HYPHEN Then
                        .Parameters("setTaxExemption").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxExemption").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text
                    End If
                End If

                ' 旧課税1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text) Then
                    .Parameters("setTaxOld1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text = STR_HYPHEN Then
                        .Parameters("setTaxOld1").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxOld1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text
                    End If
                End If

                ' 旧課税2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text) Then
                    .Parameters("setTaxOld2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text = STR_HYPHEN Then
                        .Parameters("setTaxOld2").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxOld2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text
                    End If
                End If

                ' 予備税率1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE1).Text) Then
                    .Parameters("setTaxSpare1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE1).Text = STR_HYPHEN Then
                        .Parameters("setTaxSpare1").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxSpare1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE1).Text
                    End If
                End If

                ' 予備税率2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE2).Text) Then
                    .Parameters("setTaxSpare2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE2).Text = STR_HYPHEN Then
                        .Parameters("setTaxSpare2").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxSpare2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE2).Text
                    End If
                End If

                ' 予備税率3
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE3).Text) Then
                    .Parameters("setTaxSpare3").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE3).Text = STR_HYPHEN Then
                        .Parameters("setTaxSpare3").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxSpare3").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE3).Text
                    End If
                End If

                ' 予備税率4
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE4).Text) Then
                    .Parameters("setTaxSpare4").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE4).Text = STR_HYPHEN Then
                        .Parameters("setTaxSpare4").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxSpare4").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE4).Text
                    End If
                End If

                ' 予備税率5
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE5).Text) Then
                    .Parameters("setTaxSpare5").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE5).Text = STR_HYPHEN Then
                        .Parameters("setTaxSpare5").Value = STR_MINUS1
                    Else
                        .Parameters("setTaxSpare5").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_BEFORE_TAX_SPARE5).Text
                    End If
                End If
                ' --- 2020/03/06 税区分追加対応 End E.Okuda@Compass ---

                .Parameters("setAddDt").Value = Now                                                      'メールアドレス
                '.Parameters("setAddUserCd").Value = setAdd_User_CD                                       '画面.ユーザーCD
                .Parameters("setAddUserCd").Value = CommonEXT.PropComStrUserId                            '画面.ユーザーCD
                .Parameters("setUpDt").Value = Now                                                        '更新年月日
                '.Parameters("setUpUserCd").Value = setUp_User_CD                                         '画面.ユーザーCD
                .Parameters("setUpUserCd").Value = CommonEXT.PropComStrUserId                             '画面.ユーザーCD
            End With

            '終了ログ出力()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 終了日自動修正のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0103">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアター消費税マスタに終了日の修正を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi
    '''    <p>改訂情報：2019/08/09 E.Okuda@Compass 軽減税率対応</p>
    ''' </para></remarks>
    Public Function SetUpdateTaxeDtSql(ByVal Cn As NpgsqlConnection, _
                                       ByRef Cmd As NpgsqlCommand, _
                                       ByRef j As Integer, _
                                       ByVal dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""
        Dim setUp_User_CD As String = "u1234567"    'セッション情報（仮）
        Dim StartDt As DateTime

        StartDt = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Value.ToString

        Try

            'SQL文(UPDATE)
            strSQL = strUpdateTaxmstSql

            'データアダプタに、SQLのUPDATE文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("UpdateTaxsDt", NpgsqlTypes.NpgsqlDbType.Varchar))                 '開始日
                .Add(New NpgsqlParameter("UpdateTaxeDt", NpgsqlTypes.NpgsqlDbType.Varchar))                 '終了日
                .Add(New NpgsqlParameter("UpdateTaxRitu", NpgsqlTypes.NpgsqlDbType.Integer))                '消費税割合
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("UpdateReducedRate", NpgsqlTypes.NpgsqlDbType.Integer))             ' 軽減税率
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("UpdateUntaxedRate", NpgsqlTypes.NpgsqlDbType.Integer))            ' 対象外
                .Add(New NpgsqlParameter("UpdateTaxFree", NpgsqlTypes.NpgsqlDbType.Integer))                ' 非課税
                .Add(New NpgsqlParameter("UpdateTaxExemption", NpgsqlTypes.NpgsqlDbType.Integer))           ' 免税
                .Add(New NpgsqlParameter("UpdateTaxOld1", NpgsqlTypes.NpgsqlDbType.Integer))                ' 旧課税1
                .Add(New NpgsqlParameter("UpdateTaxOld2", NpgsqlTypes.NpgsqlDbType.Integer))                ' 旧課税2
                .Add(New NpgsqlParameter("UpdateTaxSpare1", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備1
                .Add(New NpgsqlParameter("UpdateTaxSpare2", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備2
                .Add(New NpgsqlParameter("UpdateTaxSpare3", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備3
                .Add(New NpgsqlParameter("UpdateTaxSpare4", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備4
                .Add(New NpgsqlParameter("UpdateTaxSpare5", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備5
                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---


                .Add(New NpgsqlParameter("UpdateUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("UpdateUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                          '画面.SEQ番号
            End With

            'バインド変数に値をセット
            With Cmd
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                '.Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 0).Text   '開始日
                '.Parameters("UpdateTaxeDt").Value = StartDt.AddDays(-1).ToString("yyyy/MM/dd")               '終了日
                '.Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 2).Text  '消費税割合
                .Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAXS_DT).Text   '開始日
                .Parameters("UpdateTaxeDt").Value = StartDt.AddDays(-1).ToString("yyyy/MM/dd")               '終了日
                .Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_RITU).Text  '消費税割合
                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("UpdateReducedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text Then
                        .Parameters("UpdateReducedRate").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                    End If
                End If
                'If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text) Then
                '    .Parameters("UpdateReducedRate").Value = DBNull.Value
                'Else
                '    .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                'End If
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                ' 対象外
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_UNTAXED_RATE).Text) Then
                    .Parameters("UpdateUntaxedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_UNTAXED_RATE).Text = STR_HYPHEN Then
                        .Parameters("UpdateUntaxedRate").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateUntaxedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_UNTAXED_RATE).Text
                    End If
                End If

                ' 非課税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_FREE).Text) Then
                    .Parameters("UpdateTaxFree").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_FREE).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxFree").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxFree").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_FREE).Text
                    End If
                End If

                ' 免税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_EXEMPTION).Text) Then
                    .Parameters("UpdateTaxExemption").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_EXEMPTION).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxExemption").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxExemption").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_EXEMPTION).Text
                    End If
                End If

                ' 旧課税1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD1).Text) Then
                    .Parameters("UpdateTaxOld1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD1).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxOld1").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxOld1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD1).Text
                    End If
                End If

                ' 旧課税2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD2).Text) Then
                    .Parameters("UpdateTaxOld2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD2).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxOld2").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxOld2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_OLD2).Text
                    End If
                End If

                ' 予備税率1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE1).Text) Then
                    .Parameters("UpdateTaxSpare1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE1).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare1").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE1).Text
                    End If
                End If

                ' 予備税率2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE2).Text) Then
                    .Parameters("UpdateTaxSpare2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE2).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare2").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE2).Text
                    End If
                End If

                ' 予備税率3
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE3).Text) Then
                    .Parameters("UpdateTaxSpare3").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE3).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare3").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare3").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE3).Text
                    End If
                End If

                ' 予備税率4
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE4).Text) Then
                    .Parameters("UpdateTaxSpare4").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE4).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare4").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare4").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_TAX_SPARE4).Text
                    End If
                End If

                ' 予備税率5
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text) Then
                    .Parameters("UpdateTaxSpare5").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare5").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare5").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text
                    End If
                End If
                ' --- 2020/03/06 税区分追加対応 End E.Okuda@Compass ---

                .Parameters("UpdateUpDt").Value = Now                                                        '更新年月日
                .Parameters("UpdateUpUserCd").Value = setUp_User_CD                                          '更新ユーザーCD
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_SEQ).Text            '画面.SEQ番号
            End With

            '終了ログ出力()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 更新のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0103">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアターユーザーマスタにスプレッドシート内の値の更新を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi
    '''    <p>改訂情報：2019/08/09 E.Okuda@Compass 軽減税率対応</p>
    ''' </para></remarks>
    Public Function SetUpdateTaxmstSql(ByVal Cn As NpgsqlConnection, _
                                       ByRef Cmd As NpgsqlCommand, _
                                       ByRef j As Integer, _
                                       ByVal dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""
        'Dim setUp_User_CD As String = "u1234567"    'セッション情報（仮）


        Try

            'SQL文(UPDATE)
            strSQL = strUpdateTaxmstSql

            'データアダプタに、SQLのUPDATE文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("UpdateTaxsDt", NpgsqlTypes.NpgsqlDbType.Varchar))                    '開始日
                .Add(New NpgsqlParameter("UpdateTaxeDt", NpgsqlTypes.NpgsqlDbType.Varchar))                    '終了日
                .Add(New NpgsqlParameter("UpdateTaxRitu", NpgsqlTypes.NpgsqlDbType.Integer))                '消費税割合
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("UpdateReducedRate", NpgsqlTypes.NpgsqlDbType.Integer))             ' 軽減税率
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                .Add(New NpgsqlParameter("UpdateUntaxedRate", NpgsqlTypes.NpgsqlDbType.Integer))            ' 対象外
                .Add(New NpgsqlParameter("UpdateTaxFree", NpgsqlTypes.NpgsqlDbType.Integer))                ' 非課税
                .Add(New NpgsqlParameter("UpdateTaxExemption", NpgsqlTypes.NpgsqlDbType.Integer))           ' 免税
                .Add(New NpgsqlParameter("UpdateTaxOld1", NpgsqlTypes.NpgsqlDbType.Integer))                ' 旧課税1
                .Add(New NpgsqlParameter("UpdateTaxOld2", NpgsqlTypes.NpgsqlDbType.Integer))                ' 旧課税2
                .Add(New NpgsqlParameter("UpdateTaxSpare1", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備1
                .Add(New NpgsqlParameter("UpdateTaxSpare2", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備2
                .Add(New NpgsqlParameter("UpdateTaxSpare3", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備3
                .Add(New NpgsqlParameter("UpdateTaxSpare4", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備4
                .Add(New NpgsqlParameter("UpdateTaxSpare5", NpgsqlTypes.NpgsqlDbType.Integer))              ' 予備5
                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---

                .Add(New NpgsqlParameter("UpdateUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("UpdateUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                          '画面.SEQ番号
            End With

            'バインド変数に値をセット
            With Cmd
                ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
                '.Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text      '開始日
                'If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text.Trim = "" Then                        '終了日(未入力時は2099/12/31をセット)
                '    .Parameters("UpdateTaxeDt").Value = "2099/12/31"
                'Else
                '    .Parameters("UpdateTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text
                'End If

                '.Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text     '消費税割合
                '.Parameters("UpdateUpDt").Value = Now                                                       '更新年月日
                ''.Parameters("UpdateUpUserCd").Value = setUp_User_CD                                        '更新ユーザーCD
                '.Parameters("UpdateUpUserCd").Value = CommonEXT.PropComStrUserId                            '更新ユーザーCD
                '.Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 3).Text               '画面.SEQ番号

                .Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXS_DT).Text      '開始日
                If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXE_DT).Text.Trim = "" Then                        '終了日(未入力時は2099/12/31をセット)
                    .Parameters("UpdateTaxeDt").Value = "2099/12/31"
                Else
                    .Parameters("UpdateTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAXE_DT).Text
                End If

                .Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_RITU).Text     '消費税割合
                ' --- 2020/03/06 税区分追加対応 Start E.Okuda@Compass ---
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("UpdateReducedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text = STR_HYPHEN Then
                        .Parameters("UpdateReducedRate").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text     ' 軽減税率
                    End If
                End If
                'If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                '    .Parameters("UpdateReducedRate").Value = DBNull.Value
                'Else
                '    .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text     ' 軽減税率
                'End If

                ' 対象外
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text) Then
                    .Parameters("UpdateUntaxedRate").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text = STR_HYPHEN Then
                        .Parameters("UpdateUntaxedRate").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateUntaxedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_UNTAXED_RATE).Text
                    End If
                End If

                ' 非課税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text) Then
                    .Parameters("UpdateTaxFree").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxFree").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxFree").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_FREE).Text
                    End If
                End If

                ' 免税
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text) Then
                    .Parameters("UpdateTaxExemption").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxExemption").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxExemption").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_EXEMPTION).Text
                    End If
                End If

                ' 旧課税1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text) Then
                    .Parameters("UpdateTaxOld1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxOld1").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxOld1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD1).Text
                    End If
                End If

                ' 旧課税2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text) Then
                    .Parameters("UpdateTaxOld2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxOld2").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxOld2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_OLD2).Text
                    End If
                End If

                ' 予備税率1
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE1).Text) Then
                    .Parameters("UpdateTaxSpare1").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE1).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare1").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare1").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE1).Text
                    End If
                End If

                ' 予備税率2
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE2).Text) Then
                    .Parameters("UpdateTaxSpare2").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE2).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare2").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare2").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE2).Text
                    End If
                End If

                ' 予備税率3
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE3).Text) Then
                    .Parameters("UpdateTaxSpare3").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE3).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare3").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare3").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE3).Text
                    End If
                End If

                ' 予備税率4
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE4).Text) Then
                    .Parameters("UpdateTaxSpare4").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE4).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare4").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare4").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE4).Text
                    End If
                End If

                ' 予備税率5
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text) Then
                    .Parameters("UpdateTaxSpare5").Value = DBNull.Value
                Else
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text = STR_HYPHEN Then
                        .Parameters("UpdateTaxSpare5").Value = STR_MINUS1
                    Else
                        .Parameters("UpdateTaxSpare5").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_TAX_SPARE5).Text
                    End If
                End If
                ' --- 2020/03/06 税区分追加対応 End E.Okuda@Compass ---

                .Parameters("UpdateUpDt").Value = Now                                                       '更新年月日
                '.Parameters("UpdateUpUserCd").Value = setUp_User_CD                                        '更新ユーザーCD
                .Parameters("UpdateUpUserCd").Value = CommonEXT.PropComStrUserId                            '更新ユーザーCD
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_SEQ).Text               '画面.SEQ番号
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---
            End With

            '終了ログ出力()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 削除のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0103">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアター消費税マスタから指定の消費税データ削除を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' </para></remarks>
    Public Function SetDeleteTaxmstSql(ByVal Cn As NpgsqlConnection, _
                                       ByRef Cmd As NpgsqlCommand, _
                                       ByRef j As Integer, _
                                       ByVal dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(UPDATE)
            strSQL = strDeleteTaxmstSql

            'データアダプタに、SQLのUPDATE文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                          '画面.SEQ番号
            End With

            'バインド変数に値をセット
            With Cmd
                ' --- 2019/08/13 軽減税率対応 Start E.Okuda@Compass ---
                '.Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 3).Text               '画面.SEQ番号
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_SEQ).Text               '画面.SEQ番号
                ' --- 2019/08/13 軽減税率対応 End E.Okuda@Compass ---
            End With

            '終了ログ出力()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function

End Class
