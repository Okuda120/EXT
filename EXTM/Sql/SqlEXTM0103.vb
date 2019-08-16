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

    ' --- 2019/08/13 軽減税率対応 Start E.Okuda@Compass ---
    Private Const COL_SHEET_TAXS_DT As Integer = 0              ' 開始日
    Private Const COL_SHEET_TAXE_DT As Integer = 1              ' 終了日
    Private Const COL_SHEET_TAX_RITU As Integer = 2             ' 消費税率
    Private Const COL_SHEET_REDUCED_RATE As Integer = 3         ' 軽減税率
    Private Const COL_SHEET_SEQ As Integer = 4                  ' SEQ
    Private Const COL_SHEET_UPDATE_KBN As Integer = 5           ' 更新区分
    Private Const COL_SHEET_BEFORE_TAXS_DTN As Integer = 6      ' 修正前開始日
    Private Const COL_SHEET_BEFORE_TAXE_DT As Integer = 7       ' 修正前終了日
    Private Const COL_SHEET_BEFORE_TAX_RITU As Integer = 8      ' 修正前消費税率
    Private Const COL_SHEET_BEFORE_REDUCED_RATE As Integer = 9  ' 修正前軽減税率

    ' --- 2019/08/13 軽減税率対応 End E.Okuda@Compass ---

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
    Public strSelectTaxmstSearch As String = " SELECT " & vbCrLf &
                                             " TAXS_DT " & vbCrLf &
                                             ",TAXE_DT " & vbCrLf &
                                             ",TAX_RITU " & vbCrLf &
                                             ",REDUCED_RATE " & vbCrLf &
                                             ",SEQ " & vbCrLf &
                                             ",'1' " & vbCrLf &
                                             ",TAXS_DT " & vbCrLf &
                                             ",TAXE_DT " & vbCrLf &
                                             ",TAX_RITU " & vbCrLf &
                                             ",REDUCED_RATE " & vbCrLf &
                                             " FROM TAX_MST " & vbCrLf &
                                             " ORDER BY SEQ "
    ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

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
    Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf &
                                            " (SEQ, " & vbCrLf &
                                            " TAXS_DT, " & vbCrLf &
                                            " TAXE_DT, " & vbCrLf &
                                            " TAX_RITU, " & vbCrLf &
                                            " REDUCED_RATE, " & vbCrLf &
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
                                            " , :setAddDt " & vbCrLf &
                                            " , :setAddUserCd " & vbCrLf &
                                            " , :setUpDt " & vbCrLf &
                                            " , :setUpUserCd ); "
    ' 2015.10.09 UPDATE END↑ h.hagiwara 
    ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

    '<sqlid:EX20U001>消費税マスタの更新（UPDATE）SQL
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---
    Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf &
                                            " TAXS_DT        = :UpdateTaxsDt " & vbCrLf &
                                            ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf &
                                            ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf &
                                            ",REDUCED_RATE   = :UpdateReducedRate " & vbCrLf &
                                            ",UP_DT          = :UpdateUpDt " & vbCrLf &
                                            ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf &
                                            " WHERE SEQ      = :SEQ "

    'Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf & _
    '                                        " TAXS_DT        = :UpdateTaxsDt " & vbCrLf & _
    '                                        ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf & _
    '                                        ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf & _
    '                                        ",UP_DT          = :UpdateUpDt " & vbCrLf & _
    '                                        ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf & _
    '                                        " WHERE SEQ      = :SEQ "
    ' --- 2019/08/09 軽減税率対応 Start E.Okuda@Compass ---

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

                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("setReducedRate").Value = DBNull.Value
                Else
                    .Parameters("setReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                End If
                '.Parameters("setTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text      '開始日
                'If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text.Trim() = "" Then
                '    .Parameters("setTaxeDt").Value = "2099/12/31"                                        '終了日
                'Else : .Parameters("setTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text
                'End If
                '.Parameters("setTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text     '終了日更新年月日
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---

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
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("UpdateReducedRate").Value = DBNull.Value
                Else
                    .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, COL_SHEET_REDUCED_RATE).Text  ' 軽減税率
                End If
                ' --- 2019/08/09 軽減税率対応 End E.Okuda@Compass ---
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
                If String.IsNullOrEmpty(dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text) Then
                    .Parameters("UpdateReducedRate").Value = DBNull.Value
                Else
                    .Parameters("UpdateReducedRate").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, COL_SHEET_REDUCED_RATE).Text     ' 軽減税率
                End If

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
