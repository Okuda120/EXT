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

    'SQL文宣言

    '<sqlid:EX20S001>消費税の初期表示（SELECT）SQL
    Public strSelectTaxmstSearch As String = " SELECT " & vbCrLf & _
                                             " TAXS_DT " & vbCrLf & _
                                             ",TAXE_DT " & vbCrLf & _
                                             ",TAX_RITU " & vbCrLf & _
                                             ",SEQ " & vbCrLf & _
                                             ",'1' " & vbCrLf & _
                                             ",TAXS_DT " & vbCrLf & _
                                             ",TAXE_DT " & vbCrLf & _
                                             ",TAX_RITU " & vbCrLf & _
                                             " FROM TAX_MST " & vbCrLf & _
                                             " ORDER BY SEQ "

    '<sqlid:EX20I001>消費税の登録（INSERT）SQL
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
    Private strInsertTaxmstSql As String = " INSERT INTO TAX_MST" & vbCrLf & _
                                            " (SEQ, " & vbCrLf & _
                                            " TAXS_DT, " & vbCrLf & _
                                            " TAXE_DT, " & vbCrLf & _
                                            " TAX_RITU, " & vbCrLf & _
                                            " ADD_DT, " & vbCrLf & _
                                            " ADD_USER_CD, " & vbCrLf & _
                                            " UP_DT, " & vbCrLf & _
                                            " UP_USER_CD) " & vbCrLf & _
                                            " VALUES " & vbCrLf & _
                                            " (( SELECT COALESCE(MAX(SEQ),0) + 1 FROM TAX_MST ) " & vbCrLf & _
                                            " , :setTaxsDt " & vbCrLf & _
                                            " , :setTaxeDt " & vbCrLf & _
                                            " , :setTaxRitu " & vbCrLf & _
                                            " , :setAddDt " & vbCrLf & _
                                            " , :setAddUserCd " & vbCrLf & _
                                            " , :setUpDt " & vbCrLf & _
                                            " , :setUpUserCd ); "
    ' 2015.10.09 UPDATE END↑ h.hagiwara 

    '<sqlid:EX20U001>消費税マスタの更新（UPDATE）SQL
    Private strUpdateTaxmstSql As String = " UPDATE TAX_MST SET " & vbCrLf & _
                                            " TAXS_DT        = :UpdateTaxsDt " & vbCrLf & _
                                            ",TAXE_DT        = :UpdateTaxeDt " & vbCrLf & _
                                            ",TAX_RITU       = :UpdateTaxRitu " & vbCrLf & _
                                            ",UP_DT          = :UpdateUpDt " & vbCrLf & _
                                            ",UP_USER_CD     = :UpdateUpUserCd " & vbCrLf & _
                                            " WHERE SEQ      = :SEQ "
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
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアター消費税マスタメンテフォームから渡される値をもとに消費税マスタへの登録を行うSQL
    ''' <para>作成情報：2015/08/26 hayabuchi
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
                .Add(New NpgsqlParameter("setAddDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                '登録年月日
                .Add(New NpgsqlParameter("setAddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))              '登録ユーザーCD
                .Add(New NpgsqlParameter("setUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("setUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
            End With

            'バインド変数に値をセット
            With Cmd
                .Parameters("setTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text      '開始日
                If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text.Trim() = "" Then
                    .Parameters("setTaxeDt").Value = "2099/12/31"                                        '終了日
                Else : .Parameters("setTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text
                End If
                .Parameters("setTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text     '終了日更新年月日
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
                .Add(New NpgsqlParameter("UpdateUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("UpdateUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                          '画面.SEQ番号
            End With

            'バインド変数に値をセット
            With Cmd
                .Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 0).Text   '開始日
                .Parameters("UpdateTaxeDt").Value = StartDt.AddDays(-1).ToString("yyyy/MM/dd")               '終了日
                .Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 2).Text  '消費税割合
                .Parameters("UpdateUpDt").Value = Now                                                        '更新年月日
                .Parameters("UpdateUpUserCd").Value = setUp_User_CD                                          '更新ユーザーCD
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 3).Text            '画面.SEQ番号
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
                .Add(New NpgsqlParameter("UpdateUpDt", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '更新年月日
                .Add(New NpgsqlParameter("UpdateUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))               '更新ユーザーCD
                .Add(New NpgsqlParameter("SEQ", NpgsqlTypes.NpgsqlDbType.Integer))                          '画面.SEQ番号
            End With

            'バインド変数に値をセット
            With Cmd
                .Parameters("UpdateTaxsDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text      '開始日
                If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text.Trim = "" Then                        '終了日(未入力時は2099/12/31をセット)
                    .Parameters("UpdateTaxeDt").Value = "2099/12/31"
                Else
                    .Parameters("UpdateTaxeDt").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text
                End If

                .Parameters("UpdateTaxRitu").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text     '消費税割合
                .Parameters("UpdateUpDt").Value = Now                                                       '更新年月日
                '.Parameters("UpdateUpUserCd").Value = setUp_User_CD                                        '更新ユーザーCD
                .Parameters("UpdateUpUserCd").Value = CommonEXT.PropComStrUserId                            '更新ユーザーCD
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 3).Text               '画面.SEQ番号
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
                .Parameters("SEQ").Value = dataEXTM0103.PropVwList.Sheets(0).Cells(j, 3).Text               '画面.SEQ番号
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
