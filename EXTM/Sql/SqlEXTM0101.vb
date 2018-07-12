Imports Npgsql
Imports Common

''' <summary>
''' EXシアターユーザー検索画面Sqlクラス
''' </summary>
''' <remarks>EXシアターユーザー検索画面のSQLの作成・設定を行う
''' <para>作成情報：2015/08/11 hayabuchi
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTM0101

    'SQL文宣言

    '<sqlid:EX18S001>ユーザ一覧の初期表示  
    Public strSelectEXTUsrSearch As String = " SELECT " & vbCrLf & _
                                              " FALSE as CHANGEKBN " & vbCrLf & _
                                              ",t1.USER_CD " & vbCrLf & _
                                              ",t1.PW " & vbCrLf & _
                                              ",t1.USER_NM " & vbCrLf & _
                                              ",t1.MAIL " & vbCrLf & _
                                              ",t2.BUSHO_CD " & vbCrLf & _
                                              ",CASE WHEN t1.SHONIN_FLG = '1'" & vbCrLf & _
                                              " THEN TRUE " & vbCrLf & _
                                              " ELSE FALSE " & vbCrLf & _
                                              " END as SHONIN_FLG " & vbCrLf & _
                                              ",CASE WHEN t1.MST_FLG = '1' " & vbCrLf & _
                                              " THEN TRUE " & vbCrLf & _
                                              " ELSE FALSE " & vbCrLf & _
                                              " END as MST_FLG " & vbCrLf & _
                                              ",CASE WHEN t1.STS = '1' " & vbCrLf & _
                                              " THEN TRUE " & vbCrLf & _
                                              " ELSE FALSE " & vbCrLf & _
                                              " END as STS " & vbCrLf & _
                                              " ,t1.ADD_DT " & vbCrLf & _
                                              " ,t1.UP_DT " & vbCrLf & _
                                              " ,'1' " & vbCrLf & _
                                              " FROM USER_MST t1 " & vbCrLf & _
                                              " LEFT JOIN BUSHO_MST t2" & vbCrLf & _
                                              " ON t1.BUSHO_CD = t2.BUSHO_CD " & vbCrLf & _
                                              " WHERE t1.USER_CD <> '000000' "

    '部署名コンボボックス取得（SELECT）SQL
    Public strSelectBushoCD As String = " SELECT " & vbCrLf & _
                                        " BUSHO_CD " & vbCrLf & _
                                        ",BUSHO_NM " & vbCrLf & _
                                        " FROM " & vbCrLf & _
                                        " BUSHO_MST " & vbCrLf & _
                                        " ORDER BY BUSHO_CD "

    '重複登録項目データ数取得（SELECT）SQL
    Public strSelectCountSameUser As String = "SELECT 1 " & vbCrLf & _
                                                 "FROM USER_MST " & vbCrLf & _
                                                 "WHERE USER_CD = :USER_CD "

    '<sqlid:EX18I001>ユーザ一覧の登録（INSERT）SQL
    Private strInsertUsermstSql As String = " INSERT INTO " & vbCrLf & _
                                            " USER_MST " & vbCrLf & _
                                            " ( USER_CD, PW, USER_NM, MAIL, BUSHO_CD, SHONIN_FLG, MST_FLG, STS, ADD_USER_CD, UP_USER_CD, ADD_DT, UP_DT) " & _
                                            " VALUES " & vbCrLf & _
                                            " ( :setUserCD " & vbCrLf & _
                                            " , :setPW " & vbCrLf & _
                                            " , :setKanjiName " & vbCrLf & _
                                            " , :setMail " & vbCrLf & _
                                            " , :setBushoCD " & vbCrLf & _
                                            " , :setShoninFlg " & vbCrLf & _
                                            " , :setMst_Flg " & vbCrLf & _
                                            " , :setSts " & vbCrLf & _
                                            " , :setAdd_User_CD " & vbCrLf & _
                                            " , :setUp_User_CD " & vbCrLf & _
                                            " , :setAdd_DT " & vbCrLf & _
                                            " , :setUp_DT ); "

    '<sqlid:EX18U001>ユーザ一覧の更新（UPDATE）SQL
    Private strUpdateUsermstSql As String = " UPDATE USER_MST SET " & vbCrLf & _
                                            " PW             = :UpdatePW " & vbCrLf & _
                                            ",USER_NM        = :UpdateUSER_NM " & vbCrLf & _
                                            ",MAIL           = :UpdateMAIL " & vbCrLf & _
                                            ",BUSHO_CD       = :UpdateBUSHO_CD " & vbCrLf & _
                                            ",SHONIN_FLG     = :UpdateSHONIN_FLG " & vbCrLf & _
                                            ",MST_FLG        = :UpdateMST_FLG " & vbCrLf & _
                                            ",STS            = :UpdateSTS " & vbCrLf & _
                                            ",UP_USER_CD     = :UpdateUP_USER_CD " & vbCrLf & _
                                            ",UP_DT          = :UpdateUP_DT " & vbCrLf & _
                                            " WHERE USER_CD  = :USER_CD "

    ''' <summary>
    ''' 初期表示用検索のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアターユーザーマスターから、初期表示時検索を行うSQL
    ''' <para>作成情報：2015/08/11 hayabuchi 
    ''' </para></remarks>
    Public Function SetSelectInitEXTUserSearchSql(ByRef Adapter As NpgsqlDataAdapter, _
                                                  ByVal Cn As NpgsqlConnection, _
                                                  ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strSelectEXTUsrSearch & " ORDER BY USER_CD "

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
    ''' 部署名コンボボックス取得用のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>部署名リストにセットするデータの取得を行うSQL
    ''' <para>作成情報：2015/08/11 hayabuchi 
    ''' </para></remarks>
    Public Function SetSelectBushoSql(ByRef Adapter As NpgsqlDataAdapter, _
                                      ByVal Cn As NpgsqlConnection, _
                                      ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""         '実行SQL

        Try

            'SQL文(SELECT)
            strSQL = strSelectBushoCD

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
    ''' 検索のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアターユーザーマスタから、フォームから渡される値をもとに検索を行うSQL
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' </para></remarks>
    Public Function SetSelectEXTUserSearchSql(ByRef Adapter As NpgsqlDataAdapter, _
                                              ByVal Cn As NpgsqlConnection, _
                                              ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String              '実行SQL
        Dim strWhere As String = ""       '実行WHERE句
        Dim param As Npgsql.NpgsqlParameter() = New Npgsql.NpgsqlParameter(3) {}

        Try

            'SQLに渡すデータの設定
            param(0) = New NpgsqlParameter("setUserId", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(1) = New NpgsqlParameter("setKanjiName", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(2) = New NpgsqlParameter("setMail", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(3) = New NpgsqlParameter("setBushoName", NpgsqlTypes.NpgsqlDbType.Varchar)

            param(0).Value = dataEXTM0101.PropTxtSearchUserID.Text
            param(1).Value = "%" & dataEXTM0101.PropTxtSearchKanjiName.Text & "%"
            param(2).Value = "%" & dataEXTM0101.PropTxtSearchMail.Text & "%"
            param(3).Value = dataEXTM0101.PropCmbSearchBushoName.SelectedValue

            'SQL文(SELECT)
            strSQL = strSelectEXTUsrSearch

            'Where句作成
            If dataEXTM0101.PropTxtSearchUserID.Text <> System.String.Empty Then
                'ユーザーIDが入力されている
                strWhere &= " AND t1.USER_CD = :setUserId "
            End If
            If dataEXTM0101.PropTxtSearchKanjiName.Text <> System.String.Empty Then
                '漢字氏名が入力されている
                strWhere &= " AND t1.USER_NM LIKE :setKanjiName "
            End If
            If dataEXTM0101.PropTxtSearchMail.Text <> System.String.Empty Then
                'メールアドレスが入力されている
                strWhere &= " AND t1.MAIL LIKE :setMail "
            End If
            If dataEXTM0101.PropCmbSearchBushoName.Text <> System.String.Empty Then
                '部署名が入力されている
                strWhere &= " AND t1.BUSHO_CD = :setBushoName "
            End If
            If dataEXTM0101.PropChkShoninFlg.Checked = True Then
                '承認者のみ表示が選択されている
                strWhere &= " AND t1.SHONIN_FLG = '1' "
            End If
            If dataEXTM0101.PropChkMstFlg.Checked = True Then
                'マスタ操作可能者のみ表示が選択されている
                strWhere &= " AND t1.MST_FLG = '1' "
            End If
            If dataEXTM0101.PropChkStsFlg.Checked = False Then
                '無効なユーザーも含めて表示が選択されている
                strWhere &= " AND t1.STS = '0' "
            End If

            'データアダプタに、SQLを設定
            strSQL &= strWhere & " ORDER BY USER_CD "
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値をセット()
            Adapter.SelectCommand.Parameters.AddRange(param)

            '終了ログ出力()
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
    ''' 【共通】同じユーザー名のデータ有無取得用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0101">[IN]ユーザーマスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>同じユーザー名のデータ有無取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function SetSelectCountSameUserSql(ByRef Adapter As NpgsqlDataAdapter, _
                                              ByVal Cn As NpgsqlConnection, _
                                              ByRef j As Integer, _
                                              ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strSelectCountSameUser

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)



            'バインド変数に型をセット
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter("USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                '画面.ユーザー名
            End With

            'バインド変数に値をセット
            With Adapter.SelectCommand
                .Parameters("USER_CD").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(j, 1).Value    '画面.ユーザー名
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理

            Return False
        End Try

    End Function
    ''' <summary>
    ''' 新規登録のSQLの設定
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアターユーザーマスタにスプレッドシート内の値の新規登録を行うSQL
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' </para></remarks>
    Public Function SetInsertUsermstSql(ByVal Cn As NpgsqlConnection, _
                                        ByRef Cmd As NpgsqlCommand, _
                                        ByRef i As Integer, _
                                        ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""                '実行SQL
        Dim setAdd_User_CD As String = "1234"    'セッション情報（仮）
        Dim setUp_User_CD As String = "1234"     'セッション情報（仮）

        Try
            'SQL文(INSERT)
            strSQL = strInsertUsermstSql

            'データアダプタに、SQLのINSERT文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("setUserCD", NpgsqlTypes.NpgsqlDbType.Varchar))                    'ユーザーID
                .Add(New NpgsqlParameter("setPW", NpgsqlTypes.NpgsqlDbType.Varchar))                        'パスワード
                .Add(New NpgsqlParameter("setKanjiName", NpgsqlTypes.NpgsqlDbType.Varchar))                 '漢字氏名
                .Add(New NpgsqlParameter("setMail", NpgsqlTypes.NpgsqlDbType.Varchar))                      'メールアドレス
                .Add(New NpgsqlParameter("setBushoCD", NpgsqlTypes.NpgsqlDbType.Varchar))                   '部署名
                .Add(New NpgsqlParameter("setShoninFlg", NpgsqlTypes.NpgsqlDbType.Char))                    '承認者権限フラグ
                .Add(New NpgsqlParameter("setMst_Flg", NpgsqlTypes.NpgsqlDbType.Char))                      'マスタ操作フラグ
                .Add(New NpgsqlParameter("setSts", NpgsqlTypes.NpgsqlDbType.Char))                          'ステータス
                .Add(New NpgsqlParameter("setAdd_User_CD", NpgsqlTypes.NpgsqlDbType.Varchar))               '初回登録者ユーザーCD
                .Add(New NpgsqlParameter("setUp_User_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                '最終登録者ユーザーCD
                .Add(New NpgsqlParameter("setAdd_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                  '初回登録日
                .Add(New NpgsqlParameter("setUp_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                   '最終登録日
            End With

            'バインド変数に値をセット
            With Cmd
                .Parameters("setUserCD").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 1).Text         'ユーザーID
                .Parameters("setPW").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 2).Text             'パスワード
                .Parameters("setKanjiName").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 3).Text      '漢字氏名
                .Parameters("setMail").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 4).Text           'メールアドレス
                .Parameters("setBushoCD").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 5).Value       '部署名
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 6).Value = True Then                          '承認者権限フラグ
                    .Parameters("setShoninFlg").Value = 1
                Else
                    .Parameters("setShoninFlg").Value = 0
                End If
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 7).Value = True Then                          'マスタ操作フラグ
                    .Parameters("setMst_Flg").Value = 1
                Else
                    .Parameters("setMst_Flg").Value = 0
                End If
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 8).Value = True Then                          'ステータス
                    .Parameters("setSts").Value = 1
                Else
                    .Parameters("setSts").Value = 0
                End If
                .Parameters("setAdd_User_CD").Value = setAdd_User_CD                                        '画面.ユーザーCD
                .Parameters("setUp_User_CD").Value = setUp_User_CD                                          '画面.ユーザーCD
                .Parameters("setAdd_DT").Value = Now                                                        '初回登録日
                .Parameters("setUp_DT").Value = Now                                                         '最終登録日
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
    ''' <param name="dataHBKZ0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>EXシアターユーザーマスタにスプレッドシート内の値の更新を行うSQL
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' </para></remarks>
    Public Function SetUpdateUsermstSql(ByVal Cn As NpgsqlConnection, _
                                        ByRef Cmd As NpgsqlCommand, _
                                        ByRef i As Integer, _
                                        ByVal dataEXTM0101 As DataEXTM0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""
        Dim setUp_User_CD As String = "1234"    'セッション情報（仮）

        Try
            'SQL文(UPDATE)
            strSQL = strUpdateUsermstSql

            'データアダプタに、SQLのUPDATE文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("UpdatePW", NpgsqlTypes.NpgsqlDbType.Varchar))                      'パスワード
                .Add(New NpgsqlParameter("UpdateUSER_NM", NpgsqlTypes.NpgsqlDbType.Varchar))                 '漢字氏名
                .Add(New NpgsqlParameter("UpdateMAIL", NpgsqlTypes.NpgsqlDbType.Varchar))                    'メールアドレス
                .Add(New NpgsqlParameter("UpdateBUSHO_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                '部署名
                .Add(New NpgsqlParameter("UpdateSHONIN_FLG", NpgsqlTypes.NpgsqlDbType.Char))                 '承認者権限フラグ
                .Add(New NpgsqlParameter("UpdateMST_FLG", NpgsqlTypes.NpgsqlDbType.Char))                    'マスタ操作フラグ
                .Add(New NpgsqlParameter("UpdateSTS", NpgsqlTypes.NpgsqlDbType.Char))                        'ステータス
                .Add(New NpgsqlParameter("UpdateUP_USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))              '最終登録者ユーザーCD
                .Add(New NpgsqlParameter("UpdateUP_DT", NpgsqlTypes.NpgsqlDbType.Timestamp))                 '最終登録日
                .Add(New NpgsqlParameter("USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                       '画面.ユーザーCD
            End With

            'バインド変数に値をセット
            With Cmd
                .Parameters("UpdatePW").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 2).Text           'パスワード
                .Parameters("UpdateUSER_NM").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 3).Text      '漢字氏名
                .Parameters("UpdateMAIL").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 4).Text         'メールアドレス
                .Parameters("UpdateBUSHO_CD").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 5).Value    '部署名
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 6).Text = True Then                            '承認者権限フラグ
                    .Parameters("UpdateSHONIN_FLG").Value = 1
                Else
                    .Parameters("UpdateSHONIN_FLG").Value = 0
                End If
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 7).Text = True Then                            'マスタ操作フラグ
                    .Parameters("UpdateMST_FLG").Value = 1
                Else
                    .Parameters("UpdateMST_FLG").Value = 0
                End If
                If dataEXTM0101.PropVwList.Sheets(0).Cells(i, 8).Text = True Then                            'ステータス
                    .Parameters("UpdateSTS").Value = 1
                Else
                    .Parameters("UpdateSTS").Value = 0
                End If
                .Parameters("UpdateUP_USER_CD").Value = setUp_User_CD                                        '画面.ユーザーCD
                .Parameters("UpdateUP_DT").Value = Now                                                       '最終登録日
                .Parameters("USER_CD").Value = dataEXTM0101.PropVwList.Sheets(0).Cells(i, 1).Text            'ユーザーCD
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


