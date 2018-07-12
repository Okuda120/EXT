Imports Npgsql
Imports Common
Imports CommonEXT

''' <summary>
''' 利用者情報登録/詳細画面SQLクラス
''' </summary>
''' <remarks>利用者情報登録/詳細画面のSQLを定義する
''' <para>作成情報：2015/08/25 ozawa
''' <p>改訂情報：</p>
''' </para>
''' </remarks>
Public Class SqlEXTM0202
    '利用者情報取得SQL　
    Private strGetRiyoshaData As String = "SELECT " & _
                                          "t1.RIYOSHA_CD," & _
                                          "t1.RIYO_NM, " & _
                                          "t1.RIYO_KANA, " & _
                                          "COALESCE(t1.DAIHYO_NM,''), " & _
                                          "t1.RIYO_TEL11, " & _
                                          "t1.RIYO_TEL12 , " & _
                                          "t1.RIYO_TEL13 , " & _
                                          "COALESCE(t1.RIYO_NAISEN,'')," & _
                                          "COALESCE(t1.RIYO_FAX11,'')," & _
                                          "COALESCE(t1.RIYO_FAX12,'') ," & _
                                          "COALESCE(t1.RIYO_FAX13,'') ," & _
                                          "t1.RIYO_YUBIN1," & _
                                          "t1.RIYO_YUBIN2 ," & _
                                          "t1.RIYO_TODO ," & _
                                          "t1.RIYO_SHIKU ," & _
                                          "t1.RIYO_BAN ," & _
                                          "COALESCE(t1.RIYO_BUILD,'') ," & _
                                          "t1.RIYO_LVL," & _
                                          "COALESCE(t1.AITE_CD,'')," & _
                                          "COALESCE(t2.AITE_NM,'')," & _
                                          "COALESCE(t1.COM,'') " & _
                                          "FROM RIYOSHA_MST t1 " & _
                                          "LEFT JOIN AITESAKI_MST t2  " & _
                                          "ON t1.AITE_CD = t2.AITE_CD "

    '利用コメント取得用SQL　
    Private strGetComment As String = "SELECT t2.開始日, t2.終了日, " & _
                                      "CASE t1.SHISETU_KBN " & _
                                      "WHEN '1' THEN t1.SAIJI_NM " & _
                                      "WHEN '2' THEN t1.SHUTSUEN_NM " & _
                                      "END 催事名アーティスト名," & _
                                      "t1.RIYO_COM " & _
                                      "FROM YOYAKU_TBL t1 " & _
                                      "LEFT JOIN " & _
                                      "( SELECT " & _
                                      "MIN(YOYAKU_DT) 開始日, " & _
                                      "MAX(YOYAKU_DT) 終了日, " & _
                                      "YOYAKU_NO " & _
                                      "FROM YDT_TBL " & _
                                      "GROUP BY YOYAKU_NO ) t2 " & _
                                      "ON  t1.YOYAKU_NO = t2.YOYAKU_NO "

    '新規登録時、利用者番号を採番するSQL
    'システム日時を取得するSQL
    Private strSelectSysdate As String = "SELECT NOW() SYSDATE "

    '更新用SQL
    Private strUpdateRiyoshaData As String = "UPDATE RIYOSHA_MST " & _
                                             "SET RIYO_NM = :setRiyoNm, " & _
                                             "RIYO_KANA = :setRiyoKn, " & _
                                             "DAIHYO_NM = :setDaihyo," & _
                                             "RIYO_LVL = :setRiyoLvl, " & _
                                             "RIYO_TEL11 = :setTel1, " & _
                                             "RIYO_TEL12 = :setTel2, " & _
                                             "RIYO_TEL13 = :setTel3, " & _
                                             "RIYO_NAISEN = :setNaisen, " & _
                                             "RIYO_FAX11 = :setFax1, " & _
                                             "RIYO_FAX12 = :setFax2," & _
                                             "RIYO_FAX13 = :setFax3," & _
                                             "RIYO_YUBIN1 = :setYubin1, " & _
                                             "RIYO_YUBIN2 = :setYubin2, " & _
                                             "RIYO_TODO = :setToDo, " & _
                                             "RIYO_SHIKU = :setShiku, " & _
                                             "RIYO_BAN = :setBanchi, " & _
                                             "RIYO_BUILD = :setBuild, " & _
                                             "AITE_CD = :setAiteCd, " & _
                                             "COM = :setCom, " & _
                                             "UP_DT = :setSysDate, " & _
                                             "UP_USER_CD = :setUpUserCd " & _
                                             "WHERE RIYOSHA_CD = :setRiyoCd "

    '新規利用者番号採番用SQL
    Private strGetNewRiyoCd As String = "SELECT TRIM(TO_CHAR(NEXTVAL('EXTS0001'),'00000')) AS RIYOSHA_CD ,NOW() SYSDATE "

    '新規登録用SQL
    Private strInsertRiyoshaData As String = "INSERT INTO RIYOSHA_MST ( " & _
                                             "RIYOSHA_CD," & _
                                             "RIYO_NM, " & _
                                             "RIYO_KANA, " & _
                                             "DAIHYO_NM," & _
                                             "RIYO_LVL, " & _
                                             "RIYO_TEL11, " & _
                                             "RIYO_TEL12 , " & _
                                             "RIYO_TEL13 , " & _
                                             "RIYO_NAISEN , " & _
                                             "RIYO_FAX11 , " & _
                                             "RIYO_FAX12," & _
                                             "RIYO_FAX13," & _
                                             "RIYO_YUBIN1, " & _
                                             "RIYO_YUBIN2, " & _
                                             "RIYO_TODO,  " & _
                                             "RIYO_SHIKU,  " & _
                                             "RIYO_BAN , " & _
                                             "RIYO_BUILD , " & _
                                             "AITE_CD, " & _
                                             "COM , " & _
                                             "ADD_DT, " & _
                                             "ADD_USER_CD, " & _
                                             "UP_DT, " & _
                                             "UP_USER_CD " & _
                                             ")VALUES( " & _
                                             " :setRiyoCd" & _
                                             " ,:setRiyoNm" & _
                                             " ,:setRiyoKn " & _
                                             " ,:setDaihyo " & _
                                             " ,:setRiyoLvl " & _
                                             " ,:setTel1 " & _
                                             " ,:setTel2 " & _
                                             " ,:setTel3 " & _
                                             " ,:setNaisen " & _
                                             " ,:setFax1 " & _
                                             " ,:setFax2 " & _
                                             " ,:setFax3 " & _
                                             " ,:setYubin1 " & _
                                             " ,:setYubin2 " & _
                                             " ,:setToDo " & _
                                             " ,:setShiku " & _
                                             " ,:setBanchi " & _
                                             " ,:setBuild " & _
                                             " ,:setAiteCd " & _
                                             " ,:setCom " & _
                                             " ,:setSysDate " & _
                                             " ,:setUpUserCd " & _
                                             " ,:setSysDate " & _
                                             " ,:setUpUserCd )"


    ''' <summary>
    ''' 利用者情報取得用SQL
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>利用者情報取得用SQL
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function GetRiyoshaData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim param As Npgsql.NpgsqlParameter() = New Npgsql.NpgsqlParameter(0) {}

        Try
            'SQLに渡すデータの設定
            param(0) = New NpgsqlParameter("setRiyoCd", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(0).Value = DataEXTM0202.PropLblRiyo_cd.Text

            'SQL文(SELECT)
            strSQL = strGetRiyoshaData
            strWhere = ""

            'Where句作成
            strWhere = "WHERE t1.RIYOSHA_CD = :setRiyoCd "


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

    ''' <summary>
    ''' 利用コメント取得用SQL
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>利用コメント取得用SQL
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function GetComment(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim param As Npgsql.NpgsqlParameter() = New Npgsql.NpgsqlParameter(0) {}

        Try
            'SQLに渡すデータの設定
            param(0) = New NpgsqlParameter("setRiyoCd", NpgsqlTypes.NpgsqlDbType.Varchar)
            param(0).Value = dataEXTM0202.PropLblRiyo_cd.Text

            'SQL文(SELECT)
            strSQL = strGetComment
            strWhere = ""
            'Where句作成
            strWhere = "WHERE RIYOSHA_CD = :setRiyoCd "

            '並べ替える
            strWhere &= "ORDER BY t2.開始日 "

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

    ''' <summary>サーバー日付取得用SQL
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns>サーバー日付取得用SQLを定義する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </returns>
    ''' <remarks></remarks>
    Public Function SelectSysdate(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim strSQL As String
        Dim param As Npgsql.NpgsqlParameter() = New Npgsql.NpgsqlParameter(0) {}

        Try
            'SQL文設定
            strSQL = strSelectSysdate
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

    ''' <summary>利用者情報更新用SQL
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>利用者情報更新用SQLを定義する
    ''' <para>作成情報：2015/08/25　ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function UpdateRiyoshaData(ByRef Cmd As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim strSQL As String
        'セッション.登録ユーザーCD（単体テスト用）
        Dim UpUserCd As String = "1122"

        Try
            'SQL文(UPDATE)
            strSQL = strUpdateRiyoshaData


            'データアダプタに、SQLを設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数にセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("setRiyoCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用者番号
                .Add(New NpgsqlParameter("setRiyoNm", NpgsqlTypes.NpgsqlDbType.Varchar))    '利用者名
                .Add(New NpgsqlParameter("setRiyoKn", NpgsqlTypes.NpgsqlDbType.Varchar))     '利用者カナ
                .Add(New NpgsqlParameter("setDaihyo", NpgsqlTypes.NpgsqlDbType.Varchar))     '代表者名
                .Add(New NpgsqlParameter("setRiyoLvl", NpgsqlTypes.NpgsqlDbType.Char))       '利用レベル
                .Add(New NpgsqlParameter("setTel1", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号1
                .Add(New NpgsqlParameter("setTel2", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号2
                .Add(New NpgsqlParameter("setTel3", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号3
                .Add(New NpgsqlParameter("setNaisen", NpgsqlTypes.NpgsqlDbType.Varchar)) '内線
                .Add(New NpgsqlParameter("setFax1", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX1
                .Add(New NpgsqlParameter("setFax2", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX2
                .Add(New NpgsqlParameter("setFax3", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX3
                .Add(New NpgsqlParameter("setYubin1", NpgsqlTypes.NpgsqlDbType.Varchar)) '郵便番号1
                .Add(New NpgsqlParameter("setYubin2", NpgsqlTypes.NpgsqlDbType.Varchar)) '郵便番号2
                .Add(New NpgsqlParameter("setToDo", NpgsqlTypes.NpgsqlDbType.Varchar)) '都道府県
                .Add(New NpgsqlParameter("setShiku", NpgsqlTypes.NpgsqlDbType.Varchar)) '市区町村
                .Add(New NpgsqlParameter("setBanchi", NpgsqlTypes.NpgsqlDbType.Varchar)) '番地
                .Add(New NpgsqlParameter("setBuild", NpgsqlTypes.NpgsqlDbType.Varchar)) 'ビル名
                .Add(New NpgsqlParameter("setAiteCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '相手先コード
                .Add(New NpgsqlParameter("setCom", NpgsqlTypes.NpgsqlDbType.Varchar)) 'コメント
                .Add(New NpgsqlParameter("setSysDate", NpgsqlTypes.NpgsqlDbType.Timestamp)) 'システム日付
                .Add(New NpgsqlParameter("setUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar)) 'ユーザーCD
            End With

            With Cmd
                .Parameters("setRiyoCd").Value = dataEXTM0202.PropLblRiyo_cd.Text           '利用者番号
                .Parameters("setRiyoNm").Value = dataEXTM0202.PropTxtRiyo_nm.Text            '利用者名
                .Parameters("setRiyoKn").Value = dataEXTM0202.PropTxtRiyo_kana.Text          '利用者カナ

                If dataEXTM0202.PropTxtDaihyo_nm.Text = "" Then
                    .Parameters("setDaihyo").Value = DBNull.Value
                Else
                    .Parameters("setDaihyo").Value = dataEXTM0202.PropTxtDaihyo_nm.Text    '代表者名
                End If

                If dataEXTM0202.PropRdoTujyo.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "1"                                     '利用者レベル
                ElseIf dataEXTM0202.PropRdoChui.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "2"
                ElseIf dataEXTM0202.PropRdoHuka.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "3"
                End If

                .Parameters("setTel1").Value = dataEXTM0202.PropTxtTel1.Text '電話番号1
                .Parameters("setTel2").Value = dataEXTM0202.PropTxtTel2.Text '電話番号2
                .Parameters("setTel3").Value = dataEXTM0202.PropTxtTel3.Text '電話番号3

                If dataEXTM0202.PropTxtNaisen.Text = "" Then
                    .Parameters("setNaisen").Value = DBNull.Value
                Else
                    .Parameters("setNaisen").Value = dataEXTM0202.PropTxtNaisen.Text '内線
                End If

                If dataEXTM0202.PropTxtFax1.Text = "" Then
                    .Parameters("setFax1").Value = DBNull.Value
                Else
                    .Parameters("setFax1").Value = dataEXTM0202.PropTxtFax1.Text 'FAX1
                End If

                If dataEXTM0202.PropTxtFax2.Text = "" Then
                    .Parameters("setFax2").Value = DBNull.Value
                Else
                    .Parameters("setFax2").Value = dataEXTM0202.PropTxtFax2.Text 'FAX2
                End If

                If dataEXTM0202.PropTxtFax3.Text = "" Then
                    .Parameters("setFax3").Value = DBNull.Value
                Else
                    .Parameters("setFax3").Value = dataEXTM0202.PropTxtFax3.Text 'FAX3
                End If

                .Parameters("setYubin1").Value = dataEXTM0202.PropTxtYubin1.Text '郵便番号1
                .Parameters("setYubin2").Value = dataEXTM0202.PropTxtYubin2.Text '郵便番号2
                .Parameters("setToDo").Value = dataEXTM0202.PropCmbTodo.Text '都道府県
                .Parameters("setShiku").Value = dataEXTM0202.PropTxtShiku.Text '市区町村
                .Parameters("setBanchi").Value = dataEXTM0202.PropTxtBanchi.Text '番地

                If dataEXTM0202.PropTxtBuild.Text = "" Then
                    .Parameters("setBuild").Value = DBNull.Value
                Else
                    .Parameters("setBuild").Value = dataEXTM0202.PropTxtBuild.Text 'ビル名
                End If

                .Parameters("setAiteCd").Value = dataEXTM0202.PropLblAite_cd.Text '相手先コード

                If dataEXTM0202.PropTxtCom.Text = "" Then
                    .Parameters("setCom").Value = DBNull.Value
                Else
                    .Parameters("setCom").Value = dataEXTM0202.PropTxtCom.Text 'コメント
                End If

                .Parameters("setSysDate").Value = dataEXTM0202.PropDtmSysDate 'システム日付
                .Parameters("setUpUserCd").Value = UpUserCd  'ユーザーCD
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>新規利用者番号採番
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>新規利用者番号採番用SQLを定義する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function GetNewRiyoCd(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strGetNewRiyoCd

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
            puErrMsg = ex.Message
            Return False
        End Try
    End Function


    ''' <summary> 利用者情報新規登録SQL
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>利用者情報新規登録SQLを定義する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function InsertRiyoshaData(ByRef Cmd As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '変数宣言
        Dim strSQL As String = ""
        'セッション.登録ユーザーCD（単体テスト用）
        Dim UpUserCd As String = "1122"

        Try
            'SQL文(INSERT)
            strSQL = strInsertRiyoshaData

            'データアダプタに、SQLのINSERT文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数にセット
            With Cmd.Parameters
                .Add(New NpgsqlParameter("setRiyoCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '利用者番号
                .Add(New NpgsqlParameter("setRiyoNm", NpgsqlTypes.NpgsqlDbType.Varchar)) '利用者名
                .Add(New NpgsqlParameter("setRiyoKn", NpgsqlTypes.NpgsqlDbType.Varchar)) '利用者カナ
                .Add(New NpgsqlParameter("setDaihyo", NpgsqlTypes.NpgsqlDbType.Varchar)) '代表者名
                .Add(New NpgsqlParameter("setRiyoLvl", NpgsqlTypes.NpgsqlDbType.Char)) '利用レベル
                .Add(New NpgsqlParameter("setTel1", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号1
                .Add(New NpgsqlParameter("setTel2", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号2
                .Add(New NpgsqlParameter("setTel3", NpgsqlTypes.NpgsqlDbType.Varchar)) '電話番号3
                .Add(New NpgsqlParameter("setNaisen", NpgsqlTypes.NpgsqlDbType.Varchar)) '内線
                .Add(New NpgsqlParameter("setFax1", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX1
                .Add(New NpgsqlParameter("setFax2", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX2
                .Add(New NpgsqlParameter("setFax3", NpgsqlTypes.NpgsqlDbType.Varchar)) 'FAX3
                .Add(New NpgsqlParameter("setYubin1", NpgsqlTypes.NpgsqlDbType.Varchar)) '郵便番号1
                .Add(New NpgsqlParameter("setYubin2", NpgsqlTypes.NpgsqlDbType.Varchar)) '郵便番号2
                .Add(New NpgsqlParameter("setToDo", NpgsqlTypes.NpgsqlDbType.Varchar)) '都道府県
                .Add(New NpgsqlParameter("setShiku", NpgsqlTypes.NpgsqlDbType.Varchar)) '市区町村
                .Add(New NpgsqlParameter("setBanchi", NpgsqlTypes.NpgsqlDbType.Varchar)) '番地
                .Add(New NpgsqlParameter("setBuild", NpgsqlTypes.NpgsqlDbType.Varchar)) 'ビル名
                .Add(New NpgsqlParameter("setAiteCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '相手先コード
                .Add(New NpgsqlParameter("setCom", NpgsqlTypes.NpgsqlDbType.Varchar)) 'コメント
                .Add(New NpgsqlParameter("setSysDate", NpgsqlTypes.NpgsqlDbType.Timestamp)) 'システム日付
                .Add(New NpgsqlParameter("setUpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar)) 'ユーザーCD

            End With

            With Cmd
                .Parameters("setRiyoCd").Value = dataEXTM0202.PropNewRiyo_cd      '利用者番号
                .Parameters("setRiyoNm").Value = dataEXTM0202.PropTxtRiyo_nm.Text      '利用者名
                .Parameters("setRiyoKn").Value = dataEXTM0202.PropTxtRiyo_kana.Text    '利用者カナ

                If dataEXTM0202.PropTxtDaihyo_nm.Text = "" Then
                    .Parameters("setDaihyo").Value = DBNull.Value
                Else
                    .Parameters("setDaihyo").Value = dataEXTM0202.PropTxtDaihyo_nm.Text    '代表者名
                End If

                If dataEXTM0202.PropRdoTujyo.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "1"                                     '利用者レベル
                ElseIf dataEXTM0202.PropRdoChui.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "2"
                ElseIf dataEXTM0202.PropRdoHuka.Checked = True Then
                    .Parameters("setRiyoLvl").Value = "3"
                End If

                .Parameters("setTel1").Value = dataEXTM0202.PropTxtTel1.Text '電話番号1
                .Parameters("setTel2").Value = dataEXTM0202.PropTxtTel2.Text '電話番号2
                .Parameters("setTel3").Value = dataEXTM0202.PropTxtTel3.Text '電話番号3

                If dataEXTM0202.PropTxtNaisen.Text = "" Then
                    .Parameters("setNaisen").Value = DBNull.Value
                Else
                    .Parameters("setNaisen").Value = dataEXTM0202.PropTxtNaisen.Text '内線
                End If

                If dataEXTM0202.PropTxtFax1.Text = "" Then
                    .Parameters("setFax1").Value = DBNull.Value
                Else
                    .Parameters("setFax1").Value = dataEXTM0202.PropTxtFax1.Text 'FAX1
                End If

                If dataEXTM0202.PropTxtFax2.Text = "" Then
                    .Parameters("setFax2").Value = DBNull.Value
                Else
                    .Parameters("setFax2").Value = dataEXTM0202.PropTxtFax2.Text 'FAX2
                End If

                If dataEXTM0202.PropTxtFax3.Text = "" Then
                    .Parameters("setFax3").Value = DBNull.Value
                Else
                    .Parameters("setFax3").Value = dataEXTM0202.PropTxtFax3.Text 'FAX3
                End If

                .Parameters("setYubin1").Value = dataEXTM0202.PropTxtYubin1.Text '郵便番号1
                .Parameters("setYubin2").Value = dataEXTM0202.PropTxtYubin2.Text '郵便番号2
                .Parameters("setToDo").Value = dataEXTM0202.PropCmbTodo.Text '都道府県
                .Parameters("setShiku").Value = dataEXTM0202.PropTxtShiku.Text '市区町村
                .Parameters("setBanchi").Value = dataEXTM0202.PropTxtBanchi.Text '番地

                If dataEXTM0202.PropTxtBuild.Text = "" Then
                    .Parameters("setBuild").Value = DBNull.Value
                Else
                    .Parameters("setBuild").Value = dataEXTM0202.PropTxtBuild.Text 'ビル名
                End If

                .Parameters("setAiteCd").Value = dataEXTM0202.PropLblAite_cd.Text '相手先コード

                If dataEXTM0202.PropTxtCom.Text = "" Then
                    .Parameters("setCom").Value = DBNull.Value
                Else
                    .Parameters("setCom").Value = dataEXTM0202.PropTxtCom.Text 'コメント
                End If

                .Parameters("setSysDate").Value = dataEXTM0202.PropDtmSysDate 'システム日付
                .Parameters("setUpUserCd").Value = UpUserCd  'ユーザーCD
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try
    End Function

End Class
