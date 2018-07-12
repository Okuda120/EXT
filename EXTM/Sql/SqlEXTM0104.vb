Imports Common
Imports Npgsql

''' <summary>
''' 利用料マスタメンテナンス
''' </summary>
''' <remarks></remarks>
Public Class SqlEXTM0104

    Private strSelectBunruiSql As String = "SELECT " +
                                        "bunrui_cd," +
                                        "shisetu_kbn," +
                                        "bunrui_nm," +
                                        "sort," +
                                        "CASE sts " +
                                        "WHEN '0' THEN FALSE " +
                                        "WHEN '1' THEN TRUE " +
                                        "END AS sts, " +
                                        "kikan_from || kikan_to," +
                                        "SUBSTR(kikan_from,1,4)," +
                                        "SUBSTR(kikan_from,5,2)," +
                                        "SUBSTR(kikan_to,1,4)," +
                                        "SUBSTR(kikan_to,5,2) " +
                                    "FROM " +
                                        "rbunrui_mst "

    '利用料(料金)の取得(fromまで)
    Private strSelectRyokinSql As String = "SELECT " +
                                        "bunrui_cd, " +
                                        "ryokin_cd, " +
                                        "ryokin_nm, " +
                                        "ryokin_hour, " +
                                        "ryokin, " +
                                        "sort, " +
                                        "CASE sts " +
                                        "WHEN '0' THEN FALSE " +
                                        "WHEN '1' THEN TRUE " +
                                        "END AS sts " +
                                        " , shisetu_kbn " +
                                        " , '0' " +
                                    "FROM " +
                                        "riyoryo_mst "

    '倍率の取得(fromまで)
    Private strSelectBairituSql As String = "SELECT " +
                                        "bairitu_cd, " +
                                        "shisetu_kbn, " +
                                        "bairitu_nm, " +
                                        "bairitu, " +
                                        "sort," +
                                        "CASE sts " +
                                        "WHEN '0' THEN FALSE " +
                                        "WHEN '1' THEN TRUE " +
                                        "END AS sts " +
                                    "FROM " +
                                        "bairitu_mst "


    '対象期間の取得(fromまで)
    Private strSelectKikanSql As String = "SELECT distinct " +
                                        "kikan_from || kikan_to, " +
                                        "kikan_from ||'～'" +
                                            "|| kikan_to " +
                                    "FROM " +
                                        "rbunrui_mst "


    '利用料(分類)の登録
    Private strInsertBunruiSql As String = "INSERT INTO " +
                                        "rbunrui_mst " +
                                    "(" +
                                        "kikan_from, " +
                                        "kikan_to, " +
                                        "shisetu_kbn, " +
                                        "bunrui_cd, " +
                                        "bunrui_nm, " +
                                        "sort, " +
                                        "sts, " +
                                        "add_dt, " +
                                        "add_user_cd, " +
                                        "up_dt, " +
                                        "up_user_cd " +
                                    ")VALUES( " +
                                        ":kikan_from, " +
                                        ":kikan_to, " +
                                        ":shisetu_kbn, " +
                                        ":bunrui_cd, " +
                                        ":bunrui_nm, " +
                                        ":sort, " +
                                        ":sts, " +
                                        ":add_dt, " +
                                        ":add_user_cd, " +
                                        ":up_dt, " +
                                        ":up_user_cd " +
                                        ")"

    '利用料(分類)の更新
    Private strUpdataBunruiSql As String = "UPDATE " +
                                        "RBUNRUI_MST " +
                                    "SET " +
                                        "kikan_from=:kikan_from, " +
                                        "kikan_to=:kikan_to, " +
                                        "bunrui_cd=:bunrui_cd, " +
                                        "bunrui_nm=:bunrui_nm, " +
                                        "sort=:sort, " +
                                        "sts=:sts, " +
                                        "up_dt=:up_dt, " +
                                        "up_user_cd=:up_user_cd," +
                                        "shisetu_kbn=:shisetu "


    '利用料(料金)の登録
    Private strInsertRyokinSql As String = "INSERT INTO " +
                                        "RIYORYO_MST " +
                                    "( " +
                                        "kikan_from, " +
                                        "kikan_to, " +
                                        "shisetu_kbn, " +
                                        "bunrui_cd, " +
                                        "ryokin_cd, " +
                                        "ryokin_nm, " +
                                        "ryokin_hour, " +
                                        "ryokin, " +
                                        "sort, " +
                                        "sts, " +
                                        "add_dt, " +
                                        "add_user_cd, " +
                                        "up_dt, " +
                                        "up_user_cd " +
                                    ")VALUES( " +
                                        ":kikan_from, " +
                                        ":kikan_to, " +
                                        ":shisetu_kbn, " +
                                        ":bunrui_cd, " +
                                        ":ryokin_cd, " +
                                        ":ryokin_nm, " +
                                        ":ryokin_hour, " +
                                        ":ryokin, " +
                                        ":sort, " +
                                        ":sts, " +
                                        ":add_dt, " +
                                        ":add_user_cd, " +
                                        ":up_dt, " +
                                        ":up_user_cd " +
                                        ")"

    '利用料(料金)の更新
    Private strUpdataRyokinSql As String = "UPDATE " +
                                        "RIYORYO_MST " +
                                    "SET  " +
                                        "kikan_from=:kikan_from, " +
                                        "kikan_to=:kikan_to, " +
                                        "ryokin_nm=:ryokin_nm, " +
                                        "ryokin_hour=:ryokin_hour, " +
                                        "ryokin=:ryokin, " +
                                        "sort=:sort, " +
                                        "sts=:sts, " +
                                        "up_dt=:up_dt, " +
                                        "up_user_cd=:up_user_cd "


    '倍率の登録
    Private strInsertBairituSql As String = "INSERT INTO " +
                                        "BAIRITU_MST " +
                                    "(  " +
                                        "kikan_from, " +
                                        "kikan_to, " +
                                        "shisetu_kbn, " +
                                        "bairitu_cd, " +
                                        "bairitu_nm, " +
                                        "bairitu, " +
                                        "sort, " +
                                        "sts, " +
                                        "add_dt, " +
                                        "add_user_cd, " +
                                        "up_dt, " +
                                        "up_user_cd " +
                                    ")VALUES( " +
                                        ":kikan_from, " +
                                        ":kikan_to, " +
                                        ":shisetu_kbn, " +
                                        ":bairitu_cd, " +
                                        ":bairitu_nm, " +
                                        ":bairitu, " +
                                        ":sort, " +
                                        ":sts, " +
                                        ":add_dt, " +
                                        ":add_user_cd, " +
                                        ":up_dt, " +
                                        ":up_user_cd " +
                                    ")"

    '倍率の更新
    Private strUpdataBairituSql As String = "UPDATE " +
                                        "BAIRITU_MST " +
                                    "SET " +
                                        "kikan_from=:kikan_from, " +
                                        "kikan_to=:kikan_to, " +
                                        "bairitu_nm=:bairitu_nm, " +
                                        "bairitu=:bairitu, " +
                                        "sort=:sort, " +
                                        "sts=:sts, " +
                                        "up_dt=:up_dt, " +
                                        "up_user_cd=:up_user_cd," +
                                        "shisetu_kbn=:shisetu "

    ' サーバ日時取得
    Private EX26S000 As String = "SELECT Now() AS Sysdate "

    'SQLWHERE文の先頭(各メソッドで初期化する必要がある)
    Private sqlwhere As String


    ''' <summary>
    ''' 分類を取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>分類取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetBunruiData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        'sqlwhere 初期化
        sqlwhere = "WHERE "

        'sqlwhere値セット
        sqlwhere += "kikan_from || kikan_to =:kikan "

        Dim orderby As String = "order by sort"
        'SQL分
        strSQL = strSelectBunruiSql + sqlwhere + orderby
        Try

            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))    '期間
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("kikan").Value = dataEXTM0104.PropStrKikan                             '期間


            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金を取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>料金取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetRyokinData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String
        sqlwhere = "WHERE "

        'sqlwhere値セット
        sqlwhere += "kikan_from || kikan_to = :kikan "
        '20151023 sqlwhere += "AND bunrui_cd = :bunrui "
        '20151023 sqlwhere += "AND shisetu_kbn=:shisetu "
        Dim ord As String = "order by bunrui_cd,sort"

        'SQL
        strSQL = strSelectRyokinSql + sqlwhere + ord

        Try
            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))    '期間
            '20151023 Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("bunrui", NpgsqlTypes.NpgsqlDbType.Varchar))   '分類CD
            '20151023 Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("shisetu", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設

            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("kikan").Value = dataEXTM0104.PropStrKikan         '期間
            '20151023 Adapter.SelectCommand.Parameters("bunrui").Value = dataEXTM0104.PropStrBunruicd     '分類CD
            '20151023 Adapter.SelectCommand.Parameters("shisetu").Value = dataEXTM0104.PropStrShisetu    '施設

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 倍率を取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>倍率取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetBairituData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strSelectBairituSql
        sqlwhere = "WHERE "
        sqlwhere += "kikan_from || kikan_to =:kikan "

        'SQL文
        strSQL = strSQL + sqlwhere + "order by sort"

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("kikan").Value = dataEXTM0104.PropStrKikan

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 期間を取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>期間取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetKikanData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数の宣言
        Dim strSQL As String = strSelectKikanSql
        'SQL
        strSQL += "ORDER BY " +
                  "kikan_from || kikan_to "
        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 分類登録用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>分類登録用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function InsertBunrui(ByRef cmd As NpgsqlCommand, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        'SQL
        Dim strSQL As String = strInsertBunruiSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            cmd.Parameters.Add(New NpgsqlParameter("shisetu_kbn", NpgsqlTypes.NpgsqlDbType.Varchar))    '施設区分
            cmd.Parameters.Add(New NpgsqlParameter("bunrui_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類CD
            cmd.Parameters.Add(New NpgsqlParameter("bunrui_nm", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類名
            cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            cmd.Parameters.Add(New NpgsqlParameter("add_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))       '登録日付
            cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            cmd.Parameters.Add(New NpgsqlParameter("add_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))    '登録ユーザCD
            cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '更新ユーザCD

            'バインド変数に値をセット
            cmd.Parameters("kikan_from").Value = dataEXTM0104.PropStrKikanFrom                          '期間FROM
            cmd.Parameters("kikan_to").Value = dataEXTM0104.PropStrKikanTo                              '期間TO
            cmd.Parameters("shisetu_kbn").Value = dataEXTM0104.PropStrShisetu                           '施設区分
            cmd.Parameters("bunrui_cd").Value = dataEXTM0104.PropStrBunruicdInsert                      '分類CD
            cmd.Parameters("bunrui_nm").Value = dataEXTM0104.PropStrBunruinm                            '分類名
            cmd.Parameters("sort").Value = CInt(dataEXTM0104.PropStrSort)                      '並び順
            cmd.Parameters("sts").Value = dataEXTM0104.PropStrSts                                       'ステータス
            'cmd.Parameters("add_dt").Value = Now                                                        '登録日付
            'cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'cmd.Parameters("add_user_cd").Value = "super"                                               '登録ユーザCD
            'cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            cmd.Parameters("add_dt").Value = dataEXTM0104.PropDtSysDate                                 '登録日付
            cmd.Parameters("up_dt").Value = dataEXTM0104.PropDtSysDate                                  '更新日付
            cmd.Parameters("add_user_cd").Value = CommonEXT.PropComStrUserId                            '登録ユーザCD
            cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 分類更新用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>分類更新用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function UpdataBunrui(ByRef Cmd As NpgsqlCommand, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataextm0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数の宣言
        Dim sqlstr As String = strUpdataBunruiSql
        sqlwhere = "WHERE "
        sqlwhere += "kikan_from || kikan_to =:kikan "
        sqlwhere += "AND bunrui_cd=:bunrui_cd "
        sqlwhere += "AND shisetu_kbn=:shisetu "
        'SQL
        sqlstr += sqlwhere

        Try
            'データアダプタに、SQLのINSERT文を設定
            Cmd = New NpgsqlCommand(sqlstr, Cn)
            'バインド変数に型をセット
            Cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            Cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            Cmd.Parameters.Add(New NpgsqlParameter("bunrui_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類CD
            Cmd.Parameters.Add(New NpgsqlParameter("bunrui_nm", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類名
            Cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            Cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            Cmd.Parameters.Add(New NpgsqlParameter("shisetu", NpgsqlTypes.NpgsqlDbType.Varchar))        '施設区分
            Cmd.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))          '期間
            Cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            Cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '登録日付

            'バインド変数に値をセット
            Cmd.Parameters("kikan_from").Value = dataextm0104.PropStrKikanFrom                          '期間FROM
            Cmd.Parameters("kikan_to").Value = dataextm0104.PropStrKikanTo                              '期間TO
            Cmd.Parameters("bunrui_cd").Value = dataextm0104.PropStrBunruicdInsert                      '分類CD
            Cmd.Parameters("bunrui_nm").Value = dataextm0104.PropStrBunruinm                            '分類名
            Cmd.Parameters("sort").Value = CInt(dataextm0104.PropStrSort)                               '並び順
            Cmd.Parameters("sts").Value = dataextm0104.PropStrSts                                       'ステータス
            Cmd.Parameters("shisetu").Value = dataextm0104.PropStrShisetu                               '施設区分
            Cmd.Parameters("kikan").Value = dataextm0104.PropStrKikan                                   '期間
            'Cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'Cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            Cmd.Parameters("up_dt").Value = dataextm0104.PropDtSysDate                                  '更新日付
            Cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金登録用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>料金登録用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function InsertRyokin(ByRef cmd As NpgsqlCommand, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strInsertRyokinSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            cmd.Parameters.Add(New NpgsqlParameter("shisetu_kbn", NpgsqlTypes.NpgsqlDbType.Varchar))    '施設区分
            cmd.Parameters.Add(New NpgsqlParameter("bunrui_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類CD
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '料金CD
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_nm", NpgsqlTypes.NpgsqlDbType.Varchar))      '料金名
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_hour", NpgsqlTypes.NpgsqlDbType.Integer))    '貸し時間
            cmd.Parameters.Add(New NpgsqlParameter("ryokin", NpgsqlTypes.NpgsqlDbType.Integer))         '料金
            cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            cmd.Parameters.Add(New NpgsqlParameter("add_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))       '登録日付
            cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            cmd.Parameters.Add(New NpgsqlParameter("add_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))    '登録ユーザCD
            cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '更新ユーザCD

            'バインド変数に型をセット
            cmd.Parameters("kikan_from").Value = dataEXTM0104.PropStrKikanFrom                          '期間FROM
            cmd.Parameters("kikan_to").Value = dataEXTM0104.PropStrKikanTo                              '期間TO
            cmd.Parameters("shisetu_kbn").Value = dataEXTM0104.PropStrShisetu                           '施設区分
            cmd.Parameters("bunrui_cd").Value = dataEXTM0104.PropStrBunruicdInsert                      '分類CD
            cmd.Parameters("ryokin_cd").Value = dataEXTM0104.PropStrRyokincd                            '料金CD
            cmd.Parameters("ryokin_nm").Value = dataEXTM0104.PropStrRyokinnm                            '料金名
            cmd.Parameters("ryokin_hour").Value = CInt(dataEXTM0104.PropStrRyokinhour)         '貸し時間
            cmd.Parameters("ryokin").Value = CInt(dataEXTM0104.PropStrRyokin)                  '料金
            cmd.Parameters("sort").Value = CInt(dataEXTM0104.PropStrSort)                      '並び順
            cmd.Parameters("sts").Value = dataEXTM0104.PropStrSts                                       'ステータス
            'cmd.Parameters("add_dt").Value = Now                                                        '登録日付
            'cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'cmd.Parameters("add_user_cd").Value = "super"                                               '登録ユーザCD
            'cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            cmd.Parameters("add_dt").Value = dataEXTM0104.PropDtSysDate                                 '登録日付
            cmd.Parameters("up_dt").Value = dataEXTM0104.PropDtSysDate                                  '更新日付
            cmd.Parameters("add_user_cd").Value = CommonEXT.PropComStrUserId                            '登録ユーザCD
            cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD  

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金更新用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>料金更新用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function UpdataRyokin(ByRef cmd As NpgsqlCommand, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strUpdataRyokinSql

        sqlwhere = "WHERE "
        sqlwhere += "kikan_from || kikan_to = :kikan "
        sqlwhere += "AND bunrui_cd = :bunrui_cd "
        sqlwhere += "AND shisetu_kbn = :shisetu "
        sqlwhere += "AND ryokin_cd = :ryokin_cd "
        'SQL
        strSQL += sqlwhere

        Try
            'データアダプタに、SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型をセット
            cmd.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))          '期間
            cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            cmd.Parameters.Add(New NpgsqlParameter("bunrui_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '分類CD
            cmd.Parameters.Add(New NpgsqlParameter("shisetu", NpgsqlTypes.NpgsqlDbType.Varchar))        '施設区分
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_cd", NpgsqlTypes.NpgsqlDbType.Varchar))      '料金CD
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_nm", NpgsqlTypes.NpgsqlDbType.Varchar))      '料金名
            cmd.Parameters.Add(New NpgsqlParameter("ryokin_hour", NpgsqlTypes.NpgsqlDbType.Integer))    '貸し時間
            cmd.Parameters.Add(New NpgsqlParameter("ryokin", NpgsqlTypes.NpgsqlDbType.Integer))         '料金
            cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '更新ユーザCD

            'バインド変数に値をセット
            cmd.Parameters("kikan").Value = dataEXTM0104.PropStrKikan                                   '期間
            cmd.Parameters("kikan_from").Value = dataEXTM0104.PropStrKikanFrom                          '期間FROM
            cmd.Parameters("kikan_to").Value = dataEXTM0104.PropStrKikanTo                              '期間TO
            cmd.Parameters("bunrui_cd").Value = dataEXTM0104.PropStrBunruicdInsert                      '分類CD
            cmd.Parameters("shisetu").Value = dataEXTM0104.PropStrShisetu                               '施設区分
            cmd.Parameters("ryokin_cd").Value = dataEXTM0104.PropStrRyokincd                            '料金CD
            cmd.Parameters("ryokin_nm").Value = dataEXTM0104.PropStrRyokinnm                            '料金名
            cmd.Parameters("ryokin_hour").Value = CInt(dataEXTM0104.PropStrRyokinhour)         '貸し時間
            cmd.Parameters("ryokin").Value = CInt(dataEXTM0104.PropStrRyokin)                  '料金
            cmd.Parameters("sort").Value = CInt(dataEXTM0104.PropStrSort)                      '並び順
            cmd.Parameters("sts").Value = dataEXTM0104.PropStrSts                                       'ステータス
            'cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            cmd.Parameters("up_dt").Value = dataEXTM0104.PropDtSysDate                                  '更新日付
            cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 倍率登録用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>倍率登録用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function InsertBairitu(ByRef cmd As NpgsqlCommand, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数の宣言
        Dim strSQL As String = strInsertBairituSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            cmd.Parameters.Add(New NpgsqlParameter("shisetu_kbn", NpgsqlTypes.NpgsqlDbType.Varchar))    '施設区分
            cmd.Parameters.Add(New NpgsqlParameter("bairitu_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '倍率CD
            cmd.Parameters.Add(New NpgsqlParameter("bairitu_nm", NpgsqlTypes.NpgsqlDbType.Varchar))     '倍率名
            cmd.Parameters.Add(New NpgsqlParameter("bairitu", NpgsqlTypes.NpgsqlDbType.Double))         '倍率
            cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            cmd.Parameters.Add(New NpgsqlParameter("add_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))       '登録日付
            cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            cmd.Parameters.Add(New NpgsqlParameter("add_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))    '登録ユーザCD
            cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '更新ユーザCD

            'バインド変数に値をセット
            cmd.Parameters("kikan_from").Value = dataEXTM0104.PropStrKikanFrom                          '期間FROM
            cmd.Parameters("kikan_to").Value = dataEXTM0104.PropStrKikanTo                              '期間TO
            cmd.Parameters("shisetu_kbn").Value = dataEXTM0104.PropStrShisetu                           '施設区分
            cmd.Parameters("bairitu_cd").Value = dataEXTM0104.PropStrBairitucd                          '倍率CD
            cmd.Parameters("bairitu_nm").Value = dataEXTM0104.PropStrbairitunm                          '倍率名
            cmd.Parameters("bairitu").Value = Double.Parse(dataEXTM0104.propStrBairitu)                       '倍率
            cmd.Parameters("sort").Value = CInt(dataEXTM0104.PropStrSort)                      '並び順
            cmd.Parameters("sts").Value = dataEXTM0104.PropStrSts                                       'ステータス
            'cmd.Parameters("add_dt").Value = Now                                                        '登録日付
            'cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'cmd.Parameters("add_user_cd").Value = "super"                                               '登録ユーザCD
            'cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            cmd.Parameters("add_dt").Value = dataEXTM0104.PropDtSysDate                                 '登録日付
            cmd.Parameters("up_dt").Value = dataEXTM0104.PropDtSysDate                                  '更新日付
            cmd.Parameters("add_user_cd").Value = CommonEXT.PropComStrUserId                            '登録ユーザCD
            cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 倍率更新用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>倍率更新用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function UpdataBairitu(ByRef cmd As NpgsqlCommand, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数の宣言
        Dim strSQL As String = strUpdataBairituSql

        sqlwhere = "WHERE "
        sqlwhere += "kikan_from || kikan_to =:kikan "
        sqlwhere += "AND shisetu_kbn=:shisetu "
        sqlwhere += "AND bairitu_cd=:bairitu_cd "
        'SQL
        strSQL += sqlwhere
        Try
            'データアダプタに、SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            'バインド変数に型をセット
            cmd.Parameters.Add(New NpgsqlParameter("kikan_from", NpgsqlTypes.NpgsqlDbType.Varchar))     '期間FROM
            cmd.Parameters.Add(New NpgsqlParameter("kikan_to", NpgsqlTypes.NpgsqlDbType.Varchar))       '期間TO
            cmd.Parameters.Add(New NpgsqlParameter("bairitu_nm", NpgsqlTypes.NpgsqlDbType.Varchar))     '倍率名
            cmd.Parameters.Add(New NpgsqlParameter("bairitu", NpgsqlTypes.NpgsqlDbType.Double))         '倍率
            cmd.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            cmd.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            'ステータス
            cmd.Parameters.Add(New NpgsqlParameter("up_dt", NpgsqlTypes.NpgsqlDbType.Timestamp))        '更新日付
            cmd.Parameters.Add(New NpgsqlParameter("up_user_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '更新ユーザCD
            cmd.Parameters.Add(New NpgsqlParameter("kikan", NpgsqlTypes.NpgsqlDbType.Varchar))          '期間
            cmd.Parameters.Add(New NpgsqlParameter("shisetu", NpgsqlTypes.NpgsqlDbType.Varchar))        '施設区分
            cmd.Parameters.Add(New NpgsqlParameter("bairitu_cd", NpgsqlTypes.NpgsqlDbType.Varchar))     '倍率CD

            'バインド変数に型をセット
            cmd.Parameters("kikan_from").Value = dataEXTM0104.PropStrKikanFrom                          '期間FROM
            cmd.Parameters("kikan_to").Value = dataEXTM0104.PropStrKikanTo                              '期間TO
            cmd.Parameters("bairitu_nm").Value = dataEXTM0104.PropStrbairitunm                          '倍率名
            cmd.Parameters("bairitu").Value = Double.Parse(dataEXTM0104.propStrBairitu)                       '倍率
            cmd.Parameters("sort").Value = CInt(dataEXTM0104.PropStrSort)                      '並び順
            cmd.Parameters("sts").Value = dataEXTM0104.PropStrSts                                       'ステータス
            'cmd.Parameters("up_dt").Value = Now                                                         '更新日付
            'cmd.Parameters("up_user_cd").Value = "super"                                                '更新ユーザCD
            cmd.Parameters("up_dt").Value = dataEXTM0104.PropDtSysDate                                  '更新日付
            cmd.Parameters("up_user_cd").Value = CommonEXT.PropComStrUserId                             '更新ユーザCD
            cmd.Parameters("kikan").Value = dataEXTM0104.PropStrKikan                                   '期間
            cmd.Parameters("shisetu").Value = dataEXTM0104.PropStrShisetu                               '施設区分
            cmd.Parameters("bairitu_cd").Value = dataEXTM0104.PropStrBairitucd                          '倍率CD

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' サーバ日時取得
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="DataEXTM0104">[IN]利用料マスタメンテ画面データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>サーバ日時取得のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.10.16 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSysDate(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByVal Cn As NpgsqlConnection, _
                                       ByVal dataEXTM0104 As DataEXTM0104) As Boolean

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
            puErrMsg = E0000 & ex.Message
            Return False
        End Try

    End Function

End Class
