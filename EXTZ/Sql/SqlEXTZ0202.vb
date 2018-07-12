Imports Common
Imports CommonEXT
Imports Npgsql

Public Class SqlEXTZ0202
    'SQL文(ログイン)
    Private strEX99S001 As String = _
                           "select " & vbCrLf & _
                           "    sum(t1.c1) as cnt1 " & vbCrLf & _
                           "from " & vbCrLf & _
                           "    ( " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            count(*) as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            YDT_TBL t1, YOYAKU_TBL t2 " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            t1.yoyaku_dt = :YoyakuDt " & vbCrLf & _
                           "        AND t1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "        AND t1.studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "        AND (t1.yoyaku_no <> :YoyakuNo OR :YoyakuNo IS NULL) " & vbCrLf & _
                           "        AND t1.yoyaku_no = t2.yoyaku_no " & vbCrLf & _
                           "        AND t2.yoyaku_sts in ('1','2','3','4') " & vbCrLf & _
                           "        union all " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            count(*) as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            yctl_tbl " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            yoyaku_dt = :YoyakuDt " & vbCrLf & _
                           "        AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "        AND studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "        AND user_cd <> :UserId " & vbCrLf & _
                           "        union all " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            3 as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            YDT_TBL t1, YOYAKU_TBL t2 " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            t1.yoyaku_dt = :YoyakuDt " & vbCrLf & _
                           "        AND '2' = :ShisetuKbn " & vbCrLf & _
                           "        AND t1.studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "        AND t1.riyo_keitai = '2' " & vbCrLf & _
                           "        AND (t1.yoyaku_no <> :YoyakuNo OR :YoyakuNo IS NULL) " & vbCrLf & _
                           "        AND t1.yoyaku_no = t2.yoyaku_no " & vbCrLf & _
                           "        AND t2.yoyaku_sts in ('1','2','3','4') " & vbCrLf & _
                           "        union all " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            3 as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            HOLMENT_MST " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            holment_dt = :YoyakuDt " & vbCrLf & _
                           "        AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "        AND studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "    ) t1 "

    'SQL文(ログイン)
    ' 2016.02.15 UPD START↓ h.hagiwara キャンセル待ちの｢Lockout｣を３枠扱いしない様（１枠扱い）に修正
    'Private strEX99S002 As String = _
    '                       "select " & vbCrLf & _
    '                       "    sum(t1.c1) as cnt1 " & vbCrLf & _
    '                       "from " & vbCrLf & _
    '                       "    ( " & vbCrLf & _
    '                       "        select " & vbCrLf & _
    '                       "            count(*) as c1 " & vbCrLf & _
    '                       "        from " & vbCrLf & _
    '                       "            CANCEL_WAIT_DT_TBL t1, CANCEL_WAIT_TBL t2 " & vbCrLf & _
    '                       "        where " & vbCrLf & _
    '                       "            t1.riyo_dt = :CancelDt " & vbCrLf & _
    '                       "        AND t1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
    '                       "        AND t1.studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
    '                       "        AND (t1.cancel_wait_no <> :CancelWaitNo OR :CancelWaitNo IS NULL) " & vbCrLf & _
    '                       "        AND t1.cancel_wait_no = t2.cancel_wait_no " & vbCrLf & _
    '                       "        AND t2.cancel_wait_sts in ('1') " & vbCrLf & _
    '                       "        union all " & vbCrLf & _
    '                       "        select " & vbCrLf & _
    '                       "            count(*) as c1 " & vbCrLf & _
    '                       "        from " & vbCrLf & _
    '                       "            cctl_tbl " & vbCrLf & _
    '                       "        where " & vbCrLf & _
    '                       "            yoyaku_dt = :CancelDt " & vbCrLf & _
    '                       "        AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
    '                       "        AND studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
    '                       "        AND user_cd <> :UserId " & vbCrLf & _
    '                       "        union all " & vbCrLf & _
    '                       "        select " & vbCrLf & _
    '                       "            3 as c1 " & vbCrLf & _
    '                       "        from " & vbCrLf & _
    '                       "            CANCEL_WAIT_DT_TBL t1, CANCEL_WAIT_TBL t2 " & vbCrLf & _
    '                       "        where " & vbCrLf & _
    '                       "            t1.riyo_dt = :CancelDt " & vbCrLf & _
    '                       "        AND '2' = :ShisetuKbn " & vbCrLf & _
    '                       "        AND t1.studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
    '                       "        AND t1.riyo_keitai = '2' " & vbCrLf & _
    '                       "        AND (t1.cancel_wait_no <> :CancelWaitNo OR :CancelWaitNo IS NULL) " & vbCrLf & _
    '                       "        AND t1.cancel_wait_no = t2.cancel_wait_no " & vbCrLf & _
    '                       "        AND t2.cancel_wait_sts in ('1') " & vbCrLf & _
    '                       "    ) t1 "
    Private strEX99S002 As String = _
                           "select " & vbCrLf & _
                           "    sum(t1.c1) as cnt1 " & vbCrLf & _
                           "from " & vbCrLf & _
                           "    ( " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            count(*) as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            CANCEL_WAIT_DT_TBL t1, CANCEL_WAIT_TBL t2 " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            t1.riyo_dt = :CancelDt " & vbCrLf & _
                           "        AND t1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "        AND t1.studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "        AND (t1.cancel_wait_no <> :CancelWaitNo OR :CancelWaitNo IS NULL) " & vbCrLf & _
                           "        AND t1.cancel_wait_no = t2.cancel_wait_no " & vbCrLf & _
                           "        AND t2.cancel_wait_sts in ('1') " & vbCrLf & _
                           "        union all " & vbCrLf & _
                           "        select " & vbCrLf & _
                           "            count(*) as c1 " & vbCrLf & _
                           "        from " & vbCrLf & _
                           "            cctl_tbl " & vbCrLf & _
                           "        where " & vbCrLf & _
                           "            yoyaku_dt = :CancelDt " & vbCrLf & _
                           "        AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "        AND studio_kbn in (CASE WHEN :StudioKbn='3' THEN '1' ELSE :StudioKbn END, CASE WHEN :StudioKbn='3' THEN '2' ELSE :StudioKbn END, '3') " & vbCrLf & _
                           "        AND user_cd <> :UserId " & vbCrLf & _
                           "    ) t1 "
    ' 2016.02.15 UPD END↑ h.hagiwara キャンセル待ちの｢Lockout｣を３枠扱いしない様（１枠扱い）に修正

    'SQL文(予約受付制御データ削除)
    Private strEX01D001 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "YCTL_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "USER_CD = :UserId "

    'SQL文(キャンセル制御データ削除)
    Private strEX01D002 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "CCTL_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "USER_CD = :UserId "
    'SQL(予約受付制御登録)
    Private strEX04I002 As String = _
                           "insert into YCTL_TBL(user_cd, shisetu_kbn, studio_kbn, yoyaku_dt, add_dt, add_user_cd)" & vbCrLf & _
                           "values( " & vbCrLf & _
                               ":UserId, " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":YoyakuDt, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd " & vbCrLf & _
                           ") "

    'SQL(キャンセル受付制御登録)
    Private strEX04I004 As String = _
                           "insert into CCTL_TBL(user_cd, shisetu_kbn, studio_kbn, yoyaku_dt, add_dt, add_user_cd)" & vbCrLf & _
                           "values( " & vbCrLf & _
                               ":UserId, " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":YoyakuDt, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd " & vbCrLf & _
                           ") "

    ''' <summary>
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function cntRiyobi(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataRiyobi As CommonDataRiyobi) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX99S001

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(0).Value = dataRiyobi.PropStrYoyakuNo
                .Parameters(1).Value = dataRiyobi.PropStrYoyakuDt
                .Parameters(2).Value = dataRiyobi.PropStrShisetuKbn
                .Parameters(3).Value = dataRiyobi.PropStrStudioKbn
                .Parameters(4).Value = CommonEXT.PropComStrUserId
            End With

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
    ''' 予約受付制御テーブルのデータ削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteYoyakuCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal StrShisetuKbn As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX01D001
            If StrShisetuKbn = CommonDeclareEXT.SHISETU_KBN_THEATER Or StrShisetuKbn = CommonDeclareEXT.SHISETU_KBN_STUDIO Then
                strSQL = strSQL & " AND SHISETU_KBN = :ShisetuKbn " ' 条件追加
                Cmd = New NpgsqlCommand(strSQL, Cn)

                '値を設定
                Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
                Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
                Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
                Cmd.Parameters(1).Value = StrShisetuKbn
            Else
                Cmd = New NpgsqlCommand(strSQL, Cn)

                '値を設定
                Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
                Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            End If

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' キャンセル待ち登録制御テーブルのデータ削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteCancelCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX01D002
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId              'ユーザーID

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 予約受付制御テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertYoyakuCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As CommonDataRiyobi) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX04I002
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(1).Value = dataRiyobi.PropStrShisetuKbn
            Cmd.Parameters(2).Value = dataRiyobi.PropStrStudioKbn
            Cmd.Parameters(3).Value = dataRiyobi.PropStrYoyakuDt
            Cmd.Parameters(4).Value = CommonEXT.PropComStrUserId

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 予約受付制御テーブルの行削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="StrUserId"></param>
    ''' <param name="StrYoyakuDt"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteLineYoyakuCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal StrUserId As String, _
                                   ByVal StrYoyakuDt As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX01D001
            strSQL = strSQL & " AND YOYAKU_DT = :YoyakuDt " ' 条件追加
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = StrUserId              'ユーザーID
            Cmd.Parameters(1).Value = StrYoyakuDt            '日付

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' キャンセル待ち日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function cntCancelWait(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataRiyobi As CommonDataCancel) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX99S002

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":CancelDt", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(0).Value = dataRiyobi.PropStrCancelWaitNo
                .Parameters(1).Value = dataRiyobi.PropStrCancelDt
                .Parameters(2).Value = dataRiyobi.PropStrShisetuKbn
                .Parameters(3).Value = dataRiyobi.PropStrStudioKbn
                .Parameters(4).Value = CommonEXT.PropComStrUserId
            End With

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
    ''' キャンセル待ち受付制御テーブルのデータ削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteCancelCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal StrShisetuKbn As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX01D002
            If StrShisetuKbn = CommonDeclareEXT.SHISETU_KBN_THEATER Or StrShisetuKbn = CommonDeclareEXT.SHISETU_KBN_STUDIO Then
                strSQL = strSQL & " AND SHISETU_KBN = :ShisetuKbn " ' 条件追加
                Cmd = New NpgsqlCommand(strSQL, Cn)

                '値を設定
                Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
                Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
                Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
                Cmd.Parameters(1).Value = StrShisetuKbn
            Else
                Cmd = New NpgsqlCommand(strSQL, Cn)

                '値を設定
                Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
                Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            End If

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertCancelCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As CommonDataCancel) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX04I004
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(1).Value = dataRiyobi.PropStrShisetuKbn
            Cmd.Parameters(2).Value = dataRiyobi.PropStrStudioKbn
            Cmd.Parameters(3).Value = dataRiyobi.PropStrCancelDt
            Cmd.Parameters(4).Value = CommonEXT.PropComStrUserId

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' キャンセル待ち受付制御テーブルの行削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="StrUserId"></param>
    ''' <param name="StrYoyakuDt"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteLineCancelCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal StrUserId As String, _
                                   ByVal StrYoyakuDt As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX01D002
            strSQL = strSQL & " AND YOYAKU_DT = :YoyakuDt " ' 条件追加
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = StrUserId              'ユーザーID
            Cmd.Parameters(1).Value = StrYoyakuDt            '日付

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

End Class
