Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' その他利用登録（シアター）で使用する情報取得・更新SQL作成
''' </summary>
''' <remarks>その他利用登録（シアター）で使用する情報取得・更新のSQLを作成する
''' <para>作成情報：2015/08/26 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTB0105

    '*****基本SQL文*****
#Region "SQL文(その他利用情報取得)"

    'SQL文(その他利用情報取得)
    Private strSelectSonotaInfoSQL As String = _
    "select " & vbCrLf & _
    "   RIYO_DT" & vbCrLf & _
    "  ,START_TIME" & vbCrLf & _
    "  ,END_TIME" & vbCrLf & _
    "  ,RIYO_BASHO" & vbCrLf & _
    "  ,RIYO_YOTO" & vbCrLf & _
    "  ,RIYOSHA" & vbCrLf & _
    "from" & vbCrLf & _
    "  SONOTA_DT_TBL " & vbCrLf & _
    "where" & vbCrLf & _
    "  RIYO_DT = :RIYO_DT"

#End Region

#Region "SQL文(その他利用情報更新)"

    'SQL文(その他利用情報更新)
    Private strUpdateSonotaInfoSQL As String = _
    "update SONOTA_DT_TBL" & vbCrLf & _
    "   set" & vbCrLf & _
    "   RIYO_DT = :RIYO_DT" & vbCrLf & _
    "  ,START_TIME = :START_TIME" & vbCrLf & _
    "  ,END_TIME = :END_TIME " & vbCrLf & _
    "  ,RIYO_BASHO = :RIYO_BASHO" & vbCrLf & _
    "  ,RIYO_YOTO = :RIYO_YOTO" & vbCrLf & _
    "  ,RIYOSHA = :RIYOSHA" & vbCrLf & _
    "  ,ADD_DT = CURRENT_TIMESTAMP" & vbCrLf & _
    "  ,ADD_USER_CD = :USER_CD" & vbCrLf & _
    "  ,UP_DT = CURRENT_TIMESTAMP" & vbCrLf & _
    "  ,UP_USER_CD = :USER_CD" & vbCrLf & _
    "where" & vbCrLf & _
    "  RIYO_DT = :RIYO_DT"

#End Region

#Region "SQL文(その他利用情報削除)"

    'SQL文(その他利用情報削除)
    Private strDeleteSonotaInfoSQL As String = _
    "delete" & vbCrLf & _
    "from" & vbCrLf & _
    "  SONOTA_DT_TBL " & vbCrLf & _
    "where" & vbCrLf & _
    "  RIYO_DT = :RIYO_DT"

#End Region

#Region "SQL文(その他利用情報登録)"

    'SQL文(その他利用情報登録)
    Private strInsertSonotaInfoSQL As String = _
    "insert into SONOTA_DT_TBL values(" & vbCrLf & _
    "    :RIYO_DT" & vbCrLf & _
    "   ,:START_TIME" & vbCrLf & _
    "   ,:END_TIME " & vbCrLf & _
    "   ,:RIYO_BASHO" & vbCrLf & _
    "   ,:RIYO_YOTO" & vbCrLf & _
    "   ,:RIYOSHA" & vbCrLf & _
    "   ,CURRENT_TIMESTAMP" & vbCrLf & _
    "   ,:USER_CD" & vbCrLf & _
    "   ,CURRENT_TIMESTAMP" & vbCrLf & _
    "   ,:USER_CD" & vbCrLf & _
    ")"

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "その他利用情報取得"

    ''' <summary>
    ''' その他利用情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0105"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>その他利用情報を取得するSQLの作成
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectSonotaInfo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0105 As DataEXTB0105) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strSelectSonotaInfoSQL

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":RIYO_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":RIYO_DT").Value = dataEXTB0105.PropStrRiyoDt      '利用日
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

#End Region

#Region "その他利用情報更新"

    ''' <summary>
    ''' その他利用情報更新
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0105">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>その他利用情報更新を行うSQLの作成
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setUpdateSonotaInfo(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0105 As DataEXTB0105) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strUpdateSonotaInfoSQL

            'SQL文(UPDATE)
            'SQLのUPDATE文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":START_TIME", NpgsqlTypes.NpgsqlDbType.Varchar))    '開始時間
            cmd.Parameters.Add(New NpgsqlParameter(":END_TIME", NpgsqlTypes.NpgsqlDbType.Varchar))      '終了時間
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_BASHO", NpgsqlTypes.NpgsqlDbType.Varchar))    '利用場所
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_YOTO", NpgsqlTypes.NpgsqlDbType.Varchar))     '利用用途
            cmd.Parameters.Add(New NpgsqlParameter(":RIYOSHA", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用者
            cmd.Parameters.Add(New NpgsqlParameter(":USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))       'ユーザーID
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_DT", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用日

            '値を設定
            cmd.Parameters(":START_TIME").Value = dataEXTB0105.PropStrStartTime     '開始時間
            cmd.Parameters(":END_TIME").Value = dataEXTB0105.PropStrEndTime         '終了時間
            cmd.Parameters(":RIYO_BASHO").Value = dataEXTB0105.PropStrRiyoBasho     '利用場所
            cmd.Parameters(":RIYO_YOTO").Value = dataEXTB0105.PropStrRiyoYoto       '利用用途
            cmd.Parameters(":RIYOSHA").Value = dataEXTB0105.PropStrRiyosha          '利用者
            cmd.Parameters(":USER_CD").Value = CommonEXT.PropComStrUserId           'ユーザーID
            cmd.Parameters(":RIYO_DT").Value = dataEXTB0105.PropStrRiyoDt           '利用日

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

#End Region

#Region "その他利用情報削除"

    ''' <summary>
    ''' その他利用情報削除
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0105">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>その他利用情報削除を行うSQLの作成
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setDeleteSonotaInfo(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0105 As DataEXTB0105) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strDeleteSonotaInfoSQL

            'SQL文(DELETE)
            'SQLのDELETE文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_DT", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用日

            '値を設定
            cmd.Parameters(":RIYO_DT").Value = dataEXTB0105.PropStrRiyoDt        '利用日

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

#End Region

#Region "その他利用情報登録"

    ''' <summary>
    ''' その他利用情報登録
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0105">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>その他利用情報登録を行うSQLの作成
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setInsertSonotaInfo(ByRef cmd As NpgsqlCommand, _
                                     ByRef Cn As NpgsqlConnection, _
                                     ByVal dataEXTB0105 As DataEXTB0105) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strInsertSonotaInfoSQL

            'SQL文(INSERT)
            'SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_DT", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用日
            cmd.Parameters.Add(New NpgsqlParameter(":START_TIME", NpgsqlTypes.NpgsqlDbType.Varchar))    '開始時間
            cmd.Parameters.Add(New NpgsqlParameter(":END_TIME", NpgsqlTypes.NpgsqlDbType.Varchar))      '終了時間
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_BASHO", NpgsqlTypes.NpgsqlDbType.Varchar))    '利用場所
            cmd.Parameters.Add(New NpgsqlParameter(":RIYO_YOTO", NpgsqlTypes.NpgsqlDbType.Varchar))     '利用用途
            cmd.Parameters.Add(New NpgsqlParameter(":RIYOSHA", NpgsqlTypes.NpgsqlDbType.Varchar))       '利用者
            cmd.Parameters.Add(New NpgsqlParameter(":USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))       'ユーザーID

            '値を設定
            cmd.Parameters(":RIYO_DT").Value = dataEXTB0105.PropStrRiyoDt               '利用日
            cmd.Parameters(":START_TIME").Value = dataEXTB0105.PropStrStartTime         '開始時間
            cmd.Parameters(":END_TIME").Value = dataEXTB0105.PropStrEndTime             '終了時間
            cmd.Parameters(":RIYO_BASHO").Value = dataEXTB0105.PropStrRiyoBasho         '利用場所
            cmd.Parameters(":RIYO_YOTO").Value = dataEXTB0105.PropStrRiyoYoto           '利用用途
            cmd.Parameters(":RIYOSHA").Value = dataEXTB0105.PropStrRiyosha              '利用者
            cmd.Parameters(":USER_CD").Value = CommonEXT.PropComStrUserId               'ユーザーID

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

#End Region

End Class
