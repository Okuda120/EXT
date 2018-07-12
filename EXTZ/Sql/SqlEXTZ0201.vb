Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' メンテ日登録で使用する情報取得・更新SQL作成
''' </summary>
''' <remarks>メンテ日登録で使用する情報取得・更新のSQLを作成する
''' <para>作成情報：2015/08/28 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTZ0201

    '*****基本SQL文*****
#Region "SQL文(休館メンテ情報取得)"

    'SQL文(休館メンテ情報取得)
    Private strSelectHolmentInfoSQL As String = _
    "select " & vbCrLf & _
    "   HOLMENT_DT" & vbCrLf & _
    "  ,MNAIYO" & vbCrLf & _
    "from" & vbCrLf & _
    "  HOLMENT_MST " & vbCrLf & _
    "where" & vbCrLf & _
    "  HOLMENT_DT = :HOLMENT_DT" & vbCrLf & _
    "  AND SHISETU_KBN = :SHISETU_KBN" & vbCrLf

#End Region

#Region "SQL文(休館メンテ情報情報更新)"

    'SQL文(休館メンテ情報更新)
    Private strUpdateHolmentInfoSQL As String = _
    "update HOLMENT_MST " & vbCrLf & _
    "  set" & vbCrLf & _
    "  MNAIYO = :MNAIYO" & vbCrLf & _
    "  ,UP_DT = CURRENT_TIMESTAMP" & vbCrLf & _
    "  ,UP_USER_CD = :USER_CD" & vbCrLf & _
    "where" & vbCrLf & _
    "  HOLMENT_DT = :TARGET_DT" & vbCrLf & _
    "  AND SHISETU_KBN = :SHISETU_KBN" & vbCrLf & _
    "  AND STUDIO_KBN = :STUDIO_KBN"

#End Region

#Region "SQL文(休館メンテ情報削除)"

    'SQL文(休館メンテ情報削除)
    Private strDeleteHolmentInfoSQL As String = _
    "delete" & vbCrLf & _
    "from" & vbCrLf & _
    "  HOLMENT_MST " & vbCrLf & _
    "where" & vbCrLf & _
    "  HOLMENT_DT = :HOLMENT_DT" & vbCrLf & _
    "  AND SHISETU_KBN = :SHISETU_KBN" & vbCrLf

#End Region

#Region "SQL文(休館メンテ情報登録)"

    'SQL文(休館メンテ情報登録)
    Private strInsertHolmentInfoSQL As String = _
    "insert into HOLMENT_MST values(" & vbCrLf & _
    "    :HOLMENT_DT" & vbCrLf & _
    "   ,:SHISETU_KBN" & vbCrLf & _
    "   ,:STUDIO_KBN " & vbCrLf & _
    "   ,:MNAIYO" & vbCrLf & _
    "   ,:HOLMENT_KBN" & vbCrLf & _
    "   ,CURRENT_TIMESTAMP" & vbCrLf & _
    "   ,:USER_CD" & vbCrLf & _
    "   ,CURRENT_TIMESTAMP" & vbCrLf & _
    "   ,:USER_CD" & vbCrLf & _
    ")"

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "休館メンテ情報取得"

    ''' <summary>
    ''' 休館メンテ情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0201"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館メンテ情報を取得するSQLの作成
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectHolmentInfo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0201 As DataEXTZ0201) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strSelectHolmentInfoSQL
            If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                strSQL &= "  AND STUDIO_KBN = :STUDIO_KBN"
            End If

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":HOLMENT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))      '対象日
                .Add(New NpgsqlParameter(":SHISETU_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))     '施設区分
                If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                    .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))      'スタジオ区分
                End If
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":HOLMENT_DT").Value = dataEXTZ0201.PropStrHolmentDt        '対象日
                .Parameters(":SHISETU_KBN").Value = dataEXTZ0201.PropStrShisetuKbn      '施設区分
                If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                    .Parameters(":STUDIO_KBN").Value = dataEXTZ0201.PropStrStudioKbn        'スタジオ区分
                End If
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

#Region "休館メンテ情報更新"

    ''' <summary>
    ''' 休館メンテ情報更新
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0201">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館メンテ情報更新を行うSQLの作成
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setUpdateHolmentInfo(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0201 As DataEXTZ0201) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strUpdateHolmentInfoSQL

            'SQL文(UPDATE)
            'SQLのUPDATE文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":MNAIYO", NpgsqlTypes.NpgsqlDbType.Varchar))        'メンテナンス内容
            cmd.Parameters.Add(New NpgsqlParameter(":USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))       'ユーザーID
            cmd.Parameters.Add(New NpgsqlParameter(":TARGET_DT", NpgsqlTypes.NpgsqlDbType.Varchar))     '対象日
            cmd.Parameters.Add(New NpgsqlParameter(":SHISETU_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))   '施設区分
            cmd.Parameters.Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))    'スタジオ区分

            '値を設定
            cmd.Parameters(0).Value = dataEXTZ0201.PropStrMnaiyo            'メンテナンス内容
            cmd.Parameters(1).Value = CommonEXT.PropComStrUserId            'ユーザーID
            cmd.Parameters(2).Value = dataEXTZ0201.PropStrHolmentDt         '対象日
            cmd.Parameters(3).Value = dataEXTZ0201.PropStrShisetuKbn        '施設区分
            cmd.Parameters(4).Value = dataEXTZ0201.PropStrStudioKbn         'スタジオ区分

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

#Region "休館メンテ情報削除"

    ''' <summary>
    ''' 休館メンテ情報削除
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0201">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館メンテ情報削除を行うSQLの作成
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setDeleteHolmentInfo(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0201 As DataEXTZ0201) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strDeleteHolmentInfoSQL
            If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                strSQL &= "  AND STUDIO_KBN = :STUDIO_KBN"
            End If

            'SQL文(DELETE)
            'SQLのDELETE文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":HOLMENT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))        '対象日
            cmd.Parameters.Add(New NpgsqlParameter(":SHISETU_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))       '施設区分
            If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                cmd.Parameters.Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))    'スタジオ区分
            End If

            '値を設定
            cmd.Parameters(0).Value = dataEXTZ0201.PropStrHolmentDt         '対象日
            cmd.Parameters(1).Value = dataEXTZ0201.PropStrShisetuKbn        '施設区分
            If Not dataEXTZ0201.PropStrStudioKbn Is Nothing Then
                cmd.Parameters(2).Value = dataEXTZ0201.PropStrStudioKbn     'スタジオ区分
            End If

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

#Region "休館メンテ情報登録"

    ''' <summary>
    ''' 休館メンテ情報登録
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0201">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館メンテ情報登録を行うSQLの作成
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setInsertHolmentInfo(ByRef cmd As NpgsqlCommand, _
                                     ByRef Cn As NpgsqlConnection, _
                                     ByVal dataEXTZ0201 As DataEXTZ0201) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strInsertHolmentInfoSQL

            'SQL文(INSERT)
            'SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":HOLMENT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))        '対象日
            cmd.Parameters.Add(New NpgsqlParameter(":SHISETU_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))       '施設区分
            cmd.Parameters.Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))        'スタジオ区分
            cmd.Parameters.Add(New NpgsqlParameter(":MNAIYO", NpgsqlTypes.NpgsqlDbType.Varchar))            'メンテナンス内容
            cmd.Parameters.Add(New NpgsqlParameter(":HOLMENT_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))       'メンテナンス区分
            cmd.Parameters.Add(New NpgsqlParameter(":USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))           'ユーザーID

            '値を設定
            cmd.Parameters(":HOLMENT_DT").Value = dataEXTZ0201.PropStrHolmentDt             '対象日
            cmd.Parameters(":SHISETU_KBN").Value = dataEXTZ0201.PropStrShisetuKbn           '施設区分
            cmd.Parameters(":STUDIO_KBN").Value = dataEXTZ0201.PropStrStudioKbn             'スタジオ区分
            cmd.Parameters(":MNAIYO").Value = dataEXTZ0201.PropStrMnaiyo                    'メンテナンス内容
            cmd.Parameters(":HOLMENT_KBN").Value = dataEXTZ0201.PropHolmentKbn              'メンテナンス区分
            cmd.Parameters(":USER_CD").Value = CommonEXT.PropComStrUserId                   'ユーザーID

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
