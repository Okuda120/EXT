Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' 予約カレンダー（スタジオ）で使用する情報取得・更新SQL作成
''' </summary>
''' <remarks>予約カレンダー（スタジオ）で使用する情報取得・更新のSQLを作成する
''' <para>作成情報：2015/09/01 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTC0101

    '*****基本SQL文*****
#Region "SQL文(祝祭日取得)"

    Private strSelectShukusaijitsuSQL As String = _
    "select " & vbCrLf & _
    "  HOLIDAY_DT" & vbCrLf & _
    "from" & vbCrLf & _
    "  HOLIDAY_MST " & vbCrLf & _
    "where" & vbCrLf & _
    "  substr(HOLIDAY_DT,1,7) = :HOLIDAY_DT"
#End Region

#Region "SQL文(予約一覧取得)"

    'SQL文(予約一覧取得)
    Private strSelectYoyakuListSQL_SELECT As String = _
    "select " & vbCrLf & _
    "   a.SHISETU_KBN" & vbCrLf & _
    "  ,a.STUDIO_KBN" & vbCrLf & _
    "  ,a.YOYAKU_DT" & vbCrLf & _
    "  ,a.START_TIME" & vbCrLf & _
    "  ,a.END_TIME" & vbCrLf & _
    "  ,a.YOYAKU_NO" & vbCrLf & _
    "  ,a.MITEI_FLG" & vbCrLf & _
    "  ,b.STUDIO_KBN  " & vbCrLf & _
    "  ,b.SAIJI_NM" & vbCrLf & _
    "  ,b.SHUTSUEN_NM" & vbCrLf & _
    "  ,b.YOYAKU_STS" & vbCrLf & _
    "  ,a.RIYO_KEITAI" & vbCrLf & _
    "  ,c.RIYO_NM" & vbCrLf & _
    "from" & vbCrLf & _
    "  YDT_TBL a" & vbCrLf & _
    "left join YOYAKU_TBL b on(" & vbCrLf & _
    "  a.YOYAKU_NO = b.YOYAKU_NO" & vbCrLf & _
    ")" & vbCrLf & _
    "left join RIYOSHA_MST c on(" & vbCrLf & _
    "  b.RIYOSHA_CD = c.RIYOSHA_CD" & vbCrLf & _
    ")" & vbCrLf
    Private strSelectYoyakuListSQL_WHERE As String = _
    "where" & vbCrLf & _
    "  a.SHISETU_KBN = '2'" & vbCrLf & _
    "  and substr(a.YOYAKU_DT,1,7) = :YOYAKU_DT" & vbCrLf & _
    "  and b.YOYAKU_STS IN ('2', '3', '4')" & vbCrLf
    Private strSelectYoyakuListSQL_ORDER As String = _
    "order by a.YOYAKU_DT, a.START_TIME"

#End Region

#Region "SQL文(予約未確定一覧取得)"

    'SQL文(予約未確定一覧取得)
    Private strSelectMikakuYoyakuListSQL_SELECT As String = _
    "select" & vbCrLf & _
    "   USER_CD" & vbCrLf & _
    "  ,SHISETU_KBN" & vbCrLf & _
    "  ,STUDIO_KBN" & vbCrLf & _
    "  ,YOYAKU_DT" & vbCrLf & _
    "from" & vbCrLf & _
    "  YCTL_TBL" & vbCrLf
    Private strSelectMikakuYoyakuListSQL_WHERE As String = _
    "where" & vbCrLf & _
    "  SHISETU_KBN = '2'" & vbCrLf & _
    "  and substr(YOYAKU_DT,1,7) = :YOYAKU_DT"

#End Region

#Region "SQL文(キャンセル一覧取得)"

    'SQL文(キャンセル一覧取得)
    Private strSelectCancelListSQL_SELECT As String = _
    "select" & vbCrLf & _
     "a.SHISETU_KBN" & vbCrLf & _
      ",a.STUDIO_KBN" & vbCrLf & _
      ",a.RIYO_DT_FLG" & vbCrLf & _
      ",a.RIYO_YM" & vbCrLf & _
      ",a.RIYO_DT" & vbCrLf & _
      ",a.CANCEL_WAIT_NO" & vbCrLf & _
      ",a.START_TIME" & vbCrLf & _
      ",a.END_TIME" & vbCrLf & _
      ",a.RIYO_MEMO" & vbCrLf & _
      ",a.WAKU_NO" & vbCrLf & _
      ",a.RIYO_KEITAI" & vbCrLf & _
      ",a.MITEI_FLG" & vbCrLf & _
      ",b.SAIJI_NM" & vbCrLf & _
      ",b.RIYO_NM" & vbCrLf & _
      ",b.SHUTSUEN_NM" & vbCrLf & _
    "from" & vbCrLf & _
      "CANCEL_WAIT_DT_TBL a" & vbCrLf & _
    "inner join CANCEL_WAIT_TBL b on(" & vbCrLf & _
      "a.CANCEL_WAIT_NO = b.CANCEL_WAIT_NO" & vbCrLf & _
    ")" & vbCrLf
    Private strSelectCancelListSQL_WHERE As String = String.Empty

    Private strSelectCancelListSQL_ORDER As String = _
    "order by " & vbCrLf & _
      "a.RIYO_DT, a.WAKU_NO"

#End Region

#Region "SQL文(キャンセル待ち未確定一覧取得)"

    'SQL文(キャンセル待ち未確定一覧取得)
    Private strSelectMikakuCancelWaitListSQL_SELECT As String = _
    "select" & vbCrLf & _
    "   USER_CD" & vbCrLf & _
    "  ,SHISETU_KBN" & vbCrLf & _
    "  ,STUDIO_KBN" & vbCrLf & _
    "  ,YOYAKU_DT" & vbCrLf & _
    "from" & vbCrLf & _
    "  CCTL_TBL" & vbCrLf
    Private strSelectMikakuCancelWaitListSQL_WHERE As String = _
    "where" & vbCrLf & _
    "  SHISETU_KBN = '2'" & vbCrLf & _
    "  and substr(YOYAKU_DT,1,7) = :YOYAKU_DT"

#End Region

#Region "SQL文(休館日、メンテ日一覧取得)"

    'SQL文(休館日、メンテ日一覧取得)
    Private strSelectKyukanMainteListSQL_SELECT As String = _
        "select" & vbCrLf & _
        "  HOLMENT_DT" & vbCrLf & _
        "  ,SHISETU_KBN" & vbCrLf & _
        "  ,STUDIO_KBN" & vbCrLf & _
        "  ,MNAIYO" & vbCrLf & _
        "  ,HOLMENT_KBN" & vbCrLf & _
        "from" & vbCrLf & _
        "  HOLMENT_MST" & vbCrLf
    Private strSelectKyukanMainteListSQL_WHERE As String = _
        "where" & vbCrLf & _
        "  SHISETU_KBN = '2'" & vbCrLf & _
        "  and substr(HOLMENT_DT,1,7) = :HOLMENT_DT"

#End Region

#Region "SQL文(休館・メンテ解除)"

    'SQL文(休館・メンテ登録)
    Private strDeleteKyukanMainteSQL As String = _
        "delete from HOLMENT_MST" & vbCrLf & _
        "where" & vbCrLf & _
        "  SHISETU_KBN = '2'" & vbCrLf & _
        "  and HOLMENT_DT = :HOLMENT_DT" & vbCrLf

#End Region

#Region "SQL文(祝祭日登録解除判定)"

    'SQL文(祝祭日登録解除判定)
    Private strJudgeHolidaySQL As String = _
        "select" & vbCrLf & _
        "  HOLIDAY_DT" & vbCrLf & _
        "from" & vbCrLf & _
        "  HOLIDAY_MST" & vbCrLf & _
        "where" & vbCrLf & _
        "  HOLIDAY_DT = :HOLIDAY_DT"

#End Region

#Region "SQL文(祝祭日登録)"

    'SQL文(祝祭日登録)
    Private strInsertHolidaySQL As String = _
        "insert into HOLIDAY_MST values(" & vbCrLf & _
          ":HOLIDAY_DT" & vbCrLf & _
          ",CURRENT_TIMESTAMP" & vbCrLf & _
          ",:USER_CD" & vbCrLf & _
          ",CURRENT_TIMESTAMP" & vbCrLf & _
          ",:USER_CD" & vbCrLf & _
        ")"

#End Region

#Region "SQL文(祝祭日解除)"

    'SQL文(祝祭日解除)
    Private strDeleteHolidaySQL As String = _
        "delete from HOLIDAY_MST" & vbCrLf & _
        "where" & vbCrLf & _
        "  HOLIDAY_DT = :HOLIDAY_DT"

#End Region

#Region "SQL(予約受付制御登録)"
    ' 2015.11.19 ADD START↓ h.hagiwara
    'SQL(予約受付制御登録)
    Private strYCTL_INSERT As String = _
                           "insert into YCTL_TBL(user_cd, shisetu_kbn, studio_kbn, yoyaku_dt, add_dt, add_user_cd)" & vbCrLf & _
                           "values( " & vbCrLf & _
                               ":UserId, " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":YoyakuDt, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd " & vbCrLf & _
                           ") "
    ' 2015.11.19 ADD END↑ h.hagiwara
#End Region

#Region "SQL(キャンセル予約受付制御登録)"
    ' 2015.11.19 ADD START↓ h.hagiwara
    'SQL(キャンセル予約受付制御登録)
    Private strCCTL_INSERT As String = _
                           "insert into CCTL_TBL(user_cd, shisetu_kbn, studio_kbn, yoyaku_dt, add_dt, add_user_cd)" & vbCrLf & _
                           "values( " & vbCrLf & _
                               ":UserId, " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":YoyakuDt, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd " & vbCrLf & _
                           ") "
    ' 2015.11.19 ADD END↑ h.hagiwara
#End Region


    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "祝祭日取得"

    ''' <summary>
    ''' 祝祭日取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectShukusaijitsu(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strSelectShukusaijitsuSQL

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":HOLIDAY_DT").Value = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth      '祝祭日
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

#Region "予約一覧取得"

    ''' <summary>
    ''' 予約一覧取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101">データクラス</param>
    ''' <param name="Mode">利用日検索モード</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>予約一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectYoyakuList(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101, _
                                       ByVal Mode As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If Mode = YOYAKU_NENGAPPI Then
                If dataEXTC0101.PropStrStudioKbn Is Nothing Then
                    strSelectYoyakuListSQL_WHERE = _
                         "where" & vbCrLf & _
                         "  a.SHISETU_KBN = '2'" & vbCrLf & _
                         "  and a.YOYAKU_DT = :YOYAKU_DT" & vbCrLf & _
                         "  and b.YOYAKU_STS IN ('1', '2', '3', '4')" & vbCrLf
                Else
                    strSelectYoyakuListSQL_WHERE = _
                         "where" & vbCrLf & _
                         "  a.SHISETU_KBN = '2'" & vbCrLf & _
                         "  and a.STUDIO_KBN IN(:STUDIO_KBN, '3')" & vbCrLf & _
                         "  and a.YOYAKU_DT = :YOYAKU_DT" & vbCrLf & _
                         "  and b.YOYAKU_STS IN ('1', '2', '3', '4')" & vbCrLf
                End If
            Else
                strSelectYoyakuListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "  a.SHISETU_KBN = '2'" & vbCrLf & _
                    "  and substr(a.YOYAKU_DT,1,7) = :YOYAKU_DT" & vbCrLf & _
                    "  and b.YOYAKU_STS IN ('1', '2', '3', '4')" & vbCrLf
            End If
            strSQL = strSelectYoyakuListSQL_SELECT & strSelectYoyakuListSQL_WHERE & strSelectYoyakuListSQL_ORDER

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                If Mode = YOYAKU_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                End If
                .Add(New NpgsqlParameter(":YOYAKU_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                If Mode = YOYAKU_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Parameters(":STUDIO_KBN").Value = dataEXTC0101.PropStrStudioKbn    'スタジオ区分
                    End If
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYYMMDD        '利用日
                Else
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth        '利用日
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

#Region "予約未確定一覧取得"

    ''' <summary>
    ''' 予約未確定一覧取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101"></param>
    ''' <param name="Mode">利用日検索モード</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>予約未確定一覧取得を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectMikakuYoyakuList(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101, _
                                       ByVal Mode As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If Mode = YOYAKU_NENGAPPI Then
                If dataEXTC0101.PropStrStudioKbn Is Nothing Then
                    strSelectMikakuYoyakuListSQL_WHERE = _
                         "where" & vbCrLf & _
                         "  SHISETU_KBN = '2'" & vbCrLf & _
                         "  and YOYAKU_DT= :YOYAKU_DT"
                Else
                    strSelectMikakuYoyakuListSQL_WHERE = _
                        "where" & vbCrLf & _
                        "  SHISETU_KBN = '2'" & vbCrLf & _
                        "  and STUDIO_KBN IN(:STUDIO_KBN, '3')" & vbCrLf & _
                        "  and YOYAKU_DT= :YOYAKU_DT"
                End If
            Else
                strSelectMikakuYoyakuListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "  SHISETU_KBN = '2'" & vbCrLf & _
                    "  and substr(YOYAKU_DT,1,7) = :YOYAKU_DT"
            End If
            strSQL = strSelectMikakuYoyakuListSQL_SELECT & strSelectMikakuYoyakuListSQL_WHERE

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                If Mode = YOYAKU_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                End If
                .Add(New NpgsqlParameter(":YOYAKU_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                If Mode = YOYAKU_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Parameters(":STUDIO_KBN").Value = dataEXTC0101.PropStrStudioKbn    'スタジオ区分
                    End If
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYYMMDD            '利用日
                Else
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth    '利用日
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

#Region "キャンセル一覧取得"

    ''' <summary>
    ''' キャンセル一覧取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101"></param>
    ''' <param name="Mode"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>キャンセル一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectCancelList(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101, _
                                       ByVal Mode As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If Mode = CANCEL_WAITDATEONLY Then
                '日付確定のみ（年月）の場合
                strSelectCancelListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "a.RIYO_YM = :RIYO_YM" & vbCrLf & _
                    "and a.SHISETU_KBN = '2'" & vbCrLf & _
                    "and a.RIYO_DT_FLG = :RIYO_DT_FLG" & vbCrLf & _
                    "and b.CANCEL_WAIT_STS = :CANCEL_WAIT_STS" & vbCrLf

            ElseIf Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                '日付確定のみ（年月日）の場合
                If dataEXTC0101.PropStrStudioKbn Is Nothing Then
                    strSelectCancelListSQL_WHERE = _
                     "where" & vbCrLf & _
                     "a.RIYO_DT = :RIYO_DT" & vbCrLf & _
                     "and a.SHISETU_KBN = '2'" & vbCrLf & _
                     "and a.RIYO_DT_FLG = :RIYO_DT_FLG" & vbCrLf & _
                     "and b.CANCEL_WAIT_STS = :CANCEL_WAIT_STS" & vbCrLf
                Else
                    strSelectCancelListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "a.RIYO_DT = :RIYO_DT" & vbCrLf & _
                    "and a.SHISETU_KBN = '2'" & vbCrLf & _
                    "and a.STUDIO_KBN IN(:STUDIO_KBN, '3')" & vbCrLf & _
                    "and a.RIYO_DT_FLG = :RIYO_DT_FLG" & vbCrLf & _
                    "and b.CANCEL_WAIT_STS = :CANCEL_WAIT_STS" & vbCrLf
                End If
            Else
                '上記以外場合
                strSelectCancelListSQL_WHERE = _
                "where" & vbCrLf & _
                "a.RIYO_YM = :RIYO_YM" & vbCrLf & _
                "and a.SHISETU_KBN = '2'" & vbCrLf & _
                "and b.CANCEL_WAIT_STS = :CANCEL_WAIT_STS" & vbCrLf
            End If
                strSQL = strSelectCancelListSQL_SELECT & strSelectCancelListSQL_WHERE & strSelectCancelListSQL_ORDER

                Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

                '条件項目の型を指定
                'バインド変数の型を設定
                With Adapter.SelectCommand.Parameters
                    If Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                        .Add(New NpgsqlParameter(":RIYO_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
                        If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                            .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))
                        End If
                    Else
                        .Add(New NpgsqlParameter(":RIYO_YM", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                    If Mode = CANCEL_WAITDATEONLY OrElse Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                        .Add(New NpgsqlParameter(":RIYO_DT_FLG", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                    .Add(New NpgsqlParameter(":CANCEL_WAIT_STS", NpgsqlTypes.NpgsqlDbType.Varchar))
                End With

                'バインド変数の値を設定
                With Adapter.SelectCommand
                    If Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                        .Parameters(":RIYO_DT").Value = dataEXTC0101.PropStrYYMMDD          '利用日
                        If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                            .Parameters(":STUDIO_KBN").Value = dataEXTC0101.PropStrStudioKbn    'スタジオ区分
                        End If
                    Else
                        .Parameters(":RIYO_YM").Value = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth      '利用年月
                    End If
                    If Mode = CANCEL_WAITDATEONLY OrElse Mode = CANCEL_WAITDATEONLY_NENGAPPI Then      '利用日有無
                        '日付確定のみの場合、有に設定
                        .Parameters(":RIYO_DT_FLG").Value = "1"
                    End If
                    If Mode = CANCEL_WAITALL OrElse Mode = CANCEL_WAITDATEONLY OrElse Mode = CANCEL_WAITDATEONLY_NENGAPPI Then     'キャンセル待ちステータス
                        'キャンセル待ちに設定
                        .Parameters(":CANCEL_WAIT_STS").Value = "1"
                    Else
                        'キャンセル済に設定
                        .Parameters(":CANCEL_WAIT_STS").Value = "3"
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

#Region "キャンセル待ち未確定一覧取得"

    ''' <summary>
    ''' キャンセル待ち未確定一覧取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>キャンセル待ち未確定一覧取得を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectMikakuCancelWaitList(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101, _
                                       ByVal Mode As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                If dataEXTC0101.PropStrStudioKbn Is Nothing Then
                    strSelectMikakuCancelWaitListSQL_WHERE = _
                        "where" & vbCrLf & _
                        "  SHISETU_KBN = '2'" & vbCrLf & _
                         "  and YOYAKU_DT = :YOYAKU_DT"
                Else
                    strSelectMikakuCancelWaitListSQL_WHERE = _
                        "where" & vbCrLf & _
                        "  SHISETU_KBN = '2'" & vbCrLf & _
                        "  and STUDIO_KBN IN(:STUDIO_KBN, '3')" & vbCrLf & _
                        "  and YOYAKU_DT = :YOYAKU_DT"
                End If
            Else
                strSelectMikakuCancelWaitListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "  SHISETU_KBN = '2'" & vbCrLf & _
                    "  and substr(YOYAKU_DT,1,7) = :YOYAKU_DT"
            End If
            strSQL = strSelectMikakuCancelWaitListSQL_SELECT & strSelectMikakuCancelWaitListSQL_WHERE

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                If Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                End If
                .Add(New NpgsqlParameter(":YOYAKU_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                If Mode = CANCEL_WAITDATEONLY_NENGAPPI Then
                    If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                        .Parameters(":STUDIO_KBN").Value = dataEXTC0101.PropStrStudioKbn    'スタジオ区分
                    End If
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYYMMDD    '利用日
                Else
                    .Parameters(":YOYAKU_DT").Value = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth    '利用日
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

#Region "休館日、メンテ日一覧取得"

    ''' <summary>
    ''' 休館日、メンテ日一覧取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0101"></param>
    ''' <param name="Mode">休館メンテ日検索モード</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館日、メンテ日一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectKyukanMainteList(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0101 As DataEXTC0101, _
                                       ByVal Mode As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If Mode = KYUKANMAINTE_NENGAPPI Then
                If dataEXTB0101.PropStrStudioKbn Is Nothing Then
                    strSelectKyukanMainteListSQL_WHERE = _
                        "where" & vbCrLf & _
                        "  SHISETU_KBN = '2'" & vbCrLf & _
                        "  and HOLMENT_DT = :HOLMENT_DT"
                Else
                    strSelectKyukanMainteListSQL_WHERE = _
                        "where" & vbCrLf & _
                        "  SHISETU_KBN = '2'" & vbCrLf & _
                        "  and STUDIO_KBN = :STUDIO_KBN" & vbCrLf & _
                        "  and HOLMENT_DT = :HOLMENT_DT"
                End If
            Else
                strSelectKyukanMainteListSQL_WHERE = _
                    "where" & vbCrLf & _
                    "  SHISETU_KBN = '2'" & vbCrLf & _
                    "  and substr(HOLMENT_DT,1,7) = :HOLMENT_DT"
            End If
            strSQL = strSelectKyukanMainteListSQL_SELECT & strSelectKyukanMainteListSQL_WHERE

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                If Mode = KYUKANMAINTE_NENGAPPI Then
                    If Not dataEXTB0101.PropStrStudioKbn Is Nothing Then
                        .Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End If
                End If
                .Add(New NpgsqlParameter(":HOLMENT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                If Mode = KYUKANMAINTE_NENGAPPI Then
                    If Mode = KYUKANMAINTE_NENGAPPI Then
                        If Not dataEXTB0101.PropStrStudioKbn Is Nothing Then
                            .Parameters(":STUDIO_KBN").Value = dataEXTB0101.PropStrStudioKbn 'スタジオ区分
                        End If
                    End If
                    .Parameters(":HOLMENT_DT").Value = dataEXTB0101.PropStrYYMMDD '利用日
                Else
                    .Parameters(":HOLMENT_DT").Value = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth '利用日
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

#Region "休館・メンテ解除"

    ''' <summary>
    ''' 休館・メンテ解除
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>休館メンテマスタ削除を行うSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setDeleteKyukanMainte(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0101 As DataEXTC0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strDeleteKyukanMainteSQL
            If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                strSQL &= "  AND STUDIO_KBN = :STUDIO_KBN"
            End If

            'SQL文(INSERT)
            'SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":HOLMENT_DT", NpgsqlTypes.NpgsqlDbType.Varchar))        '対象日
            If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                cmd.Parameters.Add(New NpgsqlParameter(":STUDIO_KBN", NpgsqlTypes.NpgsqlDbType.Varchar))    'スタジオ区分
            End If

            '値を設定
            cmd.Parameters(":HOLMENT_DT").Value = dataEXTC0101.PropStrYYMMDD            '対象日
            If Not dataEXTC0101.PropStrStudioKbn Is Nothing Then
                cmd.Parameters(":STUDIO_KBN").Value = dataEXTC0101.PropStrStudioKbn     'スタジオ区分
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

#Region "祝祭日登録解除判定"

    ''' <summary>
    ''' 祝祭日登録解除判定
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日登録解除を判定するSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setJudgeHoliday(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0101 As DataEXTC0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strJudgeHolidaySQL

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":HOLIDAY_DT").Value = dataEXTB0101.PropStrYYMMDD       '祝祭日
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

#Region "祝祭日登録"

    ''' <summary>
    ''' 祝祭日登録
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日登録を行うSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setInsertHoliday(ByRef cmd As NpgsqlCommand, _
                                     ByRef Cn As NpgsqlConnection, _
                                     ByVal dataEXTB0101 As DataEXTC0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strInsertHolidaySQL

            'SQL文(INSERT)
            'SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))        '祝祭日
            cmd.Parameters.Add(New NpgsqlParameter(":USER_CD", NpgsqlTypes.NpgsqlDbType.Varchar))           'ユーザーID

            '値を設定
            cmd.Parameters(":HOLIDAY_DT").Value = dataEXTB0101.PropStrYYMMDD        '祝祭日
            cmd.Parameters(":USER_CD").Value = CommonEXT.PropComStrUserId           'ユーザーID

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

#Region "祝祭日解除"

    ''' <summary>
    ''' 祝祭日解除
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日マスタ削除を行うSQLの作成
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setDeleteHoliday(ByRef cmd As NpgsqlCommand, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0101 As DataEXTC0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            strSQL = strDeleteHolidaySQL

            'SQL文(INSERT)
            'SQLのINSERT文を設定
            cmd = New NpgsqlCommand(strSQL, Cn)


            '条件項目の型を指定
            cmd.Parameters.Add(New NpgsqlParameter(":HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))       '祝祭日

            '値を設定
            cmd.Parameters(":HOLIDAY_DT").Value = dataEXTB0101.PropStrYYMMDD        '祝祭日

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


#Region "予約制御表登録"

    ''' <summary>
    ''' 予約受付制御テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015.11.19 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertYoyakuCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As String, _
                                   ByRef dataEXTC0101 As DataEXTC0101) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strYCTL_INSERT
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(1).Value = "2"
            If dataEXTC0101.PropStrStudioKbn = "" Then
                Cmd.Parameters(2).Value = "0"
            Else
                Cmd.Parameters(2).Value = dataEXTC0101.PropStrStudioKbn
            End If
            Cmd.Parameters(3).Value = dataRiyobi
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

#End Region

#Region "キャンセル予約制御表登録"

    ''' <summary>
    ''' キャンセル予約受付制御テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015.11.19 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertCancelYoyakuCtl(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As String, _
                                   ByRef dataEXTC0101 As DataEXTC0101) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strCCTL_INSERT
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":UserId", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(1).Value = "2"
            If dataEXTC0101.PropStrStudioKbn = "" Then
                Cmd.Parameters(2).Value = "0"
            Else
                Cmd.Parameters(2).Value = dataEXTC0101.PropStrStudioKbn
            End If
            Cmd.Parameters(3).Value = dataRiyobi
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

#End Region

End Class
