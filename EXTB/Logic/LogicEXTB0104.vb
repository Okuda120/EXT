Imports Common
Imports CommonEXT
Imports Npgsql
Imports EXTZ

Public Class LogicEXTB0104

    '変数宣言
    Private sqlEXTB0104 As New SqlEXTB0104          'sqlクラス

    ''' <summary>
    ''' 予約情報取得処理(共通にする)
    ''' </summary>
    ''' <param name="dataEXTB0104"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetYoyakuData(ByRef dataEXTB0104 As DataEXTB0104) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTB0104.setSelectYoyakuData(Adapter, Cn, dataEXTB0104.PropStrYoyakuNo) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得", Nothing, Adapter.SelectCommand)
            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            '値の設定
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START_2", Nothing, Nothing)
            TransData(dataEXTB0104, Table)
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END_2", Nothing, Nothing)
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' DB項目値から画面への変換
    ''' 
    ''' </summary>
    ''' <param name="dataEXTB0104"></param>
    ''' <param name="Table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TransData(ByRef dataEXTB0104 As DataEXTB0104, ByVal Table As DataTable) As Boolean
        Try
            With dataEXTB0104
                .PropStrYoyakuNo = IIf(DBNull.Value.Equals(Table.Rows(0).Item("CANCEL_WAIT_NO")), "", Table.Rows(0).Item("CANCEL_WAIT_NO"))
                .PropStrCanUkeDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("CANCEL_WAIT_DT")), "", Table.Rows(0).Item("CANCEL_WAIT_DT"))
                .PropStrCanUkeUsercd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("CANCEL_WAIT_USERCD")), "", Table.Rows(0).Item("CANCEL_WAIT_USERCD"))
                .PropStrYoyakuSts = IIf(DBNull.Value.Equals(Table.Rows(0).Item("CANCEL_WAIT_STS")), "", Table.Rows(0).Item("CANCEL_WAIT_STS"))
                .PropStrShisetuKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SHISETU_KBN")), "", Table.Rows(0).Item("SHISETU_KBN"))
                .PropStrStudioKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("STUDIO_KBN")), "", Table.Rows(0).Item("STUDIO_KBN"))
                .PropStrSaijiNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SAIJI_NM")), "", Table.Rows(0).Item("SAIJI_NM"))
                .PropStrShutsuenNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SHUTSUEN_NM")), "", Table.Rows(0).Item("SHUTSUEN_NM"))
                .PropStrKashiKind = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KASHI_KIND")), "", Table.Rows(0).Item("KASHI_KIND"))
                .PropStrRiyoType = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TYPE")), "", Table.Rows(0).Item("RIYO_TYPE"))
                .PropStrDrinkFlg = IIf(DBNull.Value.Equals(Table.Rows(0).Item("DRINK_FLG")), "", Table.Rows(0).Item("DRINK_FLG"))
                .PropStrSaijiBunrui = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SAIJI_BUNRUI")), "", Table.Rows(0).Item("SAIJI_BUNRUI"))
                .PropStrTeiin = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TEIIN")), "", Table.Rows(0).Item("TEIIN"))
                .PropStrRiyoshaCd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYOSHA_CD")), "", Table.Rows(0).Item("RIYOSHA_CD"))
                .PropStrRiyoNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_NM")), "", Table.Rows(0).Item("RIYO_NM"))
                .PropStrRiyoKana = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_KANA")), "", Table.Rows(0).Item("RIYO_KANA"))
                .PropStrSekininBushoNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEKININ_BUSHO_NM")), "", Table.Rows(0).Item("SEKININ_BUSHO_NM"))
                .PropStrSekininNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEKININ_NM")), "", Table.Rows(0).Item("SEKININ_NM"))
                .PropStrSekininMail = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEKININ_MAIL")), "", Table.Rows(0).Item("SEKININ_MAIL"))
                .PropStrDaihyoNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("DAIHYO_NM")), "", Table.Rows(0).Item("DAIHYO_NM"))
                .PropStrRiyoTel11 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL11")), "", Table.Rows(0).Item("RIYO_TEL11"))
                .PropStrRiyoTel12 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL12")), "", Table.Rows(0).Item("RIYO_TEL12"))
                .PropStrRiyoTel13 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL13")), "", Table.Rows(0).Item("RIYO_TEL13"))
                .PropStrRiyoTel21 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL21")), "", Table.Rows(0).Item("RIYO_TEL21"))
                '.PropStrRiyoTel22 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL22")), "", Table.Rows(0).Item("RIYO_TEL22")) '2016.11.4 m.hayabuchi DEL 課題No.59
                '.PropStrRiyoTel23 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL23")), "", Table.Rows(0).Item("RIYO_TEL23")) '2016.11.4 m.hayabuchi DEL 課題No.59
                .PropStrRiyoNaisen = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_NAISEN")), "", Table.Rows(0).Item("RIYO_NAISEN"))
                .PropStrRiyoFax11 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_FAX11")), "", Table.Rows(0).Item("RIYO_FAX11"))
                .PropStrRiyoFax12 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_FAX12")), "", Table.Rows(0).Item("RIYO_FAX12"))
                .PropStrRiyoFax13 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_FAX13")), "", Table.Rows(0).Item("RIYO_FAX13"))
                .PropStrRiyoYubin1 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_YUBIN1")), "", Table.Rows(0).Item("RIYO_YUBIN1"))
                .PropStrRiyoYubin2 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_YUBIN2")), "", Table.Rows(0).Item("RIYO_YUBIN2"))
                .PropStrRiyoTodo = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TODO")), "", Table.Rows(0).Item("RIYO_TODO"))
                .PropStrRiyoShiku = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_SHIKU")), "", Table.Rows(0).Item("RIYO_SHIKU"))
                .PropStrRiyoBan = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_BAN")), "", Table.Rows(0).Item("RIYO_BAN"))
                .PropStrRiyoBuild = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_BUILD")), "", Table.Rows(0).Item("RIYO_BUILD"))
                .PropStrRiyoLvl = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_LVL")), "", Table.Rows(0).Item("RIYO_LVL"))
                .PropStrAiteCd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("AITE_CD")), "", Table.Rows(0).Item("AITE_CD"))
                .PropStrBiko = IIf(DBNull.Value.Equals(Table.Rows(0).Item("BIKO")), "", Table.Rows(0).Item("BIKO"))
                .PropStrAddDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ADD_DT")), "", Table.Rows(0).Item("ADD_DT"))
                .PropStrAddUserCd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ADD_USER_CD")), "", Table.Rows(0).Item("ADD_USER_CD"))
                .PropStrAddUserNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ADD_USER_NM")), "", Table.Rows(0).Item("ADD_USER_NM"))
                .PropStrUpDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("UP_DT")), "", Table.Rows(0).Item("UP_DT"))
                .PropStrUpUserCd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("UP_USER_CD")), "", Table.Rows(0).Item("UP_USER_CD"))
                .PropStrUpUserNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("UP_USER_NM")), "", Table.Rows(0).Item("UP_USER_NM"))
            End With
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            Return False
        End Try
        Return True
    End Function


    ''' <summary>
    ''' 予約日時情報取得(共通にする)
    ''' </summary>
    ''' <param name="dataEXTB0104"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRiyobiData(ByRef dataEXTB0104 As DataEXTB0104) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTB0104.setSelectRiyobiData(Adapter, Cn, dataEXTB0104.PropStrYoyakuNo) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約日時情報取得", Nothing, Adapter.SelectCommand)

            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If
            '値の設定
            TransDataRiyobiList(dataEXTB0104, Table)

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 予約日時情報をリストにつめる
    ''' </summary>
    ''' <param name="dataEXTB0104"></param>
    ''' <param name="Table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TransDataRiyobiList(ByRef dataEXTB0104 As DataEXTB0104, ByVal Table As DataTable) As Boolean

        Dim dataRiyobi As New CommonDataCancel
        Dim lstRiyobi As New ArrayList
        Dim dateDispRiyobi As New Date
        Try
            For i = 0 To Table.Rows.Count - 1
                With dataRiyobi
                    .PropIntSeq = IIf(DBNull.Value.Equals(Table.Rows(i).Item("SEQ")), Nothing, Table.Rows(i).Item("SEQ"))
                    .PropStrShisetuKbn = IIf(DBNull.Value.Equals(Table.Rows(i).Item("SHISETU_KBN")), "", Table.Rows(i).Item("SHISETU_KBN"))
                    .PropStrStudioKbn = IIf(DBNull.Value.Equals(Table.Rows(i).Item("STUDIO_KBN")), "", Table.Rows(i).Item("STUDIO_KBN"))
                    .PropStrKibobiKbn = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_DT_FLG")), "", Table.Rows(i).Item("RIYO_DT_FLG"))
                    .PropStrMiteiFlg = IIf(DBNull.Value.Equals(Table.Rows(i).Item("MITEI_FLG")), "", Table.Rows(i).Item("MITEI_FLG"))
                    .PropStrCancelYm = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_YM")), "", Table.Rows(i).Item("RIYO_YM"))
                    .PropStrCancelDt = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_DT")), "", Table.Rows(i).Item("RIYO_DT"))
                    .PropStrCancelWaitNo = IIf(DBNull.Value.Equals(Table.Rows(i).Item("CANCEL_WAIT_NO")), "", Table.Rows(i).Item("CANCEL_WAIT_NO"))
                    .PropStrStartTime = IIf(DBNull.Value.Equals(Table.Rows(i).Item("START_TIME")), "", Table.Rows(i).Item("START_TIME"))
                    .PropStrEndTime = IIf(DBNull.Value.Equals(Table.Rows(i).Item("END_TIME")), "", Table.Rows(i).Item("END_TIME"))
                    .PropStrMemo = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_MEMO")), "", Table.Rows(i).Item("RIYO_MEMO"))
                    .PropIntWakuNo = IIf(DBNull.Value.Equals(Table.Rows(i).Item("WAKU_NO")), Nothing, Table.Rows(i).Item("WAKU_NO"))
                    .PropStrRegistFlg = "1"
                End With
                If String.IsNullOrEmpty(dataRiyobi.PropStrCancelDt) Then
                    dataRiyobi.PropStrCancelDtDisp = ""
                Else
                    dateDispRiyobi = dataRiyobi.PropStrCancelDt
                    dataRiyobi.PropStrCancelDtDisp = dateDispRiyobi.ToString(CommonDeclareEXT.FMT_DATE_DISP)
                End If
                lstRiyobi.Add(dataRiyobi)
                dataRiyobi = New CommonDataCancel
                dateDispRiyobi = New Date
            Next
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            Return False
        End Try
        dataEXTB0104.PropListRiyobi = lstRiyobi
        Return True
    End Function

    ''' <summary>
    ''' 予約情報登録／更新
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報登録
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegYoyakuInfo(ByRef dataEXTB0104 As DataEXTB0104, _
                                       ByVal blnIsUpdate As Boolean) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            sqlEXTB0104.registerYoyakuInfo(Cmd, Cn, dataEXTB0104, blnIsUpdate)
            'Cmd.ExecuteNonQuery()
            Dim strSeqNo As String = Cmd.ExecuteScalar()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Complate SQL", Nothing, Cmd)
            'COMMIT
            Tsx.Commit()

            'シーケンスで作成した予約NO取得
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "登録予約番号:" & strSeqNo, Nothing, Cmd)
            If blnIsUpdate = False Then
                dataEXTB0104.PropStrYoyakuNo = strSeqNo
            End If

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約日程表削除
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約日程データを予約Noですべて削除する
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteYoyakuList(ByVal YoyakuNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約日程表削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTB0104.deleteYoyakuList(Cmd, Cn, YoyakuNo)
            Cmd.ExecuteNonQuery()

            'エラーがなければコミットする
            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Cmd)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約日程表の登録
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報登録
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function InsertYoyakuList(ByVal yoyakuNo As String, _
                                       ByVal dataList As ArrayList) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            For Each dataRiyobi As CommonDataCancel In dataList
                'ログ出力
                sqlEXTB0104.insertYoyakuList(Cmd, Cn, dataRiyobi, yoyakuNo)
                Cmd.ExecuteNonQuery()
            Next
            Tsx.Commit()
            Cn.Close()

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        End Try
    End Function

    ''' <summary>
    ''' 予約削除
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約データを削除する
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteYoyaku(ByVal YoyakuNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約日程表削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTB0104.deleteYoyaku(Cmd, Cn, YoyakuNo)
            Cmd.ExecuteNonQuery()

            'エラーがなければコミットする
            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Cmd)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 予約受付制御登録チェック
    ''' </summary>
    ''' <param name="dataRiyobi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetWakuNo(ByVal dataRiyobi As CommonDataCancel) As Integer
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Try
            Cn.Open()

            If dataRiyobi.PropIntWakuNo Is Nothing = False And dataRiyobi.PropIntWakuNo > 0 Then
                If sqlEXTB0104.selectWakuNo(Adapter, Cn, dataRiyobi, dataRiyobi.PropIntWakuNo) = True Then
                    Return dataRiyobi.PropIntWakuNo
                End If
            End If
            If sqlEXTB0104.selectWakuNo(Adapter, Cn, dataRiyobi, 1) = True Then
                Return 1
            End If
            If sqlEXTB0104.selectWakuNo(Adapter, Cn, dataRiyobi, 2) = True Then
                Return 2
            End If
            If sqlEXTB0104.selectWakuNo(Adapter, Cn, dataRiyobi, 3) = True Then
                Return 3
            End If
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "枠番号取得失敗", Nothing, Adapter.SelectCommand)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return Nothing

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 責任者名取得
    ''' </summary>
    ''' <param name="strRiyoshaCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSekininshaList(ByVal strRiyoshaCd As String) As ArrayList
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()
        Dim list As New ArrayList
        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTB0104.setSelectSekininName(Adapter, Cn, strRiyoshaCd) = False Then
                '異常終了
                Return list
            End If

            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                '正常終了
                Return list
            Else
                '存在する
                For i = 0 To Table.Rows.Count - 1
                    list.Add(IIf(DBNull.Value.Equals(Table.Rows(i).Item("SEKININ_NM")), "", Table.Rows(i).Item("SEKININ_NM")))
                Next
                Return list
            End If

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return list
        Finally
            '終了処理
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
        Return list
    End Function

    ''' <summary>
    ''' 責任者メールアドレス・携帯電話番号取得
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>責任者名のコンボボックス選択時メールアドレス・携帯電話番号取得
    ''' <para>作成情報：2016/11/4 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetSekininshaMailTel(ByVal dataEXTB0104 As DataEXTB0104) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        Try
            'コネクションを開く
            Cn.Open()

            'SQLの作成・設定
            If sqlEXTB0104.GetSqlSekininshaMailTelData(Adapter, Cn, dataEXTB0104) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "責任者メールアドレス・携帯電話番号情報取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得した項目をDataクラスに格納
            With dataEXTB0104
                'メールアドレス
                .PropStrSekininMail = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEKININ_MAIL")), "", Table.Rows(0).Item("SEKININ_MAIL"))
                '携帯電話番号
                .PropStrRiyoTel21 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL21")), "", Table.Rows(0).Item("RIYO_TEL21"))
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception

            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

End Class
