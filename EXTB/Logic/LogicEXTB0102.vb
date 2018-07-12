Imports Common
Imports CommonEXT
Imports Npgsql
Imports EXTZ

Public Class LogicEXTB0102

    '変数宣言
    Private sqlEXTB0102 As New SqlEXTB0102          'sqlクラス

    ''' <summary>
    ''' 予約情報取得処理(共通にする)
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetYoyakuData(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0102.setSelectYoyakuData(Adapter, Cn, dataEXTB0102.PropStrYoyakuNo) = False Then
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
            TransData(dataEXTB0102, Table)
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
    ''' <param name="dataEXTB0102"></param>
    ''' <param name="Table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TransData(ByRef dataEXTB0102 As DataEXTB0102, ByVal Table As DataTable) As Boolean
        Try
            With dataEXTB0102
                .PropStrYoyakuNo = IIf(DBNull.Value.Equals(Table.Rows(0).Item("YOYAKU_NO")), "", Table.Rows(0).Item("YOYAKU_NO"))
                .PropStrKariukeDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KARIUKE_DT")), "", Table.Rows(0).Item("KARIUKE_DT"))
                .PropStrKariUsercd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KARI_USERCD")), "", Table.Rows(0).Item("KARI_USERCD"))
                .PropStrKakuteiDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KAKUTEI_DT")), "", Table.Rows(0).Item("KAKUTEI_DT"))
                .PropStrKakuUsercd = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KAKU_USERCD")), "", Table.Rows(0).Item("KAKU_USERCD"))
                .PropStrYoyakuSts = IIf(DBNull.Value.Equals(Table.Rows(0).Item("YOYAKU_STS")), "", Table.Rows(0).Item("YOYAKU_STS"))
                .PropStrShisetuKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SHISETU_KBN")), "", Table.Rows(0).Item("SHISETU_KBN"))
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
                .PropStrAiteNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("AITE_NM")), "", Table.Rows(0).Item("AITE_NM"))
                .PropStrSendKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEND_KBN")), "", Table.Rows(0).Item("SEND_KBN"))
                .PropStrSendSts = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEND_STS")), "", Table.Rows(0).Item("SEND_STS"))
                .PropStrSendDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEND_DT")), "", Table.Rows(0).Item("SEND_DT"))
                .PropStrHensoDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("HENSO_DT")), "", Table.Rows(0).Item("HENSO_DT"))
                .PropStrBiko = IIf(DBNull.Value.Equals(Table.Rows(0).Item("BIKO")), "", Table.Rows(0).Item("BIKO"))
                .PropStrOnkyoNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_NM")), "", Table.Rows(0).Item("ONKYO_NM"))
                .PropStrOnkyoTantoNm = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_TANTO_NM")), "", Table.Rows(0).Item("ONKYO_TANTO_NM"))
                .PropStrOnkyoTel11 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_TEL11")), "", Table.Rows(0).Item("ONKYO_TEL11"))
                .PropStrOnkyoTel12 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_TEL12")), "", Table.Rows(0).Item("ONKYO_TEL12"))
                .PropStrOnkyoTel13 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_TEL13")), "", Table.Rows(0).Item("ONKYO_TEL13"))
                .PropStrOnkyoNaisen = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_NAISEN")), "", Table.Rows(0).Item("ONKYO_NAISEN"))
                .PropStrOnkyoFax11 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_FAX11")), "", Table.Rows(0).Item("ONKYO_FAX11"))
                .PropStrOnkyoFax12 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_FAX12")), "", Table.Rows(0).Item("ONKYO_FAX12"))
                .PropStrOnkyoFax13 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_FAX13")), "", Table.Rows(0).Item("ONKYO_FAX13"))
                .PropStrOnkyoMail = IIf(DBNull.Value.Equals(Table.Rows(0).Item("ONKYO_MAIL")), "", Table.Rows(0).Item("ONKYO_MAIL"))
                .PropStrTotalRiyoKin = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TOTAL_RIYO_KIN")), "", Table.Rows(0).Item("TOTAL_RIYO_KIN"))
                .PropStrRiyoCom = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_COM")), "", Table.Rows(0).Item("RIYO_COM"))
                .PropStrTicketEnterKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TICKET_ENTER_KBN")), "", Table.Rows(0).Item("TICKET_ENTER_KBN"))
                .PropStrTicketDrinkKbn = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TICKET_DRINK_KBN")), "", Table.Rows(0).Item("TICKET_DRINK_KBN"))
                .PropStrHpKeisai = IIf(DBNull.Value.Equals(Table.Rows(0).Item("HP_KEISAI")), "", Table.Rows(0).Item("HP_KEISAI"))
                .PropStrJohoKokaiDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("JOHO_KOKAI_DT")), "", Table.Rows(0).Item("JOHO_KOKAI_DT"))
                .PropStrJohoKokaiTime = IIf(DBNull.Value.Equals(Table.Rows(0).Item("JOHO_KOKAI_TIME")), "", Table.Rows(0).Item("JOHO_KOKAI_TIME"))
                .PropStrKokaiDt = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KOKAI_DT")), "", Table.Rows(0).Item("KOKAI_DT"))
                .PropStrKokaiTime = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KOKAI_TIME")), "", Table.Rows(0).Item("KOKAI_TIME"))
                .PropStrFinputSts = IIf(DBNull.Value.Equals(Table.Rows(0).Item("FINPUT_STS")), "", Table.Rows(0).Item("FINPUT_STS"))
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
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRiyobiData(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0102.setSelectRiyobiData(Adapter, Cn, dataEXTB0102.PropStrYoyakuNo) = False Then
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
            TransDataRiyobiList(dataEXTB0102, Table)

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
    ''' <param name="dataEXTB0102"></param>
    ''' <param name="Table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TransDataRiyobiList(ByRef dataEXTB0102 As DataEXTB0102, ByVal Table As DataTable) As Boolean
        Dim dataRiyobi As New CommonDataRiyobi
        Dim lstRiyobi As New ArrayList
        Dim dateDispRiyobi As New Date
        Try
            For i = 0 To Table.Rows.Count - 1
                With dataRiyobi
                    .PropIntSeq = IIf(DBNull.Value.Equals(Table.Rows(i).Item("SEQ")), Nothing, Table.Rows(i).Item("SEQ"))
                    .PropStrShisetuKbn = IIf(DBNull.Value.Equals(Table.Rows(i).Item("SHISETU_KBN")), "", Table.Rows(i).Item("SHISETU_KBN"))
                    .PropStrStudioKbn = IIf(DBNull.Value.Equals(Table.Rows(i).Item("STUDIO_KBN")), "", Table.Rows(i).Item("STUDIO_KBN"))
                    .PropStrYoyakuDt = IIf(DBNull.Value.Equals(Table.Rows(i).Item("YOYAKU_DT")), "", Table.Rows(i).Item("YOYAKU_DT"))
                    .PropStrStartTime = IIf(DBNull.Value.Equals(Table.Rows(i).Item("START_TIME")), "", Table.Rows(i).Item("START_TIME"))
                    .PropStrEndTime = IIf(DBNull.Value.Equals(Table.Rows(i).Item("END_TIME")), "", Table.Rows(i).Item("END_TIME"))
                    .PropStrYoyakuNo = IIf(DBNull.Value.Equals(Table.Rows(i).Item("YOYAKU_NO")), "", Table.Rows(i).Item("YOYAKU_NO"))
                    .PropStrRiyoKeitai = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_KEITAI")), "", Table.Rows(i).Item("RIYO_KEITAI"))
                    .PropStrMiteiFlg = IIf(DBNull.Value.Equals(Table.Rows(i).Item("MITEI_FLG")), "", Table.Rows(i).Item("MITEI_FLG"))
                    .PropIntTanka = IIf(DBNull.Value.Equals(Table.Rows(i).Item("TANKA")), Nothing, Table.Rows(i).Item("TANKA"))
                    .PropDblBairitu = IIf(DBNull.Value.Equals(Table.Rows(i).Item("BAIRITU")), Nothing, Table.Rows(i).Item("BAIRITU"))
                    .PropIntSu = IIf(DBNull.Value.Equals(Table.Rows(i).Item("SU")), Nothing, Table.Rows(i).Item("SU"))
                    .PropIntRiyoKin = IIf(DBNull.Value.Equals(Table.Rows(i).Item("RIYO_KIN")), Nothing, Table.Rows(i).Item("RIYO_KIN"))
                    .PropStrRegistFlg = "1"
                End With
                dateDispRiyobi = dataRiyobi.PropStrYoyakuDt
                dataRiyobi.PropStrYoyakuDtDisp = dateDispRiyobi.ToString(CommonDeclareEXT.FMT_DATE_DISP)
                lstRiyobi.Add(dataRiyobi)
                dataRiyobi = New CommonDataRiyobi
                dateDispRiyobi = New Date
            Next
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            Return False
        End Try
        dataEXTB0102.PropListRiyobi = lstRiyobi
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
    Public Function RegYoyakuInfo(ByRef dataEXTB0102 As DataEXTB0102, _
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
            sqlEXTB0102.registerYoyakuInfo(Cmd, Cn, dataEXTB0102, blnIsUpdate)
            'Cmd.ExecuteNonQuery()
            Dim strSeqNo As String = Cmd.ExecuteScalar()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Complate SQL", Nothing, Cmd)
            'COMMIT
            Tsx.Commit()

            'シーケンスで作成した予約NO取得
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "登録予約番号:" & strSeqNo, Nothing, Cmd)
            If blnIsUpdate = False Then
                dataEXTB0102.PropStrYoyakuNo = strSeqNo
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
            sqlEXTB0102.deleteYoyakuList(Cmd, Cn, YoyakuNo)
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
            For Each dataRiyobi As CommonDataRiyobi In dataList
                'ログ出力
                sqlEXTB0102.insertYoyakuList(Cmd, Cn, dataRiyobi, yoyakuNo)
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
            sqlEXTB0102.deleteYoyaku(Cmd, Cn, YoyakuNo)
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
    ''' 利用者処理チェック
    ''' </summary>
    ''' <param name="strRiyoshaKana"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkRiyoshaRegister(ByVal strRiyoshaKana As String) As Boolean
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
            If sqlEXTB0102.setSelectRiyosha(Adapter, Cn, strRiyoshaKana) = False Then
                '異常終了
                Return False
            End If

            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                '正常終了
                Return True
            Else
                '存在する
                Return False
            End If

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
        Return False
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
            If sqlEXTB0102.setSelectSekininName(Adapter, Cn, strRiyoshaCd) = False Then
                '異常終了
                Return list
            End If

            '予約情報表から責任者名を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合(コンボボックス作成)
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
    ''' <para>作成情報：2016/11/04 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetSekininshaMailTel(ByVal dataEXTB0102 As DataEXTB0102) As Boolean

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
            If sqlEXTB0102.GetSqlSekininshaMailTelData(Adapter, Cn, dataEXTB0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "責任者メールアドレス・携帯電話番号情報取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得した項目をDataクラスに格納
            With dataEXTB0102
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
