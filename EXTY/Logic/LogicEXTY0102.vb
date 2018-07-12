Imports Common
Imports CommonEXT
Imports System.Text
Imports System.IO
Imports Npgsql
Imports FarPoint.Win.Spread

Public Class LogicEXTY0102

    Private sqlEXTY0102 As New SqlEXTY0102          'SQLクラス
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT    'EXT共通ロジッククラス

    ''' <summary>
    ''' 電子マネー用コンボ情報取得処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>電子マネー用のコンボボックス（レジ・店舗）の情報を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetCmbInf(ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'コネクションを開く
            Cn.Open()

            ' レジ情報の取得
            If GetMastaALSOKRegister(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' 店舗情報の取得
            If GetMastaTenpo(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' フォーム情報の初期化
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>フォームの情報を初期化する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitForm(ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim DtIncAlsok As New DataTable          '
        Dim DtIncCash As New DataTable           '

        Try
            With dataEXTY0102
                If .PropVwAlsokData.ActiveSheet.RowCount > 0 Then
                    .PropVwAlsokData.Sheets(0).Rows.Remove(0, .PropVwAlsokData.ActiveSheet.RowCount)
                End If
                If .PropVwDegitalCashData.ActiveSheet.RowCount > 0 Then
                    .PropVwDegitalCashData.Sheets(0).Rows.Remove(0, .PropVwDegitalCashData.ActiveSheet.RowCount)
                End If

                .PropTxtFilePath.Text = Nothing
                .PropDtpDspFrom.txtDate.Text = Nothing
                .PropDtpDspTo.txtDate.Text = Nothing
            End With

            With DtIncAlsok
                .Columns.Add("DepositCd", Type.GetType("System.String"))                ' 入金機コード
                .Columns.Add("Seq", Type.GetType("System.String"))                      ' 連番
                .Columns.Add("DepositKbn", Type.GetType("System.String"))               ' 入金区分
                .Columns.Add("DepositDt", Type.GetType("System.String"))                ' 入金日
                .Columns.Add("YoyakuNo", Type.GetType("System.String"))                 ' 予約NO_
                .Columns.Add("RegisterCd", Type.GetType("System.String"))               ' レジコード
                .Columns.Add("TenpoCd", Type.GetType("System.String"))                  ' 店舗コード
                .Columns.Add("DepositAmount", Type.GetType("System.String"))            ' 入金額
                .Columns.Add("KeiyakusakiCd", Type.GetType("System.String"))            ' 契約先コード
                'テーブルの変更を確定
                .AcceptChanges()
            End With

            With DtIncCash
                .Columns.Add("DepositCd", Type.GetType("System.String"))                ' 入金機コード
                .Columns.Add("Seq", Type.GetType("System.String"))                      ' 連番
                .Columns.Add("DepositKbn", Type.GetType("System.String"))               ' 入金区分
                .Columns.Add("DepositDt", Type.GetType("System.String"))                ' 入金日
                .Columns.Add("YoyakuNo", Type.GetType("System.String"))                 ' 予約NO_
                .Columns.Add("RegisterCd", Type.GetType("System.String"))               ' レジコード
                .Columns.Add("TenpoCd", Type.GetType("System.String"))                  ' 店舗コード
                .Columns.Add("DepositAmount", Type.GetType("System.String"))            ' 入金額
                .Columns.Add("KeiyakusakiCd", Type.GetType("System.String"))            ' 契約先コード
                'テーブルの変更を確定
                .AcceptChanges()
            End With

            dataEXTY0102.PropDtAlsokDataSet = DtIncAlsok
            dataEXTY0102.PropDtDegitalCashDataSet = DtIncCash

            ' コンボボックス用データクラス作成
            If VwDipositCmbCre(dataEXTY0102) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' ダイアログ表示
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>ファイルダイアログを表示し取り込むＣＳＶファイルを選択する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function DspDialog(ByRef dataEXTY0102 As DataEXTY0102)

        Dim strInitFile As String = ""                                ' 初期表示ファイル名
        Dim strInitPath As String = ""                                ' 初期表示ディレクトリ
        Dim strFileType As String = "CSVファイル(*.csv)|*.csv"        ' 選択可能なファイル形式
        Dim intSelFileType As Integer = 1
        Dim strFilePath As String = ""

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        If commonLogicEXT.OpenFileDialog(strInitFile, strInitPath, strFileType, intSelFileType, strFilePath) = False Then

            Return False

        End If
        If strFilePath <> "" Then
            dataEXTY0102.PropTxtFilePath.Text = strFilePath
        End If

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' ファイルパス＆ＣＳＶファイルのチェック
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>ダイアログ指定 または 入力されたＣＳＶファイルの存在チェック
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CsvInputCheck(ByRef dataEXTY0102 As DataEXTY0102)

        Dim strPathText As String = ""

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 必須チェック
        If dataEXTY0102.PropTxtFilePath.Text.Trim() = "" Then
            puErrMsg = Y0102_E003
            Return False
        End If

        ' ファイル存在チェック
        If File.Exists(dataEXTY0102.PropTxtFilePath.Text) = False Then
            puErrMsg = Y0102_E004
            Return False
        End If

        ' 拡張子チェック
        If Path.GetExtension(dataEXTY0102.PropTxtFilePath.Text) = CON_Y0102_CSV1 Or _
           Path.GetExtension(dataEXTY0102.PropTxtFilePath.Text) = CON_Y0102_CSV2 Then
        Else
            puErrMsg = Y0102_E005
            Return False
        End If

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' ＣＳＶファイル取り込み処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>指定されたＣＳＶファイルをスプレッドに編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetVwAlsokData(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 変数宣言
        Dim intDupCnt As Integer = 0        ' 重複があった件数
        Dim aryDupData As New ArrayList     ' 重複のあった入金コード＆連番を格納
        Dim intSetCnt As Integer = 0        ' スプレッド上に編集している件数
        Dim strDspsaiji As String = ""
        Dim intUnionStart As Integer = 0
        Dim intUnionCnt As Integer = 1
        Dim intUnionAllyen As Integer = 0
        Dim strUnionDt As String = ""
        Dim dt As DateTime

        Try
            With dataEXTY0102
                If .PropVwAlsokData.ActiveSheet.RowCount > 0 Then
                    .PropVwAlsokData.Sheets(0).Rows.Remove(0, .PropVwAlsokData.ActiveSheet.RowCount)
                End If
            End With

            ' CSV読み込みクラスのインスタンス作成
            Using txtParser = New FileIO.TextFieldParser(dataEXTY0102.PropTxtFilePath.Text, System.Text.Encoding.GetEncoding(932))

                txtParser.TextFieldType = FileIO.FieldType.Delimited
                txtParser.SetDelimiters(",")

                '読み込む行がなくなるまで繰り返し
                Dim strAryBuffer As String() = Nothing

                While Not txtParser.EndOfData
                    strAryBuffer = txtParser.ReadFields()

                    'If strAryBuffer.Length <> CON_Y0102_CSV_ITEMCNT Then
                    '    puErrMsg = Y0102_E009
                    '    Return False
                    'End If

                    ' 入金機コード・連番を取得し重複チェックを行う
                    dataEXTY0102.PropStrDepositCd = strAryBuffer(0)        ' 入金機コード
                    'dataEXTY0102.PropStrSeq = strAryBuffer(2)              ' 連番
                    dataEXTY0102.PropStrSeq = strAryBuffer(3)              ' 連番
                    dt = Date.Parse(Format(CLng(strAryBuffer(1)), "0000/00/00"))
                    strAryBuffer(1) = dt.ToShortDateString

                    If DepositSeqDupCheck(dataEXTY0102) = True Then
                        If intSetCnt > 0 Then
                            If strUnionDt = strAryBuffer(1) Then
                                'intUnionAllyen += Integer.Parse(strAryBuffer(6))
                                intUnionAllyen += Integer.Parse(strAryBuffer(17))
                                intUnionCnt += 1
                            Else
                                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).RowSpan = intUnionCnt
                                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).Value = intUnionAllyen
                                strUnionDt = strAryBuffer(1)
                                'intUnionAllyen = Integer.Parse(strAryBuffer(6))
                                intUnionAllyen = Integer.Parse(strAryBuffer(17))
                                intUnionCnt = 1
                                intUnionStart = intSetCnt
                            End If
                        Else
                            strUnionDt = strAryBuffer(1)
                            'intUnionAllyen = Integer.Parse(strAryBuffer(6))
                            intUnionAllyen = Integer.Parse(strAryBuffer(17))
                            intUnionCnt = 1
                        End If

                        ' スプレッド編集
                        With dataEXTY0102.PropVwAlsokData.Sheets(0)
                            .Rows.Add(.RowCount, 1)
                            '.Cells(intSetCnt, COL_ALSOK_NYUKIN_DATE).Value = strAryBuffer(1)     ' 入金日
                            '.Cells(intSetCnt, COL_ALSOK_REGISTER_NO).Value = strAryBuffer(5)     ' レジ№
                            '.Cells(intSetCnt, COL_ALSOK_TENPO_NO).Value = strAryBuffer(4)        ' 店舗№
                            '.Cells(intSetCnt, COL_ALSOK_NYUKIN_GAKU).Value = strAryBuffer(6)     ' 入金額
                            '.Cells(intSetCnt, COL_ALSOK_DEPOSIT_NO).Value = strAryBuffer(0)      ' 入金機№
                            '.Cells(intSetCnt, COL_ALSOK_DEPOSIT_SEQ).Value = strAryBuffer(2)     ' 入金機№連番
                            '.Cells(intSetCnt, COL_ALSOK_KEIYAKUSAKI).Value = strAryBuffer(3)     ' 契約先
                            .Cells(intSetCnt, COL_ALSOK_NYUKIN_DATE).Value = strAryBuffer(1)     ' 入金日
                            .Cells(intSetCnt, COL_ALSOK_REGISTER_NO).Value = strAryBuffer(6)     ' レジ№
                            .Cells(intSetCnt, COL_ALSOK_TENPO_NO).Value = strAryBuffer(5)        ' 店舗№
                            .Cells(intSetCnt, COL_ALSOK_NYUKIN_GAKU).Value = strAryBuffer(17)     ' 入金額
                            .Cells(intSetCnt, COL_ALSOK_DEPOSIT_NO).Value = strAryBuffer(0)      ' 入金機№
                            .Cells(intSetCnt, COL_ALSOK_DEPOSIT_SEQ).Value = strAryBuffer(3)     ' 入金機№連番
                            .Cells(intSetCnt, COL_ALSOK_KEIYAKUSAKI).Value = strAryBuffer(4)     ' 契約先

                            ' レジ名・店舗名・催事名の各名称を取得する
                            dataEXTY0102.PropStrDepositDt = strAryBuffer(1)
                            'dataEXTY0102.PropStrRegisterCd = strAryBuffer(5)
                            'dataEXTY0102.PropStrTenpoCd = strAryBuffer(4)
                            dataEXTY0102.PropStrRegisterCd = strAryBuffer(6)
                            dataEXTY0102.PropStrTenpoCd = strAryBuffer(5)
                            dataEXTY0102.PropStrUnionFlg = "0"                                 ' 2015.12.04 ADD h.hagiwara
                            dataEXTY0102.PropStrProc = "0"                                     ' 2015.12.18 ADD h.hagiwara

                            If GetAlsokdataCellData(dataEXTY0102) = True Then
                                .Cells(intSetCnt, COL_ALSOK_REGISTER_NAME).Value = dataEXTY0102.PropStrRegisterNm
                                .Cells(intSetCnt, COL_ALSOK_TENPO_NAME).Value = dataEXTY0102.PropStrTenpoNm

                                ' 予約情報の設定
                                Dim aryComboVal1 As New ArrayList
                                Dim aryComboTxt1 As New ArrayList
                                Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
                                comboYoyaku.Items = New String() {Nothing}
                                comboYoyaku.ItemData = New String() {""}

                                If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                                Else
                                    aryComboVal1.AddRange(comboYoyaku.Items)
                                    aryComboTxt1.AddRange(comboYoyaku.ItemData)
                                    For i As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                                        aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(0))
                                        aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1))
                                        If i = 0 Then
                                            strDspsaiji = dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1).ToString
                                        End If
                                    Next

                                    With comboYoyaku
                                        .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                                        .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                                        .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                                        '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                                    End With
                                    .Cells(intSetCnt, COL_ALSOK_SAIJI_NAME).CellType = comboYoyaku
                                    If dataEXTY0102.PropDtYoyakuData.Rows.Count = 1 Then
                                        .Cells(intSetCnt, COL_ALSOK_SAIJI_NAME).Text = strDspsaiji
                                    End If
                                End If
                            End If

                            intSetCnt += 1

                        End With
                    Else
                        ' 重複件数をカウントアップ
                        intDupCnt += 1
                        aryDupData.Add(dataEXTY0102.PropStrDepositCd + "-" + dataEXTY0102.PropStrSeq)
                    End If

                End While

            End Using

            If intSetCnt > 0 Then
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).RowSpan = intUnionCnt
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).Value = intUnionAllyen
            End If

            dataEXTY0102.PropIntDupCnt = intDupCnt
            dataEXTY0102.PropAryDupData = aryDupData

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' 重複チェック処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>ＣＳＶファイルで取り込んだデータの重複チェックを行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function DepositSeqDupCheck(ByRef dataEXTY0102 As DataEXTY0102)

        ' 開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言

        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()
            ' 登録済のALSOK入金情報表に存在するか判定
            If GetDupData(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' 同一処理内で重複情報が存在するか判定
            With dataEXTY0102
                If .PropVwAlsokData.Sheets(0).RowCount <> 0 Then
                    For i = 0 To .PropVwAlsokData.Sheets(0).RowCount - 1
                        If .PropStrDepositCd = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_NO).Value And _
                           Integer.Parse(.PropStrSeq) = Integer.Parse(.PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_SEQ).Value) Then
                            Return False
                        End If
                    Next
                End If
            End With

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 電子マネー分入金情報編集処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上に入金情報表から取得した情報を編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetAlsokdataCellData(ByRef dataEXTY0102 As DataEXTY0102, Optional ByVal GetFlg As String = "0")

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            If GetFlg <> "1" Then
                ' レジ名の取得
                If GetRegisterNm(Adapter, Cn, dataEXTY0102) = False Then
                    Return False
                End If

                ' 店舗名の取得
                If GetTenpoNm(Adapter, Cn, dataEXTY0102) = False Then
                    Return False
                End If
            End If

            ' 予約情報の取得
            If GetYoyakudata(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 既存情報読み込みチェック処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>検索条件のチェックを行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function DspInputCheck(ByRef dataEXTY0102 As DataEXTY0102)

        ' 開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 日付入力チェック
        If dataEXTY0102.PropDtpDspFrom.txtDate.Text = Nothing Then
            If dataEXTY0102.PropDtpDspTo.txtDate.Text = Nothing Then
                puErrMsg = Y0102_E006
                Return False
            End If
        Else
            If dataEXTY0102.PropDtpDspTo.txtDate.Text = Nothing Then
            Else
                If dataEXTY0102.PropDtpDspFrom.txtDate.Text > dataEXTY0102.PropDtpDspTo.txtDate.Text Then
                    puErrMsg = Y0102_E010
                    Return False
                End If
            End If

        End If

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' 既存情報読み込み処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>検索条件に合致する情報を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetVwInf(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            With dataEXTY0102
                If .PropVwAlsokData.ActiveSheet.RowCount > 0 Then
                    .PropVwAlsokData.Sheets(0).Rows.Remove(0, .PropVwAlsokData.ActiveSheet.RowCount)
                End If
                If .PropVwDegitalCashData.ActiveSheet.RowCount > 0 Then
                    .PropVwDegitalCashData.Sheets(0).Rows.Remove(0, .PropVwDegitalCashData.ActiveSheet.RowCount)
                End If

                .PropTxtFilePath.Text = Nothing
            End With

            ' ALSOK入金情報の取得
            If GetAlsokDeposit(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' 電子マネー情報の取得
            If GetDegitalCash(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' ALSOK入金情報の画面編集
            If SetAlsokDeposit(dataEXTY0102) = False Then
                Return False
            End If

            ' 電子マネー情報の画面編集
            If SetDegitalCash(dataEXTY0102) = False Then
                Return False
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 入力チェック処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上の選択および値入力のチェックを行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertInputCheck(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 変数宣言
        Dim intSetCnt As Integer = 0
        Dim intInpCnt As Integer = 0

        Try

            ' 登録時に１件以上のデータが存在しない場合エラー
            If dataEXTY0102.PropRdoProcNew.Checked = True Then
                If dataEXTY0102.PropVwAlsokData.Sheets(0).RowCount = 0 Then
                Else
                    For i = 0 To dataEXTY0102.PropVwAlsokData.Sheets(0).RowCount - 1
                        If dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(i, COL_CASH_PROP_KBN).Value = "1" Then
                        Else
                            intInpCnt += 1
                        End If
                    Next
                End If

                If dataEXTY0102.PropVwDegitalCashData.Sheets(0).RowCount = 0 Then
                Else
                    For i = 0 To dataEXTY0102.PropVwDegitalCashData.Sheets(0).RowCount - 1
                        If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_PROP_KBN).Value = "1" Then
                        Else
                            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Text = "" Then
                            Else
                                intInpCnt += 1
                            End If
                        End If
                    Next
                End If

                If intInpCnt = 0 Then
                    puErrMsg = Y0102_E008
                    Return False
                End If
            End If

            ' ＣＳＶ取り込み用の場合、催事は必須
            For i As Integer = 0 To dataEXTY0102.PropVwAlsokData.Sheets(0).RowCount - 1
                If dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_PROP_KBN).Text <> "1" Then
                    If dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_SAIJI_NAME).Text = Nothing Then
                        puErrMsg = Y0102_E007
                        Return False
                    End If
                End If
            Next

            ' 電子マネーで催事選択済の場合、店舗・金額の未入力はエラー
            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).RowCount = 0 Then
            Else
                For i = 0 To dataEXTY0102.PropVwDegitalCashData.Sheets(0).RowCount - 1
                    If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_PROP_KBN).Value = "1" Then
                    Else
                        If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Text = "" Then
                        Else
                            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_TENPO_NAME).Text = "" Then
                                puErrMsg = Y0102_E011
                                Return False
                            End If
                            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_GAKU).Text = "" Then
                                puErrMsg = Y0102_E012
                                Return False
                            End If
                        End If
                    End If
                Next
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' ＣＳＶ取り込み分入金情報編集処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上に入金情報表から取得した情報を編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetAlsokDeposit(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 変数宣言
        Dim intSetCnt As Integer = 0
        Dim intUnionStart As Integer = 0
        Dim intUnionCnt As Integer = 1
        Dim intUnionAllyen As Integer = 0
        Dim strUnionDt As String = ""

        Try

            For i As Integer = 0 To dataEXTY0102.PropDtAlsokData.Rows.Count - 1
                ' 日毎金額合算表示
                If intSetCnt > 0 Then
                    If strUnionDt = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_DT") Then
                        intUnionAllyen += Integer.Parse(dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_AMOUNT"))
                        intUnionCnt += 1
                    Else
                        dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).RowSpan = intUnionCnt
                        dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).Value = intUnionAllyen
                        strUnionDt = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_DT")
                        intUnionAllyen = Integer.Parse(dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_AMOUNT"))
                        intUnionCnt = 1
                        intUnionStart = intSetCnt
                    End If
                Else
                    strUnionDt = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_DT")
                    intUnionAllyen = Integer.Parse(dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_AMOUNT"))
                    intUnionCnt = 1
                End If

                ' スプレッド編集
                With dataEXTY0102.PropVwAlsokData.Sheets(0)
                    .Rows.Add(.RowCount, 1)
                    .Cells(intSetCnt, COL_ALSOK_NYUKIN_DATE).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_DT")     ' 入金日
                    .Cells(intSetCnt, COL_ALSOK_REGISTER_NO).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("REGISTER_CD")     ' レジ№
                    .Cells(intSetCnt, COL_ALSOK_TENPO_NO).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("TENPO_CD")         ' 店舗№
                    .Cells(intSetCnt, COL_ALSOK_NYUKIN_GAKU).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_AMOUNT")     ' 入金額
                    .Cells(intSetCnt, COL_ALSOK_DEPOSIT_NO).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_MACHINE_CD")      ' 入金機№
                    .Cells(intSetCnt, COL_ALSOK_DEPOSIT_SEQ).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("SEQ")     ' 入金機№連番
                    .Cells(intSetCnt, COL_ALSOK_KEIYAKUSAKI).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("KEIYAKUSAKI_CD")     ' 契約先
                    .Cells(intSetCnt, COL_ALSOK_REGISTER_NAME).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("REGI_NM")
                    .Cells(intSetCnt, COL_ALSOK_TENPO_NAME).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("TENPO_NM")

                    ' 催事名を取得する
                    dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropDtAlsokData.Rows(i).Item("DEPOSIT_DT")
                    dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropDtAlsokData.Rows(i).Item("REGISTER_CD")
                    dataEXTY0102.PropStrYoyakuNo = dataEXTY0102.PropDtAlsokData.Rows(i).Item("YOYAKU_NO")          ' 2015.12.04 ADD h.hagiwara
                    dataEXTY0102.PropStrUnionFlg = "1"                                                             ' 2015.12.04 ADD h.hagiwara
                    dataEXTY0102.PropStrProc = "0"                                                                 ' 2015.12.18 ADD h.hagiwara

                    If GetAlsokdataCellData(dataEXTY0102, "1") = True Then

                        ' 予約情報の設定
                        Dim aryComboVal1 As New ArrayList
                        Dim aryComboTxt1 As New ArrayList
                        Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
                        comboYoyaku.Items = New String() {Nothing}
                        comboYoyaku.ItemData = New String() {""}

                        If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                        Else
                            aryComboVal1.AddRange(comboYoyaku.Items)
                            aryComboTxt1.AddRange(comboYoyaku.ItemData)
                            For j As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                                aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(0))
                                aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(1))
                            Next

                            With comboYoyaku
                                .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                                .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                                .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                                '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                            End With
                            .Cells(intSetCnt, COL_ALSOK_SAIJI_NAME).CellType = comboYoyaku
                            .Cells(intSetCnt, COL_ALSOK_SAIJI_NAME).Value = dataEXTY0102.PropDtAlsokData.Rows(i).Item("YOYAKU_NO")
                        End If
                    End If

                    intSetCnt += 1

                End With

            Next

            If intSetCnt > 0 Then
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).RowSpan = intUnionCnt
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).Value = intUnionAllyen
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' 電子マネー分入金情報編集処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上に入金情報表から取得した情報を編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetDegitalCash(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim intSetCnt As Integer = 0
        Dim strDspsaiji As String = ""

        Try

            For i As Integer = 0 To dataEXTY0102.PropDtDegitalCashData.Rows.Count - 1
                ' スプレッド編集
                With dataEXTY0102.PropVwDegitalCashData.Sheets(0)
                    .Rows.Add(.RowCount, 1)
                    .Cells(intSetCnt, COL_CASH_NYUKIN_DATE).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("DEPOSIT_DT")     ' 入金日
                    .Cells(intSetCnt, COL_CASH_REGISTER_NO).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("REGISTER_CD")     ' レジ№
                    .Cells(intSetCnt, COL_CASH_TENPO_NO).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("TENPO_CD")         ' 店舗№
                    .Cells(intSetCnt, COL_CASH_NYUKIN_GAKU).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("DEPOSIT_AMOUNT")     ' 入金額
                    .Cells(intSetCnt, COL_CASH_DEPOSIT_NO).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("DEPOSIT_MACHINE_CD")      ' 入金機№
                    .Cells(intSetCnt, COL_CASH_DEPOSIT_SEQ).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("SEQ")     ' 入金機№連番
                    .Cells(intSetCnt, COL_CASH_REGISTER_NAME).CellType = dataEXTY0102.PropCmbReji
                    .Cells(intSetCnt, COL_CASH_TENPO_NAME).CellType = dataEXTY0102.PropCmbTenpo
                    .Cells(intSetCnt, COL_CASH_REGISTER_NAME).Text = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("REGI_NM")
                    .Cells(intSetCnt, COL_CASH_TENPO_NAME).Text = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("TENPO_NM")
                    .Cells(intSetCnt, COL_CASH_SHISETU_KBN).Text = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("SHISETU_KBN")              ' 2015.12.15 ADD h.hagiwara

                    ' 催事名を取得する
                    dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("DEPOSIT_DT")
                    dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("REGISTER_CD")
                    dataEXTY0102.PropStrYoyakuNo = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("YOYAKU_NO")    ' 2015.12.04 ADD h.hagiwara
                    dataEXTY0102.PropStrUnionFlg = "1"                                                             ' 2015.12.04 ADD h.hagiwara
                    dataEXTY0102.PropStrProc = "0"                                                                 ' 2015.12.18 ADD h.hagiwara

                    If GetAlsokdataCellData(dataEXTY0102, "1") = True Then

                        ' 予約情報の設定
                        Dim aryComboVal1 As New ArrayList
                        Dim aryComboTxt1 As New ArrayList
                        Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
                        comboYoyaku.Items = New String() {Nothing}
                        comboYoyaku.ItemData = New String() {""}

                        If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                        Else
                            aryComboVal1.AddRange(comboYoyaku.Items)
                            aryComboTxt1.AddRange(comboYoyaku.ItemData)
                            For j As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                                aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(0))
                                aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(1))
                            Next

                            With comboYoyaku
                                .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                                .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                                .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                                '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                            End With
                            .Cells(intSetCnt, COL_CASH_SAIJI_NAME).CellType = comboYoyaku
                            .Cells(intSetCnt, COL_ALSOK_SAIJI_NAME).Value = dataEXTY0102.PropDtDegitalCashData.Rows(i).Item("YOYAKU_NO")
                        End If
                    End If

                    intSetCnt += 1

                End With

            Next

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' 電子マネー用Spread上のコンボ編集
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>電子マネー用のコンボボックス（レジ・店舗）を作成する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function VwDipositCmbCre(ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' レジ情報の設定
        Dim aryComboVal1 As New ArrayList
        Dim aryComboTxt1 As New ArrayList
        Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        comboYoyaku.Items = New String() {Nothing}
        comboYoyaku.ItemData = New String() {""}
        aryComboVal1.AddRange(comboYoyaku.Items)
        aryComboTxt1.AddRange(comboYoyaku.ItemData)

        For i As Integer = 0 To dataEXTY0102.PropDtRejiMasta.Rows.Count - 1
            aryComboVal1.Add(dataEXTY0102.PropDtRejiMasta.Rows(i).Item(0))
            aryComboTxt1.Add(dataEXTY0102.PropDtRejiMasta.Rows(i).Item(1))
        Next

        Dim comboRejiM As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        With comboRejiM
            .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
            .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
            .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
        End With
        dataEXTY0102.PropCmbReji = comboRejiM

        ' 店舗情報の設定
        Dim aryComboVal2 As New ArrayList
        Dim aryComboTxt2 As New ArrayList
        aryComboVal2.AddRange(comboYoyaku.Items)
        aryComboTxt2.AddRange(comboYoyaku.ItemData)

        For i As Integer = 0 To dataEXTY0102.PropDtTenpoMasta.Rows.Count - 1
            aryComboVal2.Add(dataEXTY0102.PropDtTenpoMasta.Rows(i).Item(0))
            aryComboTxt2.Add(dataEXTY0102.PropDtTenpoMasta.Rows(i).Item(1))
        Next

        Dim comboTenpoM As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        With comboTenpoM
            .ItemData = CType(aryComboVal2.ToArray(Type.GetType("System.String")), String())
            .Items = CType(aryComboTxt2.ToArray(Type.GetType("System.String")), String())
            .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
        End With
        dataEXTY0102.PropCmbTenpo = comboTenpoM

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' 登録処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>新規取り込み時のＤＢ登録
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsUpdDB(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        If dataEXTY0102.PropRdoProcNew.Checked = True Then
            If InsertDb(dataEXTY0102) = False Then
                'エラーメッセージ表示
                puErrMsg = ""
                Return False
            End If
        ElseIf dataEXTY0102.PropRdoProcUpd.Checked = True Then
            If UpdateDb(dataEXTY0102) = False Then
                'エラーメッセージ表示
                puErrMsg = ""
                Return False
            End If
        End If

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' 登録処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>新規取り込み時のＤＢ登録
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertDb(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '新規Inc番号、システム日付取得（SELECT）
            If SelectSysDate(Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' ＣＳＶ取り込み分の登録
            If InsertAlsokdata(Tsx, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' 電子マネー分の登録
            If InsertDegitalCash(Tsx, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False

        Finally
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 更新処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既存表示時のＤＢ更新
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function UpdateDb(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '新規Inc番号、システム日付取得（SELECT）
            If SelectSysDate(Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' ＣＳＶ取り込み分の更新・削除
            If UpdateAlsokdata(Tsx, Cn, dataEXTY0102) = False Then
                Return False
            End If

            ' 電子マネー分の更新・削除
            If UpdateDegitalCash(Tsx, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False

        Finally
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 電子マネー分入金情報編集処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上に入金情報表から取得した情報を編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetCashYoyaku(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim strDspsaiji As String = ""

        Try

            ' スプレッド編集
            With dataEXTY0102.PropVwDegitalCashData.Sheets(0)
                dataEXTY0102.PropStrUnionFlg = "0"                                                             ' 2015.12.04 ADD h.hagiwara
                ' 催事名を取得する
                If GetAlsokdataCellData(dataEXTY0102, "1") = True Then

                    ' 予約情報の設定
                    Dim aryComboVal1 As New ArrayList
                    Dim aryComboTxt1 As New ArrayList
                    Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
                    comboYoyaku.Items = New String() {Nothing}
                    comboYoyaku.ItemData = New String() {""}
                    dataEXTY0102.PropStrSaiji = ""

                    If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                        dataEXTY0102.PropCmbSaiji = comboYoyaku
                    Else
                        aryComboVal1.AddRange(comboYoyaku.Items)
                        aryComboTxt1.AddRange(comboYoyaku.ItemData)
                        For j As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                            aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(0))
                            aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(j).Item(1))
                            If j = 0 Then
                                strDspsaiji = dataEXTY0102.PropDtYoyakuData.Rows(j).Item(1).ToString
                            End If
                        Next

                        With comboYoyaku
                            .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                            .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                            .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                            '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                        End With
                        dataEXTY0102.PropCmbSaiji = comboYoyaku
                        If dataEXTY0102.PropDtYoyakuData.Rows.Count = 1 Then
                            dataEXTY0102.PropStrSaiji = strDspsaiji
                        End If
                    End If
                Else
                    Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
                    comboYoyaku.Items = New String() {Nothing}
                    comboYoyaku.ItemData = New String() {""}
                    dataEXTY0102.PropCmbSaiji = comboYoyaku
                    dataEXTY0102.PropStrSaiji = ""
                End If

            End With

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' ＣＳＶ取り込み入金情報：合算金額再設定処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>スプレッド上に入金情報表から取得した情報を編集する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub SetAlsokNukin(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim intUnionStart As Integer = 0
        Dim intUnionAllyen As Integer = 0
        Dim strUnionDt As String = ""

        Try
            If dataEXTY0102.PropVwAlsokData.Sheets(0).RowCount > 0 Then
                For i = 0 To dataEXTY0102.PropVwAlsokData.Sheets(0).RowCount - 1
                    ' スプレッド編集
                    With dataEXTY0102.PropVwAlsokData.Sheets(0)
                        If i > 0 Then
                            If strUnionDt = .Cells(i, COL_ALSOK_NYUKIN_DATE).Text Then
                                If .Cells(i, COL_ALSOK_PROP_KBN).Text = "1" Then
                                Else
                                    intUnionAllyen += Integer.Parse(.Cells(i, COL_ALSOK_NYUKIN_GAKU).Value)
                                End If
                            Else
                                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(intUnionStart, COL_ALSOK_DAY_GASSAN).Value = intUnionAllyen
                                strUnionDt = .Cells(i, COL_ALSOK_NYUKIN_DATE).Text
                                If .Cells(i, COL_ALSOK_PROP_KBN).Text = "1" Then
                                    intUnionAllyen = 0
                                Else
                                    intUnionAllyen = Integer.Parse(.Cells(i, COL_ALSOK_NYUKIN_GAKU).Value)
                                End If
                                intUnionStart = i
                            End If
                        Else
                            strUnionDt = .Cells(i, COL_ALSOK_NYUKIN_DATE).Text
                            If .Cells(i, COL_ALSOK_PROP_KBN).Text = "1" Then
                            Else
                                intUnionAllyen += Integer.Parse(.Cells(i, COL_ALSOK_NYUKIN_GAKU).Value)
                            End If
                        End If
                    End With
                Next
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
        End Try

    End Sub

    ''' <summary>
    ''' レジ情報取得
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>レジ情報を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetMastaALSOKRegister(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetCmbRejiMstData(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOKレジマスタ", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
                puErrMsg = Y0102_E001
                Return False
            End If

            '取得データをデータクラスにセット
            dataEXTY0102.PropDtRejiMasta = dtmst

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 店舗情報取得
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>店舗情報を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetMastaTenpo(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetCmbTenpMstData(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "店舗マスタ", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
                puErrMsg = Y0102_E002
                Return False
            End If

            '取得データをデータクラスにセット
            dataEXTY0102.PropDtTenpoMasta = dtmst

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金データ取得取得
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ALSOK入金情報表からＣＳＶ取り込みで登録した情報を取得しデータクラスに格納
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetAlsokDeposit(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetAlsokDepositInf(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表（ＣＳＶ取り込み分）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
                'Return False
            End If

            '取得データをデータクラスにセット
            dataEXTY0102.PropDtAlsokData = dtmst

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金データ取得取得（電子マネー登録分）
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ALSOK入金情報表から電子マネー登録分で登録した情報を取得しデータクラスに格納
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetDegitalCash(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetDegitalCashInf(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表（ＣＳＶ取り込み分）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
                'Return False
            End If

            '取得データをデータクラスにセット
            dataEXTY0102.PropDtDegitalCashData = dtmst

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' サーバ日時取得
    ''' </summary>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>サーバ日時を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SelectSysDate(ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlDataAdapter
        Dim dtResult As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetSysDate(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "サーバ日時", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtResult)

            'データが取得できなかった場合、エラー
            If dtResult.Rows.Count = 0 Then
                Return False
            End If

            '取得データをデータクラスにセット
            dataEXTY0102.PropDtSysDate = dtResult.Rows(0).Item(0)

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' ALSOK入金情報表取得：重複チェック用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報が存在するか件数取得を行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetDupData(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetDupDataCnt(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表取得（重複チェック）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
            Else
                If dtmst.Rows(0).Item(0) > 0 Then
                    Return False
                End If
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 予約情報取得：ＣＳＶデータ取り込み時用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報に対応する予約情報を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetYoyakudata(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            ' 2015.12.18 UPD START↓ h.hagiwara
            '取得用SQLの作成・設定
            'If sqlEXTY0102.GeSqlYoyakudata(Adapter, Cn, dataEXTY0102) = False Then
            '    Return False
            'End If
            If dataEXTY0102.PropStrProc = "1" Then
                If sqlEXTY0102.GeSqlYoyakudata2(Adapter, Cn, dataEXTY0102) = False Then
                    Return False
                End If
            Else
                If sqlEXTY0102.GeSqlYoyakudata(Adapter, Cn, dataEXTY0102) = False Then
                    Return False
                End If
            End If
            ' 2015.12.18 UPD END↑ h.hagiwara

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得（ＣＳＶ取り込み処理）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
                dataEXTY0102.PropDtYoyakuData = Nothing
            Else
                If dtmst.Rows.Count > 0 Then
                    dataEXTY0102.PropDtYoyakuData = dtmst
                End If
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' レジ名取得：ＣＳＶデータ取り込み時用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報のレジ名を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetRegisterNm(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetSqlRegisterNm(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "レジ名称取得（ＣＳＶ取り込み処理）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
            Else
                dataEXTY0102.PropStrRegisterNm = dtmst.Rows(0).Item(0).ToString
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 店舗名取得：ＣＳＶデータ取り込み時用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＣＳＶファイルで取り込んだ情報の店舗名を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetTenpoNm(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GetSqlTenpoNm(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "店舗名称取得（ＣＳＶ取り込み処理）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count = 0 Then
            Else
                dataEXTY0102.PropStrTenpoNm = dtmst.Rows(0).Item(0).ToString
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' ＣＳＶ取り込み分 ALSOK入金情報表登録
    ''' </summary>
    ''' <param name="Tsx">[IN/OUT]NpgsqlTransaction</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks> ALSOK入金情報表にＣＳＶで取り込んだデータを登録する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function InsertAlsokdata(ByRef Tsx As NpgsqlTransaction, _
                                     ByVal Cn As NpgsqlConnection, _
                                     ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            With dataEXTY0102

                ' ＣＳＶ取り込み一覧の行数分繰り返し、登録処理を行う
                For i As Integer = 0 To .PropVwAlsokData.Sheets(0).RowCount - 1
                    ' 削除情報は登録しない
                    If .PropVwAlsokData.ActiveSheet.Cells(i, COL_ALSOK_PROP_KBN).Value = "1" Then
                    Else
                        ' データ設定
                        Dim row As DataRow = .PropDtAlsokDataSet.NewRow

                        row.Item("DepositCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_NO).Value        ' 入金機コード
                        row.Item("Seq") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_SEQ).Value             ' 連番
                        row.Item("DepositKbn") = "1"                                                                   ' 入金区分
                        row.Item("DepositDt") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_NYUKIN_DATE).Value       ' 入金日
                        row.Item("YoyakuNo") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_SAIJI_NAME).Value         ' 予約NO
                        row.Item("RegisterCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_REGISTER_NO).Value      ' レジコード
                        row.Item("TenpoCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_TENPO_NO).Value            ' 店舗コード
                        row.Item("DepositAmount") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_NYUKIN_GAKU).Value   ' 入金額
                        row.Item("KeiyakusakiCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_KEIYAKUSAKI).Value   ' 契約先コード

                        '作成した行をデータクラスにセット
                        .PropRowReg = row

                        ' ALSOK入金情報表登録SQLを作成
                        If sqlEXTY0102.SetInsertAlsokDepositInf(Cmd, Cn, dataEXTY0102) = False Then
                            Return False
                        End If

                        'ログ出力
                        commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(CSV取り込み分) 新規登録", Nothing, Cmd)

                        'SQL実行
                        Cmd.ExecuteNonQuery()

                    End If

                Next

            End With


            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 電子マネー入力分 ALSOK入金情報表登録
    ''' </summary>
    ''' <param name="Tsx">[IN/OUT]NpgsqlTransaction</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks> ALSOK入金情報表に電子マネー入力データを登録する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function InsertDegitalCash(ByRef Tsx As NpgsqlTransaction, _
                                     ByVal Cn As NpgsqlConnection, _
                                     ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            With dataEXTY0102

                ' 電子マネー一覧の行数分繰り返し、登録処理を行う
                For i As Integer = 0 To .PropVwDegitalCashData.Sheets(0).RowCount - 1
                    ' 削除情報は登録しない
                    If .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_PROP_KBN).Value = "1" Or _
                        .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Text = "" Then
                    Else
                        ' データ設定
                        Dim row As DataRow = .PropDtDegitalCashDataSet.NewRow

                        row.Item("DepositCd") = DEPOSITCASH_MACHINE_CD                                                       ' 入金機コード
                        row.Item("DepositKbn") = "2"                                                                         ' 入金区分
                        row.Item("DepositDt") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_DATE).Value        ' 入金日
                        row.Item("YoyakuNo") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Value          ' 予約NO
                        row.Item("RegisterCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_REGISTER_NO).Value       ' レジコード
                        row.Item("TenpoCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_TENPO_NO).Value             ' 店舗コード
                        row.Item("DepositAmount") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_GAKU).Value    ' 入金額
                        row.Item("KeiyakusakiCd") = ""                                                                       ' 契約先コード

                        '作成した行をデータクラスにセット
                        .PropRowReg = row

                        ' ALSOK入金情報表登録SQLを作成
                        If sqlEXTY0102.SetInsertDegitalCashInf(Cmd, Cn, dataEXTY0102) = False Then
                            Return False
                        End If

                        'ログ出力
                        commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(電子マネー入力分) 新規登録", Nothing, Cmd)

                        'SQL実行
                        Cmd.ExecuteNonQuery()

                    End If

                Next

            End With

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' ＣＳＶ取り込み分 ALSOK入金情報表 更新・削除
    ''' </summary>
    ''' <param name="Tsx">[IN/OUT]NpgsqlTransaction</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks> ALSOK入金情報表にＣＳＶで取り込んだデータを更新・削除する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function UpdateAlsokdata(ByRef Tsx As NpgsqlTransaction, _
                                     ByVal Cn As NpgsqlConnection, _
                                     ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            With dataEXTY0102

                ' ＣＳＶ取り込み一覧の行数分繰り返し、更新・削除処理を行う
                For i As Integer = 0 To .PropVwAlsokData.Sheets(0).RowCount - 1
                    ' 更新処理 削除処理の判定を行う
                    If .PropVwAlsokData.ActiveSheet.Cells(i, COL_ALSOK_PROP_KBN).Value = "1" Then
                        ' 削除処理
                        ' データ設定
                        Dim row As DataRow = .PropDtAlsokDataSet.NewRow

                        row.Item("DepositCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_NO).Value        ' 入金機コード
                        row.Item("Seq") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_SEQ).Value             ' 連番

                        '作成した行をデータクラスにセット
                        .PropRowReg = row

                        ' ALSOK入金情報表削除SQLを作成
                        If sqlEXTY0102.SetDeleteAlsokDepositInf(Cmd, Cn, dataEXTY0102) = False Then
                            Return False
                        End If

                        'ログ出力
                        commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(CSV取り込み分) 削除", Nothing, Cmd)

                        'SQL実行
                        Cmd.ExecuteNonQuery()

                    Else
                        ' データ設定
                        Dim row As DataRow = .PropDtAlsokDataSet.NewRow

                        row.Item("DepositCd") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_NO).Value        ' 入金機コード
                        row.Item("Seq") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_DEPOSIT_SEQ).Value             ' 連番
                        row.Item("YoyakuNo") = .PropVwAlsokData.Sheets(0).Cells(i, COL_ALSOK_SAIJI_NAME).Value         ' 予約NO

                        '作成した行をデータクラスにセット
                        .PropRowReg = row

                        ' ALSOK入金情報表登録SQLを作成
                        If sqlEXTY0102.SetUpdateAlsokDepositInf(Cmd, Cn, dataEXTY0102) = False Then
                            Return False
                        End If

                        'ログ出力
                        commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(CSV取り込み分) 新規登録", Nothing, Cmd)

                        'SQL実行
                        Cmd.ExecuteNonQuery()

                    End If

                Next

            End With

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 電子マネー入力分 ALSOK入金情報表 更新・削除
    ''' </summary>
    ''' <param name="Tsx">[IN/OUT]NpgsqlTransaction</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks> ALSOK入金情報表に電子マネー入力データを 更新・削除する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function UpdateDegitalCash(ByRef Tsx As NpgsqlTransaction, _
                                     ByVal Cn As NpgsqlConnection, _
                                     ByVal dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            With dataEXTY0102

                ' 電子マネー一覧の行数分繰り返し、登録処理を行う
                For i As Integer = 0 To .PropVwDegitalCashData.Sheets(0).RowCount - 1
                    ' 削除ボタンで削除対象となっているか判定
                    If .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_PROP_KBN).Value = "1" Then
                        ' 隠し列：入金機コードが未設定はＤＢ上未存在のため処理なし
                        If .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value = Nothing Then
                        Else
                            ' 隠し列：入金機コードに設定があった場合はＤＢ上にデータが存在するため削除する
                            ' データ設定
                            Dim row As DataRow = .PropDtDegitalCashDataSet.NewRow

                            row.Item("DepositCd") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value        ' 入金機コード
                            row.Item("Seq") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_SEQ).Value             ' 連番

                            '作成した行をデータクラスにセット
                            .PropRowReg = row

                            ' ALSOK入金情報表登録SQLを作成
                            If sqlEXTY0102.SetDeleteDegitalCashInf(Cmd, Cn, dataEXTY0102) = False Then
                                Return False
                            End If

                            'ログ出力
                            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(電子マネー入力分) 削除", Nothing, Cmd)

                            'SQL実行
                            Cmd.ExecuteNonQuery()
                        End If

                    Else
                        ' 削除ボタンで削除対象となってなく、催事が未選択
                        If .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Value = Nothing Then
                            ' 隠し列：入金機コードが未設定はＤＢ上未存在のため処理なし
                            If .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value = Nothing Then
                            Else
                                ' 隠し列：入金機コードに設定があった場合はＤＢ上にデータが存在するため削除する
                                ' データ設定
                                Dim row As DataRow = .PropDtDegitalCashDataSet.NewRow

                                row.Item("DepositCd") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value        ' 入金機コード
                                row.Item("Seq") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_SEQ).Value             ' 連番

                                '作成した行をデータクラスにセット
                                .PropRowReg = row

                                ' ALSOK入金情報表登録SQLを作成
                                If sqlEXTY0102.SetDeleteDegitalCashInf(Cmd, Cn, dataEXTY0102) = False Then
                                    Return False
                                End If

                                'ログ出力
                                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(電子マネー入力分) 削除", Nothing, Cmd)

                                'SQL実行
                                Cmd.ExecuteNonQuery()
                            End If
                        Else
                            ' 催事選択あり ＆ 隠し列：入金機コードが未設定 は新規登録
                            If .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value = Nothing Then
                                ' データ設定
                                Dim row As DataRow = .PropDtDegitalCashDataSet.NewRow

                                row.Item("DepositCd") = DEPOSITCASH_MACHINE_CD                                                       ' 入金機コード
                                row.Item("DepositKbn") = "2"                                                                         ' 入金区分
                                row.Item("DepositDt") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_DATE).Value        ' 入金日
                                row.Item("YoyakuNo") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Value          ' 予約NO
                                row.Item("RegisterCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_REGISTER_NO).Value       ' レジコード
                                row.Item("TenpoCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_TENPO_NO).Value             ' 店舗コード
                                row.Item("DepositAmount") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_GAKU).Value    ' 入金額
                                row.Item("KeiyakusakiCd") = ""                                                                       ' 契約先コード

                                '作成した行をデータクラスにセット
                                .PropRowReg = row

                                ' ALSOK入金情報表登録SQLを作成
                                If sqlEXTY0102.SetInsertDegitalCashInf(Cmd, Cn, dataEXTY0102) = False Then
                                    Return False
                                End If

                                'ログ出力
                                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(電子マネー入力分) 新規登録", Nothing, Cmd)

                                'SQL実行
                                Cmd.ExecuteNonQuery()

                            Else
                                ' 更新
                                ' データ設定
                                Dim row As DataRow = .PropDtDegitalCashDataSet.NewRow

                                row.Item("DepositCd") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_NO).Value        ' 入金機コード
                                row.Item("Seq") = .PropVwDegitalCashData.ActiveSheet.Cells(i, COL_CASH_DEPOSIT_SEQ).Value             ' 連番
                                row.Item("DepositDt") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_DATE).Value         ' 入金日
                                row.Item("YoyakuNo") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SAIJI_NAME).Value           ' 予約NO
                                row.Item("RegisterCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_REGISTER_NO).Value        ' レジコード
                                row.Item("TenpoCd") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_TENPO_NO).Value              ' 店舗コード
                                row.Item("DepositAmount") = .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_NYUKIN_GAKU).Value     ' 入金額
                                row.Item("KeiyakusakiCd") = ""                                                                        ' 契約先コード

                                '作成した行をデータクラスにセット
                                .PropRowReg = row

                                ' ALSOK入金情報表登録SQLを作成
                                If sqlEXTY0102.SetUpdateDegitalCashInf(Cmd, Cn, dataEXTY0102) = False Then
                                    Return False
                                End If

                                'ログ出力
                                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ALSOK入金情報表(電子マネー入力分) 新規登録", Nothing, Cmd)

                                'SQL実行
                                Cmd.ExecuteNonQuery()

                            End If
                        End If
                    End If
                Next

            End With

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' レジ番号からの施設区分取得処理
    ''' <paramref name="dataEXTY0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上選択されたレジ番号の施設区分を取得する
    ''' <para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetRejiShisetuKbn(ByRef dataEXTY0102 As DataEXTY0102)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            ' 予約情報の取得
            If GetRejiShisetuKbndata(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' レジ番号からの施設区分取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTY0102">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>スプレッド上選択されたレジ番号の施設区分を取得する
    ''' <para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetRejiShisetuKbndata(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTY0102 As DataEXTY0102) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTY0102.GeSqRejiShisetuKbn(Adapter, Cn, dataEXTY0102) = False Then
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "レジ施設区分取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合
            If dtmst.Rows.Count = 0 Then
                dataEXTY0102.PropStrShisetuKbn = Nothing
            Else
                If dtmst.Rows.Count > 0 Then
                    dataEXTY0102.PropStrShisetuKbn = dtmst.Rows(0).Item(0)
                End If
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function

End Class
