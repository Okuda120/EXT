Imports Common
Imports CommonEXT
Imports System.Configuration
Imports EXTZ

''' <summary>
''' ALSOK現金入金機データ登録Interfaceクラス
''' </summary>
''' <remarks>入金データと予約情報の紐付け 及び 電子マネー入金情報と予約情報の紐付けを行う
''' <para>作成情報：2015.08.14 h.hagiwara
''' <p>改訂情報:</p>
''' </para></remarks>

Public Class EXTY0102

    Public dataEXTY0102 As New DataEXTY0102         'データクラス
    Private logicEXTY0102 As New LogicEXTY0102      'ロジッククラス
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT    'EXT共通ロジッククラス
    Public dataEXTZ0101 As New DataEXTZ0101         'データクラス

    ''' <summary>
    ''' フォームロード時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>画面の初期設定を行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub EXTY0102_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'データクラスの初期設定を行う
            With dataEXTY0102
                .PropRdoProcNew = Me.rdoProcNew
                .PropTxtFilePath = Me.txtFilePath
                .PropBtnRef = Me.btnRef
                .PropBtnCsvInsert = Me.btnCsvInsert

                .PropRdoProcUpd = Me.rdoProcUpd
                .PropDtpDspFrom = Me.dtpDspFrom
                .PropDtpDspTo = Me.dtpDspTo
                .PropBtnDataDsp = Me.btnDataDsp

                .PropVwAlsokData = Me.vwAlsokData
                .PropVwDegitalCashData = Me.vwDegitalCashData
                .PropBtnAdd = Me.btnAdd

                .PropBtnBack = Me.btnBack
                .PropBtnUpdate = Me.btnUpdate

                .PropBtnErrDiep = Me.btnErrDisp                                          ' 2015.12.09 ADD h.hagiwara
            End With

            ' コンボ情報取得（電子マネー用：レジ＆店舗）
            If logicEXTY0102.GetCmbInf(dataEXTY0102) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Exit Sub
            End If

            'フォームの初期化
            If logicEXTY0102.InitForm(dataEXTY0102) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Exit Sub
            End If

            ' 入力制御
            ' 新規取り込み対応項目を活性化
            Me.txtFilePath.Enabled = True
            Me.btnRef.Enabled = True
            Me.btnCsvInsert.Enabled = True
            Me.btnErrDisp.Enabled = False                                              ' 2015.12.09 ADD h.hagiwara

            '既存データ表示対応項目を非活性化
            Me.dtpDspFrom.Enabled = False
            Me.dtpDspTo.Enabled = False
            Me.btnDataDsp.Enabled = False

            ' 背景色設定
            Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
            If CommonEXT.PropConfigrationFlg = "1" Then
                ' 検証機の場合には背景イメージも表示しない
                Me.BackgroundImage = Nothing
            End If

        Catch ex As Exception
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'エラーメッセージ表示
            MsgBox(EXT_E001 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    ''' <summary>
    ''' ファイル選択処理（ダイアログ表示）
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>取り込み対象のCSVファイルを選択する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnRef_Click(sender As Object, e As EventArgs) Handles btnRef.Click

        ' ダイアログ表示
        If logicEXTY0102.DspDialog(dataEXTY0102) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
        End If

    End Sub

    ''' <summary>
    ''' ＣＳＶファイル取り込み処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>指定されたＣＳＶファイルを取り込み一覧表示する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnCsvInsert_Click(sender As Object, e As EventArgs) Handles btnCsvInsert.Click

        Dim strMsg As String = ""

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        ' ファイルパス＆ＣＳＶファイルのチェック
        If logicEXTY0102.CsvInputCheck(dataEXTY0102) = False Then
            'エラーメッセージ表示
            Me.Cursor = Cursors.Default
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        Else
            ' ファイルの取り込み
            If logicEXTY0102.GetVwAlsokData(dataEXTY0102) = False Then
                'エラーメッセージ表示
                Me.Cursor = Cursors.Default
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Exit Sub
            End If
            ' 2015.12.09 UPD START↓ h.hagiwara 重複エラー情報をテキスト出力
            'If dataEXTY0102.PropIntDupCnt > 0 Then
            '    For i = 0 To dataEXTY0102.PropIntDupCnt - 1
            '        strMsg = strMsg + vbCrLf + dataEXTY0102.PropAryDupData(i).ToString
            '    Next
            '    MsgBox(String.Format(Y0102_I002, dataEXTY0102.PropIntDupCnt) + strMsg)
            'End If
            ' 2015.12.09 UPD END↑ h.hagiwara 重複エラー情報をテキスト出力
            If dataEXTY0102.PropIntDupCnt > 0 Then
                Dim textFile As System.IO.StreamWriter
                Dim OutputPath As String = ConfigurationManager.AppSettings("SystemLogPath")
                Dim OutputFile As String

                dataEXTY0102.PropStrErrFileNm = "ALSOK_INSERT_ERR_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
                OutputFile = OutputPath & dataEXTY0102.PropStrErrFileNm

                textFile = New System.IO.StreamWriter(OutputFile, True, System.Text.Encoding.Default)

                For i = 0 To dataEXTY0102.PropIntDupCnt - 1
                    textFile.WriteLine(dataEXTY0102.PropAryDupData(i).ToString)
                Next
                textFile.Close()

                MsgBox(String.Format(Y0102_I002, dataEXTY0102.PropIntDupCnt))
                Me.btnErrDisp.Enabled = True

            Else
                Me.btnErrDisp.Enabled = False
            End If
        End If

        Me.Cursor = Cursors.Default

    End Sub

    ''' <summary>
    ''' 既存情報表示処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>紐付け済の入金データを一覧表示する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnDataDsp_Click(sender As Object, e As EventArgs) Handles btnDataDsp.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Me.btnErrDisp.Enabled = False                                                ' 2015.12.09 ADD h.hagiwara

        Me.Cursor = Cursors.WaitCursor

        ' 現金機利用日のチェック
        If logicEXTY0102.DspInputCheck(dataEXTY0102) = False Then
            'エラーメッセージ表示
            Me.Cursor = Cursors.Default
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
        Else
            ' データ取得＆編集
            If logicEXTY0102.GetVwInf(dataEXTY0102) = False Then
                'エラーメッセージ表示
                Me.Cursor = Cursors.Default
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            End If
        End If

        Me.Cursor = Cursors.Default

    End Sub

    ''' <summary>
    ''' 電子マネー入力行追加処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>電子マネー情報を入力する行を追加する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

        Dim i As Integer

        ' 下Spreadに空行追加
        With dataEXTY0102
            .PropVwDegitalCashData.Sheets(0).Rows.Add(.PropVwDegitalCashData.Sheets(0).RowCount, 1)
            i = .PropVwDegitalCashData.Sheets(0).RowCount - 1

            ' 追加行のコンボリストの設定
            .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_REGISTER_NAME).CellType = .PropCmbReji
            .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_TENPO_NAME).CellType = .PropCmbTenpo
            .PropVwDegitalCashData.Sheets(0).Cells(i, COL_CASH_SHISETU_KBN).Value = "1"                                 ' 2015.12.18 ADD h.hagiwara
        End With

    End Sub

    ''' <summary>
    ''' 戻り処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>自画面を閉じてメニュー画面に戻る
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        Me.Close()

    End Sub

    ''' <summary>
    ''' 登録更新処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>一覧に設定された内容をＤＢに反映する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 入力チェック
        If logicEXTY0102.InsertInputCheck(dataEXTY0102) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        If MsgBox(Y0102_I003, vbYesNo) = MsgBoxResult.No Then
            Exit Sub
        End If
        ' 登録・変更処理
        If logicEXTY0102.InsUpdDB(dataEXTY0102) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        ' 完了メッセージを表示
        MsgBox(Y0102_I001, MsgBoxStyle.Information)

        ' 表示
        'フォームの初期化
        If logicEXTY0102.InitForm(dataEXTY0102) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        ' 入力制御
        ' 新規取り込み対応項目を活性化
        Me.rdoProcNew.Checked = True
        Me.txtFilePath.Enabled = True
        Me.btnRef.Enabled = True
        Me.btnCsvInsert.Enabled = True
        Me.btnErrDisp.Enabled = False                                              ' 2015.12.09 ADD h.hagiwara

        ' 既存データ表示対応項目を非活性化
        Me.dtpDspFrom.Enabled = False
        Me.dtpDspFrom.txtDate.Text = Nothing
        Me.dtpDspTo.Enabled = False
        Me.dtpDspTo.txtDate.Text = Nothing
        Me.btnDataDsp.Enabled = False

    End Sub

    ''' <summary>
    ''' 処理区分による入力制御処理１
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>新規取り込み選択時、日付範囲設定＆ボタンを非活性化する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub rdoProcNew_CheckedChanged(sender As Object, e As EventArgs) Handles rdoProcNew.CheckedChanged

        'フォームの初期化
        If logicEXTY0102.InitForm(dataEXTY0102) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        ' 新規取り込み対応項目を活性化
        Me.txtFilePath.Enabled = True
        Me.btnRef.Enabled = True
        Me.btnCsvInsert.Enabled = True
        Me.btnErrDisp.Enabled = False                                              ' 2015.12.09 ADD h.hagiwara

        ' 既存データ表示対応項目を非活性化
        Me.dtpDspFrom.Enabled = False
        Me.dtpDspTo.Enabled = False
        Me.btnDataDsp.Enabled = False

    End Sub

    ''' <summary>
    ''' 処理区分による入力制御処理２
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>既存データ表示選択時、日付範囲設定＆ボタンを活性化する ＆ 新規取り込みの各ボタンを非活性化する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub rdoProcUpd_CheckedChanged(sender As Object, e As EventArgs) Handles rdoProcUpd.CheckedChanged

        'フォームの初期化
        If logicEXTY0102.InitForm(dataEXTY0102) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        ' 新規取り込み対応項目を非活性化
        Me.txtFilePath.Enabled = False
        Me.btnRef.Enabled = False
        Me.btnCsvInsert.Enabled = False
        Me.btnErrDisp.Enabled = False                                              ' 2015.12.09 ADD h.hagiwara

        ' 既存データ表示対応項目を活性化
        Me.dtpDspFrom.Enabled = True
        Me.dtpDspTo.Enabled = True
        Me.btnDataDsp.Enabled = True

    End Sub

    ' 2015.12.04 UPD START↓ h.hagiwara
    '' ''' <summary>
    '' ''' スプレッド上の選択ボタンクリック時の処理
    '' ''' </summary>
    '' ''' <param name="sender">[IN]</param>
    '' ''' <param name="e">[IN]</param>
    '' ''' <remarks>スプレッドＡＬＳＯＫ入金情報の選択ボタンクリック時の画面表示＆戻り値取得
    '' ''' <para>作成情報：2015.08.14 h.hagiwara
    '' ''' <p>改訂情報 : </p>
    '' ''' </para></remarks>
    ''Private Sub vwAlsokData_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwAlsokData.CellClick

    ''    ' 選択ボタンクリック時、予約一覧をＨＥＬＰ表示する
    ''    If e.Column = COL_ALSOK_SELECT_BTN Then

    ''        ' DBレプリケーション
    ''        If commonLogicEXT.CheckDBCondition() = False Then
    ''            'メッセージを出力 
    ''            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
    ''            Exit Sub
    ''        End If

    ''        '「予約検索一覧」インスタンス作成
    ''        Dim frm As New EXTZ0101
    ''        Dim strYoyakuno As String = ""
    ''        Dim strSaijiNm As String = ""

    ''        frm.dataEXTZ0101.PropStrSelectMode = "1"
    ''        frm.ShowDialog()

    ''        ' DBレプリケーション
    ''        If commonLogicEXT.CheckDBCondition() = False Then
    ''            'メッセージを出力 
    ''            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
    ''            Exit Sub
    ''        End If

    ''        ''グループ検索画面を表示し、戻り値としてデータテーブルを取得
    ''        strYoyakuno = frm.dataEXTZ0101.PropStrRtnYoyakuNo
    ''        strSaijiNm = frm.dataEXTZ0101.PropStrRtnSaijiNm

    ''        '予約番号＆名称を更新
    ''        If strYoyakuno <> "" Then
    ''            dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_REGISTER_NO).Value
    ''            dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_NYUKIN_DATE).Value
    ''            Dim aryComboVal1 As New ArrayList
    ''            Dim aryComboTxt1 As New ArrayList
    ''            Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

    ''            comboYoyaku.Items = New String() {""}
    ''            comboYoyaku.ItemData = New String() {Nothing}
    ''            aryComboVal1.AddRange(comboYoyaku.Items)
    ''            aryComboTxt1.AddRange(comboYoyaku.ItemData)
    ''            dataEXTY0102.PropStrUnionFlg = "0"                                 ' 2015.12.04 ADD h.hagiwara

    ''            If logicEXTY0102.GetAlsokdataCellData(dataEXTY0102, "1") = True Then
    ''                ' 取得した予約番号＆名称をコンボに追加
    ''                If dataEXTY0102.PropDtYoyakuData Is Nothing Then
    ''                Else
    ''                    aryComboVal1.AddRange(comboYoyaku.Items)
    ''                    aryComboTxt1.AddRange(comboYoyaku.ItemData)
    ''                    For i As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
    ''                        aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(0))
    ''                        aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1))
    ''                    Next
    ''                End If
    ''            End If
    ''            aryComboVal1.Add(strYoyakuno)
    ''            aryComboTxt1.Add(strSaijiNm)

    ''            With comboYoyaku
    ''                .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
    ''                .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
    ''                .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    ''                '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
    ''            End With

    ''            ' 追加行のコンボリストの設定
    ''            dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_SAIJI_NAME).CellType = comboYoyaku
    ''            ' 取得した名称を表示
    ''            dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_SAIJI_NAME).Text = strSaijiNm
    ''        End If

    ''    End If

    ''    ' 削除ボタンクリック時、スプレッド列に値を設定し非表示にする
    ''    If e.Column = COL_ALSOK_DELETE_BTN Then

    ''        ' 削除区分を設定
    ''        dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_PROP_KBN).Text = "1"

    ''        ' 選択行を非表示
    ''        dataEXTY0102.PropVwAlsokData.Sheets(0).Rows(e.Row).Visible = False

    ''        ' 日毎合算額の再設定
    ''        logicEXTY0102.SetAlsokNukin(dataEXTY0102)
    ''    End If

    ''End Sub
    ''' <summary>
    ''' スプレッド上の選択ボタンクリック時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>スプレッドＡＬＳＯＫ入金情報の選択ボタンクリック時の画面表示＆戻り値取得
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub vwAlsokData_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwAlsokData.ButtonClicked

        ' 選択ボタンクリック時、予約一覧をＨＥＬＰ表示する
        If e.Column = COL_ALSOK_SELECT_BTN Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「予約検索一覧」インスタンス作成
            Dim frm As New EXTZ0101
            Dim strYoyakuno As String = ""
            Dim strSaijiNm As String = ""

            frm.dataEXTZ0101.PropStrSelectMode = "1"
            frm.dataEXTZ0101.PropStrSelectShisetu = "1"
            frm.ShowDialog()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            ''グループ検索画面を表示し、戻り値としてデータテーブルを取得
            strYoyakuno = frm.dataEXTZ0101.PropStrRtnYoyakuNo
            strSaijiNm = frm.dataEXTZ0101.PropStrRtnSaijiNm

            '予約番号＆名称を更新
            If strYoyakuno <> "" Then
                dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_REGISTER_NO).Value
                dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_NYUKIN_DATE).Value
                Dim aryComboVal1 As New ArrayList
                Dim aryComboTxt1 As New ArrayList
                Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                comboYoyaku.Items = New String() {Nothing}
                comboYoyaku.ItemData = New String() {""}
                dataEXTY0102.PropStrUnionFlg = "0"                                 ' 2015.12.04 ADD h.hagiwara
                dataEXTY0102.PropStrProc = "0"                                     ' 2015.12.18 ADD h.hagiwara

                If logicEXTY0102.GetAlsokdataCellData(dataEXTY0102, "1") = True Then
                    ' 取得した予約番号＆名称をコンボに追加
                    If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                    Else
                        aryComboVal1.AddRange(comboYoyaku.Items)
                        aryComboTxt1.AddRange(comboYoyaku.ItemData)
                        For i As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                            aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(0))
                            aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1))
                        Next
                    End If
                End If
                aryComboVal1.Add(strYoyakuno)
                aryComboTxt1.Add(strSaijiNm)

                With comboYoyaku
                    .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                    .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                    .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                    '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                End With

                ' 追加行のコンボリストの設定
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_SAIJI_NAME).CellType = comboYoyaku
                ' 取得した名称を表示
                dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_SAIJI_NAME).Text = strSaijiNm
            End If

        End If

        ' 削除ボタンクリック時、スプレッド列に値を設定し非表示にする
        If e.Column = COL_ALSOK_DELETE_BTN Then

            ' 削除区分を設定
            dataEXTY0102.PropVwAlsokData.Sheets(0).Cells(e.Row, COL_ALSOK_PROP_KBN).Text = "1"

            ' 選択行を非表示
            dataEXTY0102.PropVwAlsokData.Sheets(0).Rows(e.Row).Visible = False

            ' 日毎合算額の再設定
            logicEXTY0102.SetAlsokNukin(dataEXTY0102)
        End If

    End Sub

    ' 2015.12.04 UPD END↑ h.hagiwara

    ' 2015.12.04 UPD START↓ h.hagiwara
    '' ''' <summary>
    '' ''' スプレッド上の選択ボタンクリック時の処理
    '' ''' </summary>
    '' ''' <param name="sender">[IN]</param>
    '' ''' <param name="e">[IN]</param>
    '' ''' <remarks>スプレッド電子マネー入金情報の選択ボタンクリック時の画面表示＆戻り値取得
    '' ''' <para>作成情報：2015.08.14 h.hagiwara
    '' ''' <p>改訂情報 : </p>
    '' ''' </para></remarks>
    ''Private Sub vwDegitalCashData_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwDegitalCashData.CellClick

    ''    ' 選択ボタンクリック時、予約一覧をＨＥＬＰ表示する
    ''    If e.Column = COL_CASH_SELECT_BTN Then
    ''        '「予約検索一覧」インスタンス作成
    ''        Dim frm As New EXTZ0101
    ''        Dim strYoyakuno As String = ""
    ''        Dim strSaijiNm As String = ""

    ''        ' DBレプリケーション
    ''        If commonLogicEXT.CheckDBCondition() = False Then
    ''            'メッセージを出力 
    ''            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
    ''            Exit Sub
    ''        End If

    ''        frm.dataEXTZ0101.PropStrSelectMode = "1"
    ''        frm.ShowDialog()

    ''        ' DBレプリケーション
    ''        If commonLogicEXT.CheckDBCondition() = False Then
    ''            'メッセージを出力 
    ''            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
    ''            Exit Sub
    ''        End If

    ''        ''グループ検索画面を表示し、戻り値としてデータテーブルを取得
    ''        strYoyakuno = frm.dataEXTZ0101.PropStrRtnYoyakuNo
    ''        strSaijiNm = frm.dataEXTZ0101.PropStrRtnSaijiNm

    ''        '予約番号＆名称を更新
    ''        If strYoyakuno <> "" Then
    ''            dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Value
    ''            dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Value
    ''            Dim aryComboVal1 As New ArrayList
    ''            Dim aryComboTxt1 As New ArrayList
    ''            Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

    ''            comboYoyaku.Items = New String() {""}
    ''            comboYoyaku.ItemData = New String() {Nothing}
    ''            aryComboVal1.AddRange(comboYoyaku.Items)
    ''            aryComboTxt1.AddRange(comboYoyaku.ItemData)
    ''            dataEXTY0102.PropStrUnionFlg = "0"                                 ' 2015.12.04 ADD h.hagiwara

    ''            If logicEXTY0102.GetAlsokdataCellData(dataEXTY0102, "1") = True Then
    ''                ' 取得した予約番号＆名称をコンボに追加

    ''                If dataEXTY0102.PropDtYoyakuData Is Nothing Then
    ''                Else
    ''                    aryComboVal1.AddRange(comboYoyaku.Items)
    ''                    aryComboTxt1.AddRange(comboYoyaku.ItemData)
    ''                    For i As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
    ''                        aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(0))
    ''                        aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1))
    ''                    Next
    ''                End If
    ''            End If
    ''            aryComboVal1.Add(strYoyakuno)
    ''            aryComboTxt1.Add(strSaijiNm)

    ''            With comboYoyaku
    ''                .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
    ''                .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
    ''                .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    ''                '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
    ''            End With

    ''            ' 追加行のコンボリストの設定
    ''            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = comboYoyaku
    ''            ' 取得した名称を表示
    ''            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = strSaijiNm
    ''        End If

    ''    End If

    ''    ' 削除ボタンクリック時、スプレッド列に値を設定し非表示にする
    ''    If e.Column = COL_CASH_DELETE_BTN Then

    ''        ' 削除区分を設定
    ''        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_PROP_KBN).Text = "1"

    ''        ' 選択行を非表示
    ''        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Rows(e.Row).Visible = False

    ''    End If

    ''End Sub
    ''' <summary>
    ''' スプレッド上の選択ボタンクリック時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>スプレッド電子マネー入金情報の選択ボタンクリック時の画面表示＆戻り値取得
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub vwDegitalCashData_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwDegitalCashData.ButtonClicked
        ' 選択ボタンクリック時、予約一覧をＨＥＬＰ表示する
        If e.Column = COL_CASH_SELECT_BTN Then
            '「予約検索一覧」インスタンス作成
            Dim frm As New EXTZ0101
            Dim strYoyakuno As String = ""
            Dim strSaijiNm As String = ""

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            frm.dataEXTZ0101.PropStrSelectMode = "1"
            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Value = "2" Then
                frm.dataEXTZ0101.PropStrSelectShisetu = "2"
            Else
                frm.dataEXTZ0101.PropStrSelectShisetu = "1"
            End If
            frm.ShowDialog()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            ''グループ検索画面を表示し、戻り値としてデータテーブルを取得
            strYoyakuno = frm.dataEXTZ0101.PropStrRtnYoyakuNo
            strSaijiNm = frm.dataEXTZ0101.PropStrRtnSaijiNm

            '予約番号＆名称を更新
            If strYoyakuno <> "" Then
                dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Value
                dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Value
                Dim aryComboVal1 As New ArrayList
                Dim aryComboTxt1 As New ArrayList
                Dim comboYoyaku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                comboYoyaku.Items = New String() {Nothing}
                comboYoyaku.ItemData = New String() {""}
                dataEXTY0102.PropStrUnionFlg = "0"                                 ' 2015.12.04 ADD h.hagiwara
                dataEXTY0102.PropStrProc = "0"                                     ' 2015.12.18 ADD h.hagiwara

                If logicEXTY0102.GetAlsokdataCellData(dataEXTY0102, "1") = True Then
                    ' 取得した予約番号＆名称をコンボに追加

                    If dataEXTY0102.PropDtYoyakuData Is Nothing Then
                    Else
                        aryComboVal1.AddRange(comboYoyaku.Items)
                        aryComboTxt1.AddRange(comboYoyaku.ItemData)
                        For i As Integer = 0 To dataEXTY0102.PropDtYoyakuData.Rows.Count - 1
                            aryComboVal1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(0))
                            aryComboTxt1.Add(dataEXTY0102.PropDtYoyakuData.Rows(i).Item(1))
                        Next
                    End If
                End If
                aryComboVal1.Add(strYoyakuno)
                aryComboTxt1.Add(strSaijiNm)

                With comboYoyaku
                    .ItemData = CType(aryComboVal1.ToArray(Type.GetType("System.String")), String())
                    .Items = CType(aryComboTxt1.ToArray(Type.GetType("System.String")), String())
                    .EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                    '.Editable = True                           ' 2015.12.03 DEL h.hagiwara
                End With

                ' 追加行のコンボリストの設定
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = comboYoyaku
                ' 取得した名称を表示
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = strSaijiNm
            End If

        End If

        ' 削除ボタンクリック時、スプレッド列に値を設定し非表示にする
        If e.Column = COL_CASH_DELETE_BTN Then

            ' 削除区分を設定
            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_PROP_KBN).Text = "1"

            ' 選択行を非表示
            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Rows(e.Row).Visible = False

        End If

    End Sub
    ' 2015.12.04 UPD END↑ h.hagiwara

    ''' <summary>
    ''' スプレッド上の値変更時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>スプレッド電子マネー入金情報の値変更時の処理
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub vwDegitalCashData_Changed(sender As Object, e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwDegitalCashData.Change

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' レジ名変更時のレジ№表示変更
        If e.Column = COL_CASH_REGISTER_NAME Then
            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NAME).Text = Nothing Then
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Text = ""
            Else
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Text = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NAME).Value
            End If
        End If

        ' 店舗名変更時の店舗№表示変更
        If e.Column = COL_CASH_TENPO_NAME Then
            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_TENPO_NAME).Text = Nothing Then
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_TENPO_NO).Text = ""
            Else
                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_TENPO_NO).Text = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_TENPO_NAME).Value
            End If
        End If

        ' 入金日入力時＆レジ入力時に予約情報を取得
        If e.Column = COL_CASH_NYUKIN_DATE Then
            ' 2015.12.18 DEL START↓ h.hagiwara
            'dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text
            'dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Text
            ' 2015.12.18 DEL END↑ h.hagiwara
            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text = Nothing Then
                dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text                          ' 2015.12.18 ADD h.hagiwara
                dataEXTY0102.PropStrShisetuKbn = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Text                         ' 2015.12.18 ADD h.hagiwara
                dataEXTY0102.PropStrProc = "1"                                                                                                                ' 2015.12.18 ADD h.hagiwara
                ' 予約情報のクリア
                If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                End If
            Else
                ' 2015.12.18 UPD START↓ h.hagiwara 入金日が変更された場合に予約情報をクリア＆情報取得
                'If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Text = Nothing Then
                '    ' 予約情報のクリア
                '    If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                '        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                '        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                '    End If
                'Else
                '    ' 予約情報のクリア
                '    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = Nothing
                '    ' 催事名を取得する
                '    If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                '        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                '        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                '    End If
                'End If
                If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text = dataEXTY0102.PropStrDepositDt Then
                Else
                    dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text
                    dataEXTY0102.PropStrShisetuKbn = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Text
                    dataEXTY0102.PropStrProc = "1"
                    ' 予約情報のクリア
                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = Nothing
                    ' 催事名を取得する
                    If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                    End If
                End If
                ' 2015.12.18 UPD END↑ h.hagiwara 入金日が変更された場合に予約情報をクリア＆情報取得
            End If
        Else
            If e.Column = COL_CASH_REGISTER_NAME Then
                dataEXTY0102.PropStrDepositDt = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text
                dataEXTY0102.PropStrRegisterCd = dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NO).Text
                dataEXTY0102.PropStrProc = "0"                                                                                                                       ' 2015.12.18 ADD h.hagiwara
                If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_REGISTER_NAME).Text = Nothing Then
                    ' 予約情報のクリア
                    If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                    End If
                Else
                    If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_NYUKIN_DATE).Text = Nothing Then
                        ' 予約情報のクリア
                        If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                        End If
                    Else
                        ' 2015.12.09 ADD START↓ h.hagiwara    レジ変更した場合に同じ施設区分は予約情報の再取得を行わない
                        If logicEXTY0102.GetRejiShisetuKbn(dataEXTY0102) = True Then
                            If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Text = Nothing Then                                     ' 2015.12.09 ADD h.hagiwara
                                ' 予約情報のクリア
                                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = Nothing
                                ' 催事名を取得する
                                If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                                End If
                            Else                                                                                                                                       ' 2015.12.09 ADD h.hagiara 
                                If dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Text <> dataEXTY0102.PropStrShisetuKbn Then         ' 2015.12.09 ADD h.hagiwara
                                    ' 予約情報のクリア
                                    dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = Nothing
                                    ' 催事名を取得する
                                    If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                                        dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                                    End If
                                End If                                                                                                                                ' 2015.12.09 ADD h.hagiara 
                            End If                                                                                                                                    ' 2015.12.09 ADD h.hagiara 
                            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SHISETU_KBN).Text = dataEXTY0102.PropStrShisetuKbn
                        Else
                            ' 予約情報のクリア
                            dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = Nothing
                            ' 催事名を取得する
                            If logicEXTY0102.GetCashYoyaku(dataEXTY0102) = True Then
                                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).CellType = dataEXTY0102.PropCmbSaiji
                                dataEXTY0102.PropVwDegitalCashData.Sheets(0).Cells(e.Row, COL_CASH_SAIJI_NAME).Text = dataEXTY0102.PropStrSaiji
                            End If
                        End If
                        ' 2015.12.09 ADD END↑ h.hagiwara
                    End If
                End If

            End If
        End If

    End Sub

    ''' <summary>
    ''' エラー内容表示処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>取り込み時に重複エラーとなった情報を表示する
    ''' <para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnErrDisp_Click(sender As Object, e As EventArgs) Handles btnErrDisp.Click

        Dim OutputPath As String = ConfigurationManager.AppSettings("SystemLogPath")
        Dim OutputFile As String

        OutputFile = OutputPath & dataEXTY0102.PropStrErrFileNm
        System.Diagnostics.Process.Start("notepad.exe", OutputFile)

    End Sub

End Class