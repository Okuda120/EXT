Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win
Imports FarPoint.Win.Spread.CellType
Imports EXTM

''' <summary>
''' EXTC0102
''' </summary>
''' <remarks>仮予約（シアター）画面
''' <para>作成情報：2015/09/20 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTC0102

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogicEXTC As New CommonLogicEXTC      '共通クラス
    Private mailRegexUtilities As New MailRegexUtilities
    '変数宣言
    Public dataCommon As New CommonDataEXT      '共通データクラス
    Public dataEXTC0102 As New DataEXTC0102     'データクラス
    Public logicEXTC0102 As New LogicEXTC0102   'ロジッククラス
    Public logicEXTZ0202 As New LogicEXTZ0202   'ロジッククラス
    Public logicEXTC0104 As New LogicEXTC0104   'Cancel Logic
    Private strCmbLoadFlg As String             'コンボボックス初期制御フラグ            ' 2016.08.08 ADD e.watanabe

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTC0102_Load(sender As Object, e As EventArgs) Handles Me.Load

        strCmbLoadFlg = "0"                   '2016.08.08 ADD e.watanabe

        ''画面項目初期化 2015.11.19 DEL h.hagiwara
        'If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
        '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        'End If
        dataEXTC0102.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        'コンボボックスの内容設定
        commonLogicEXT.TodohukenLst(Me.cmbRiyoTodo)
        '予約NO無い場合==================================================
        If String.IsNullOrEmpty(dataEXTC0102.PropStrYoyakuNo) = True And String.IsNullOrEmpty(dataEXTC0102.PropStrCancelNo) = True Then
            'ヘッダー
            Me.lblYoyakuNo.Text = String.Empty
            Me.dtpKariuke.txtDate.Text = String.Empty

            dataEXTC0102.PropStrYoyakuSts = CommonDeclareEXT.YOYAKU_STS_KARI
            Me.lblStatus.Text = commonLogicEXTC.GetYoyakuSts(dataEXTC0102.PropStrYoyakuSts)
            Me.lblStatus.BackColor = commonLogicEXTC.GetYoyakuStsColor(dataEXTC0102.PropStrYoyakuSts)

            Me.lblAddUserCd.Text = String.Empty
            Me.lblAddUserNm.Text = String.Empty
            Me.lblAddUserDate.Text = String.Empty
            Me.lblUpUserCd.Text = String.Empty
            Me.lblUpUserNm.Text = String.Empty
            Me.lblUpUserDate.Text = String.Empty
            '催事
            Me.txtShutuen.Text = String.Empty
            '貸し出し種別
            Me.chkKashi.Checked = True
            Me.rdoKashi1.Checked = True
            '利用スタジオ
            'Me.chkStudio.Checked = True                                            ' 2015.12.09 DEL h.hagiwara
            'Me.rdoStudio1.Checked = True                                           ' 2015.12.09 UPD h.hagiwara
            Me.rdoStudio1.Checked = False                                           ' 2015.12.09 UPD h.hagiwara
            Me.rdoStudio2.Checked = False                                           ' 2015.12.09 ADD h.hagiwara
            Me.rdoStudio3.Checked = False                                           ' 2015.12.09 ADD h.hagiwara
            '音響オペレーター
            Me.chkOpe.Checked = True
            Me.rdoOpe1.Checked = True
            '利用者情報
            Me.lblRiyoshaCd.Text = String.Empty
            Me.lblRiyoshaLvl.Text = String.Empty
            Me.txtRiyoshaNmKana.Text = String.Empty
            Me.txtRiyoshaNm.Text = String.Empty
            Me.txtDaihyoNm.Text = String.Empty
            Me.txtSekininBushoNm.Text = String.Empty
            Me.cmbSekininNm.Text = String.Empty
            Me.txtSekininMail.Text = String.Empty
            Me.lblExasAiteNm.Text = String.Empty
            Me.lblExasAite.Text = String.Empty
            Me.txtRiyoPost1.Text = String.Empty
            Me.txtRiyoPost2.Text = String.Empty
            Me.cmbRiyoTodo.Text = String.Empty
            Me.txtRiyoShiku.Text = String.Empty
            Me.txtRiyoBan.Text = String.Empty
            Me.txtRiyoBuild.Text = String.Empty
            Me.txtRiyoTel1.Text = String.Empty
            Me.txtRiyoTel2.Text = String.Empty
            Me.txtRiyoTel3.Text = String.Empty
            Me.txtRiyoNaisen.Text = String.Empty
            Me.txtRiyoMobileTel1.Text = String.Empty
            'Me.txtRiyoMobileTel2.Text = String.Empty ' 2016.11.2 m.hayabuchi DEL 課題No.59
            'Me.txtRiyoMobileTel3.Text = String.Empty ' 2016.11.2 m.hayabuchi DEL 課題No.59
            Me.txtRiyoFax1.Text = String.Empty
            Me.txtRiyoFax2.Text = String.Empty
            Me.txtRiyoFax3.Text = String.Empty
            '事務処理欄
            Me.chkSendKbn.Checked = True
            Me.rdoSendKbn1.Checked = True
            '特記事項
            Me.txtBiko.Text = String.Empty
            '予約NoないのでDISABLE
            Me.btnOfficial.Enabled = False
            Me.btnCancel.Enabled = False

            '活性/非活性
            Me.pnlKashi.SuspendLayout()
            'Me.pnlStudio.SuspendLayout()                                ' 2015.12.09 DEL h.hagiwara
            Me.pnlOpe.SuspendLayout()

            '初期利用日の作成
            Dim dataRiyobi As New CommonDataRiyobi
            Dim lstRiyobi As New ArrayList
            For Each riyobi As Date In dataEXTC0102.PropAryStrRiyoDate
                With dataRiyobi
                    .PropStrYoyakuDt = riyobi.ToString(CommonDeclareEXT.FMT_DATE) '日付
                    .PropStrYoyakuDtDisp = riyobi.ToString(CommonDeclareEXT.FMT_DATE_DISP) '日付
                    .PropStrShisetuKbn = SHISETU_KBN_STUDIO  '施設区分(スタジオ)
                    .PropStrStudioKbn = STUDIO_MITEI   'スタジオ区分(スタジオ)
                    .PropStrMiteiFlg = "1" '未定
                    .PropStrRegistFlg = "0" 'DB登録済データ：未登録
                    .PropIntTanka = Nothing
                    .PropDblBairitu = Nothing
                    .PropIntSu = Nothing
                    .PropIntRiyoKin = Nothing
                    .PropStrRiyoKeitai = RIYOKEITAI_LOCKOUT
                End With
                lstRiyobi.Add(dataRiyobi)
                dataRiyobi = New CommonDataRiyobi
            Next
            SetSpreadDataRiyobi(lstRiyobi)
            dataEXTC0102.PropListRiyobi = lstRiyobi

            '設定処理
            SetEnabledSettings()
        Else
            If String.IsNullOrEmpty(dataEXTC0102.PropStrYoyakuNo) = False Then
                '予約NOがある場合==================================================
                'メインデータ取得
                Me.lblYoyakuNo.Text = dataEXTC0102.PropStrYoyakuNo
                logicEXTC0102.GetYoyakuData(dataEXTC0102)
                If logicEXTC0102.GetRiyobiData(dataEXTC0102) = True Then
                    '利用日データ設定
                    SetSpreadDataRiyobi(dataEXTC0102.PropListRiyobi)
                End If
            ElseIf String.IsNullOrEmpty(dataEXTC0102.PropStrCancelNo) = False Then
                'キャンセルNOがある場合==================================================
                '利用日データ設定
                SetSpreadDataRiyobi(dataEXTC0102.PropListRiyobi)
            End If
            'ヘッダ
            Me.dtpKariuke.txtDate.Text = dataEXTC0102.PropStrKariukeDt
            'Me.btnKariukeCal
            Me.lblStatus.Text = commonLogicEXTC.GetYoyakuSts(dataEXTC0102.PropStrYoyakuSts)
            Me.lblStatus.BackColor = commonLogicEXTC.GetYoyakuStsColor(dataEXTC0102.PropStrYoyakuSts)

            Me.lblAddUserCd.Text = dataEXTC0102.PropStrAddUserCd
            Me.lblAddUserNm.Text = dataEXTC0102.PropStrAddUserNm
            Me.lblAddUserDate.Text = dataEXTC0102.PropStrAddDt
            Me.lblUpUserCd.Text = dataEXTC0102.PropStrUpUserCd
            Me.lblUpUserNm.Text = dataEXTC0102.PropStrUpUserNm
            Me.lblUpUserDate.Text = dataEXTC0102.PropStrUpDt
            '催事
            Me.txtShutuen.Text = dataEXTC0102.PropStrShutsuenNm
            '貸し出し種別
            If dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI Then
                Me.chkKashi.Checked = True
                Me.rdoKashi1.Checked = True
            ElseIf dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
                Me.chkKashi.Checked = False
                Me.rdoKashi1.Checked = True
            ElseIf dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
                Me.chkKashi.Checked = False
                Me.rdoKashi2.Checked = True
            End If
            '利用スタジオ
            If dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_MITEI Then
                'Me.rdoStudio1.Checked = True                                                          ' 2015.12.09 DEL h.hagiwara
                'Me.chkStudio.Checked = True                                                           ' 2015.12.09 DEL h.hagiwara
                Me.rdoStudio1.Checked = False                                                          ' 2015.12.09 ADD h.hagiwara
                Me.rdoStudio2.Checked = False                                                          ' 2015.12.09 ADD h.hagiwara
                Me.rdoStudio3.Checked = False                                                          ' 2015.12.09 ADD h.hagiwara
            ElseIf dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_201 Then
                'Me.chkStudio.Checked = False                                                          ' 2015.12.09 DEL h.hagiwara
                Me.rdoStudio1.Checked = True
            ElseIf dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_202 Then
                'Me.chkStudio.Checked = False                                                          ' 2015.12.09 DEL h.hagiwara
                Me.rdoStudio2.Checked = True
            ElseIf dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_HOUSE_LOCK Then
                'Me.chkStudio.Checked = False                                                          ' 2015.12.09 DEL h.hagiwara
                Me.rdoStudio3.Checked = True
            End If
            '音響オペレーター
            If dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_MITEI Then
                Me.chkOpe.Checked = True
                Me.rdoOpe1.Checked = True
            ElseIf dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_ARI Then
                Me.chkOpe.Checked = False
                Me.rdoOpe1.Checked = True
            ElseIf dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_NASHI Then
                Me.chkOpe.Checked = False
                Me.rdoOpe2.Checked = True
            End If
            '利用者情報
            Me.lblRiyoshaCd.Text = dataEXTC0102.PropStrRiyoshaCd
            Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTC0102.PropStrRiyoLvl)
            Me.txtRiyoshaNmKana.Text = dataEXTC0102.PropStrRiyoKana
            Me.txtRiyoshaNm.Text = dataEXTC0102.PropStrRiyoNm
            Me.txtDaihyoNm.Text = dataEXTC0102.PropStrDaihyoNm
            Me.txtSekininBushoNm.Text = dataEXTC0102.PropStrSekininBushoNm
            '責任者名コンボボックス作成
            cmbSekininNm.Items.Clear()                            'コンボボックス初期化  2016.11.04 m.hayabuchi ADD
            If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then
                Dim list As ArrayList = logicEXTC0102.GetSekininshaList(dataEXTC0102.PropStrRiyoshaCd)
                For i = 0 To list.Count - 1
                    cmbSekininNm.Items.Add(list(i))
                Next
            End If
            Me.cmbSekininNm.Text = dataEXTC0102.PropStrSekininNm
            Me.txtSekininMail.Text = dataEXTC0102.PropStrSekininMail
            Me.lblExasAiteNm.Text = dataEXTC0102.PropStrAiteNm
            Me.lblExasAite.Text = dataEXTC0102.PropStrAiteCd
            Me.txtRiyoPost1.Text = dataEXTC0102.PropStrRiyoYubin1
            Me.txtRiyoPost2.Text = dataEXTC0102.PropStrRiyoYubin2
            Me.cmbRiyoTodo.Text = dataEXTC0102.PropStrRiyoTodo
            Me.txtRiyoShiku.Text = dataEXTC0102.PropStrRiyoShiku
            Me.txtRiyoBan.Text = dataEXTC0102.PropStrRiyoBan
            Me.txtRiyoBuild.Text = dataEXTC0102.PropStrRiyoBuild
            Me.txtRiyoTel1.Text = dataEXTC0102.PropStrRiyoTel11
            Me.txtRiyoTel2.Text = dataEXTC0102.PropStrRiyoTel12
            Me.txtRiyoTel3.Text = dataEXTC0102.PropStrRiyoTel13
            Me.txtRiyoNaisen.Text = dataEXTC0102.PropStrRiyoNaisen
            Me.txtRiyoMobileTel1.Text = dataEXTC0102.PropStrRiyoTel21
            'Me.txtRiyoMobileTel2.Text = dataEXTC0102.PropStrRiyoTel22 '2016.11.2 m.hayabuchi DEL 課題No.59
            'Me.txtRiyoMobileTel3.Text = dataEXTC0102.PropStrRiyoTel23 '2016.11.2 m.hayabuchi DEL 課題No.59
            Me.txtRiyoFax1.Text = dataEXTC0102.PropStrRiyoFax11
            Me.txtRiyoFax2.Text = dataEXTC0102.PropStrRiyoFax12
            Me.txtRiyoFax3.Text = dataEXTC0102.PropStrRiyoFax13
            '音響
            Me.txtOnkyoNm.Text = dataEXTC0102.PropStrOnkyoNm
            Me.txtOnkyoTantoNm.Text = dataEXTC0102.PropStrOnkyoTantoNm
            Me.txtOnkyoTel11.Text = dataEXTC0102.PropStrOnkyoTel11
            Me.txtOnkyoTel12.Text = dataEXTC0102.PropStrOnkyoTel12
            Me.txtOnkyoTel13.Text = dataEXTC0102.PropStrOnkyoTel13
            Me.txtOnkyoNaisen.Text = dataEXTC0102.PropStrOnkyoNaisen
            Me.txtOnkyoFax11.Text = dataEXTC0102.PropStrOnkyoFax11
            Me.txtOnkyoFax12.Text = dataEXTC0102.PropStrOnkyoFax12
            Me.txtOnkyoFax13.Text = dataEXTC0102.PropStrOnkyoFax13
            Me.txtOnkyoMail.Text = dataEXTC0102.PropStrOnkyoMail
            '事務処理欄
            If dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_MITEI Then
                Me.chkSendKbn.Checked = True
                Me.rdoSendKbn1.Checked = True
            ElseIf dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_DL Then
                Me.chkSendKbn.Checked = False
                Me.rdoSendKbn1.Checked = True
            ElseIf dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_KIBOU Then
                Me.chkSendKbn.Checked = False
                Me.rdoSendKbn2.Checked = True
            End If

            '特記事項
            Me.txtBiko.Text = dataEXTC0102.PropStrBiko
            ' 2015.12.10 UPD START↓ キャンセル情報は正式にしない
            'Me.btnOfficial.Enabled = True
            If dataEXTC0102.PropStrYoyakuSts > YOYAKU_STS_KARI Then
                Me.btnOfficial.Enabled = False
                Me.btnRegister.Enabled = False
            Else
                Me.btnOfficial.Enabled = True
                Me.btnRegister.Enabled = True
            End If
            ' 2015.12.10 UPD END↑ キャンセル情報は正式にしない
            Me.btnCancel.Enabled = True

            Me.pnlKashi.SuspendLayout()
            Me.pnlStudio.SuspendLayout()
            Me.pnlOpe.SuspendLayout()

            '設定処理
            SetEnabledSettings()

        End If

        'スクロールバーを必要な場合のみ表示させます
        fpRiyobi.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fpRiyobi.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

        'コンボボックス初期制御フラグ（ロード完了）          
        strCmbLoadFlg = "1"                                       '2016.08.08 ADD e.watanabe

    End Sub

    ''' <summary>
    ''' 画面項目活性判定処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetEnabledSettings()
        '画面項目活性判定
        '未定がチェック状態の場合、非活性
        If Me.chkKashi.Checked = True Then
            Me.pnlKashi.Enabled = False
        Else
            Me.pnlKashi.Enabled = True
        End If
        ' 2015.12.09 DEL START↓ h.hagiwara  
        'If Me.chkStudio.Checked = True Then
        '    Me.pnlStudio.Enabled = False
        'Else
        '    Me.pnlStudio.Enabled = True
        'End If
        ' 2015.12.09 DEL END↑ h.hagiwara  
        If Me.chkOpe.Checked = True Then
            Me.pnlOpe.Enabled = False
        Else
            Me.pnlOpe.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 【画面表示用】利用日時リストの値を画面表示するSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadDataRiyobi(ByVal dataList As ArrayList)
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        sheet.RowCount = dataList.Count
        sheet.ColumnCount = 7
        Dim index As New Integer
        Dim lineCnt As New Integer

        index = 0
        lineCnt = 1

        Dim maskcell As New FarPoint.Win.Spread.CellType.MaskCellType()
        maskcell.Mask = "99:99"
        maskcell.NullDisplay = "     "
        Dim lineCntCell As FarPoint.Win.Spread.Cell
        Dim cmb As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        'リストに表示されるアイテムを定義
        cmb.Items = New String() {"", "時間貸し", "Lock out"}
        cmb.ItemData = New String() {"", "1", "2"}
        cmb.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

        For Each dataRiyobi As CommonDataRiyobi In dataList
            sheet.Cells(index, 0).Value = False
            lineCntCell = sheet.Cells(index, 1)
            lineCntCell.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
            lineCntCell.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
            sheet.Cells(index, 1).Value = lineCnt.ToString
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = dataRiyobi.PropStrYoyakuDtDisp
            sheet.Cells(index, 2).Locked = True
            sheet.Cells(index, 4).CellType = cmb
            sheet.Cells(index, 4).Value = dataRiyobi.PropStrRiyoKeitai
            sheet.Cells(index, 5).CellType = maskcell
            sheet.Cells(index, 6).CellType = maskcell
            If dataRiyobi.PropStrMiteiFlg = "0" Then
                sheet.Cells(index, 3).Value = False
                sheet.Cells(index, 5).Value = dataRiyobi.PropStrStartTime
                sheet.Cells(index, 6).Value = dataRiyobi.PropStrEndTime
            Else
                sheet.Cells(index, 3).Value = True
                sheet.Cells(index, 5).Value = Nothing
                sheet.Cells(index, 6).Value = Nothing
            End If

            index = index + 1
            lineCnt = lineCnt + 1
        Next

    End Sub

    ''' <summary>
    ''' 【画面→DB用】Dataに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRiyobiSpreadData(ByRef dataList As ArrayList)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        sheet.RowCount = dataList.Count
        sheet.ColumnCount = 7
        Dim index As New Integer
        Dim lineCnt As New Integer

        index = 0
        lineCnt = 1
        For Each dataRiyobi As CommonDataRiyobi In dataList
            If sheet.Cells(index, 3).Value = True Then
                dataRiyobi.PropStrMiteiFlg = "1"
            Else
                dataRiyobi.PropStrMiteiFlg = "0"
            End If
            dataRiyobi.PropStrRiyoKeitai = sheet.Cells(index, 4).Value
            dataRiyobi.PropStrStartTime = sheet.Cells(index, 5).Value
            dataRiyobi.PropStrEndTime = sheet.Cells(index, 6).Value
            dataRiyobi.PropStrStudioKbn = dataEXTC0102.PropStrStudioKbn
            index = index + 1
            lineCnt = lineCnt + 1
        Next

    End Sub

    ''' <summary>
    ''' 貸出種別
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkKashi_CheckedChanged(sender As Object, e As EventArgs) Handles chkKashi.CheckedChanged
        If Me.chkKashi.Checked = True Then
            Me.pnlKashi.Enabled = False
        Else
            Me.pnlKashi.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' スタジオ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkStudio_CheckedChanged(sender As Object, e As EventArgs)
        ' 2015.12.09 DEL START↓ h.hagiwara
        'If Me.chkStudio.Checked = True Then
        '    Me.pnlStudio.Enabled = False
        'Else
        '    Me.pnlStudio.Enabled = True
        'End If
        ' 2015.12.09 DEL END↑ h.hagiwara
    End Sub

    ''' <summary>
    ''' 音響オペ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkOpe_CheckedChanged(sender As Object, e As EventArgs) Handles chkOpe.CheckedChanged
        If Me.chkOpe.Checked = True Then
            Me.pnlOpe.Enabled = False
        Else
            Me.pnlOpe.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 送信区分
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkSendKbn_CheckedChanged(sender As Object, e As EventArgs) Handles chkSendKbn.CheckedChanged
        If Me.chkSendKbn.Checked = True Then
            Me.pnlSendKbn.Enabled = False
        Else
            Me.pnlSendKbn.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 利用日追加ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnYoyakuAdd_Click(sender As Object, e As EventArgs) Handles btnYoyakuAdd.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '画面入力情報をDataに格納
        convertData(dataEXTC0102)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0102.PropListRiyobi)
        Dim frm As New EXTZ0202
        '「利用日追加」画面を表示
        frm.PropStrTorokuKbn = TOUROKU_KBN_KARI
        frm.PropLstRiyobi = dataEXTC0102.PropListRiyobi
        frm.PropStrYoyakuNo = dataEXTC0102.PropStrYoyakuNo
        frm.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.PropStrStudioKbn = dataEXTC0102.PropStrStudioKbn

        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim comp As New CommonDataRiyobiCompareter()
        dataEXTC0102.PropListRiyobi.Sort(comp)
        SetSpreadDataRiyobi(dataEXTC0102.PropListRiyobi)
    End Sub

    ''' <summary>
    ''' 利用日削除ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnYoyakuDel_Click(sender As Object, e As EventArgs) Handles btnYoyakuDel.Click
        '画面入力情報をDataに格納
        convertData(dataEXTC0102)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0102.PropListRiyobi)
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        Dim i As Integer = 0
        Dim rowCnt As Integer = sheet.RowCount
        Dim lstRiyobi As ArrayList = dataEXTC0102.PropListRiyobi
        Dim dataRiyobi As CommonDataRiyobi

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '1件の場合、0件になる場合は削除させない
        Dim blnIsNotDeleteLine As Boolean = False
        Do While i < rowCnt
            If sheet.Cells(i, 0).Value = False Then
                blnIsNotDeleteLine = True
            End If
            i = i + 1
        Loop
        If blnIsNotDeleteLine = False Then
            MsgBox(CommonDeclareEXT.E2015, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        SetRiyobiSpreadData(dataEXTC0102.PropListRiyobi)
        i = 0
        '削除開始
        Dim j As Integer = 0
        Do While i < rowCnt
            If sheet.Cells(i, 0).Value = True Then
                'DBから制御データ削除
                dataRiyobi = lstRiyobi(i - j)
                If logicEXTZ0202.DeleteYoyakuCtlData(dataRiyobi) = False Then
                    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                'リストから削除
                lstRiyobi.RemoveAt(i - j)
                j = j + 1
            End If
            i = i + 1
        Loop
        SetSpreadDataRiyobi(dataEXTC0102.PropListRiyobi)
    End Sub

    ''' <summary>
    ''' 正式登録ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOfficial_Click(sender As Object, e As EventArgs) Handles btnOfficial.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '入力チェック
        If inputCheckMain() = False Then
            Exit Sub
        End If
        '画面入力情報をDataに格納
        convertData(dataEXTC0102)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0102.PropListRiyobi)
        If inputCheckCondition(True) = False Then
            Exit Sub
        End If
        '確認処理
        If Me.chkSendKbn.Checked Then
            If MsgBox(String.Format(CommonEXT.C0001), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
                Return
            End If
        ElseIf MsgBox(String.Format(CommonEXT.C0003), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '更新処理
        '予約NOあり
        '予約制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '更新
        If logicEXTC0102.RegYoyakuInfo(dataEXTC0102, True) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表削除
        If logicEXTC0102.DeleteYoyakuList(dataEXTC0102.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTC0102.InsertYoyakuList(dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        '正式予約画面
        Dim frm As New EXTC0103
        'パラメータセット
        frm.dataEXTC0102 = dataEXTC0102
        '「正式予約」画面を表示
        Me.Hide()
        frm.ShowDialog()
        Me.Close()

    End Sub

    ''' <summary>
    ''' 更新ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '入力チェック
        If inputCheckMain() = False Then
            Exit Sub
        End If
        '画面入力情報をDataに格納
        convertData(dataEXTC0102)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0102.PropListRiyobi)
        If inputCheckCondition(False) = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0002, "予約情報"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '更新処理
        If String.IsNullOrEmpty(dataEXTC0102.PropStrYoyakuNo) = True Then
            '予約NOなし
            ''予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            'If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '登録
            If logicEXTC0102.RegYoyakuInfo(dataEXTC0102, False) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTC0102.InsertYoyakuList(dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            'キャンセル情報更新
            If String.IsNullOrEmpty(dataEXTC0102.PropStrCancelNo) = False Then
                Dim dataEXTC0104 As New DataEXTC0104
                dataEXTC0104.PropStrYoyakuNo = dataEXTC0102.PropStrCancelNo
                If logicEXTC0104.GetYoyakuData(dataEXTC0104) = True Then
                    dataEXTC0104.PropStrYoyakuSts = CANCEL_STS_KARI
                    If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, True) = False Then
                        MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                        Return
                    End If
                End If
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        Else
            '予約NOあり
            ''予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            'If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '更新
            If logicEXTC0102.RegYoyakuInfo(dataEXTC0102, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表削除
            If logicEXTC0102.DeleteYoyakuList(dataEXTC0102.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTC0102.InsertYoyakuList(dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTC0102_Load(sender, e)
    End Sub

    ''' <summary>
    ''' 画面入力情報をDataに格納
    ''' </summary>
    ''' <remarks>画面入力情報をDataに格納
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertData(ByRef dataEXTC0102 As DataEXTC0102)
        'ヘッダ
        '仮予約受付USERCDが空の場合設定する
        If String.IsNullOrEmpty(dataEXTC0102.PropStrKariUsercd) = True Then
            dataEXTC0102.PropStrKariUsercd = CommonDeclareEXT.PropComStrUserId
        End If
        '受付日
        dataEXTC0102.PropStrKariukeDt = Me.dtpKariuke.txtDate.Text

        '催事
        dataEXTC0102.PropStrShutsuenNm = Me.txtShutuen.Text

        '貸し出し種別
        If Me.chkKashi.Checked = True Then
            dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI
        Else
            If Me.rdoKashi1.Checked = True Then
                dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
            ElseIf Me.rdoKashi2.Checked = True Then
                dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
            End If
        End If
        'STUDIO
        'If Me.chkStudio.Checked = True Then                                                      ' 2015.12.09 DEL h.hagiwara                                           
        '    dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_MITEI                        ' 2015.12.09 DEL h.hagiwara
        'Else                                                                                     ' 2015.12.09 DEL h.hagiwara
        If Me.rdoStudio1.Checked = True Then
            dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_201
        ElseIf Me.rdoStudio2.Checked = True Then
            dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_202
        ElseIf Me.rdoStudio3.Checked = True Then
            dataEXTC0102.PropStrStudioKbn = CommonDeclareEXT.STUDIO_HOUSE_LOCK
        End If
        'End If                                                                                   ' 2015.12.09 DEL h.hagiwara
        '音響
        If Me.chkOpe.Checked = True Then
            dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_MITEI
        Else
            If Me.rdoOpe1.Checked = True Then
                dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_ARI
            ElseIf Me.rdoOpe2.Checked = True Then
                dataEXTC0102.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_NASHI
            End If
        End If
        '利用者情報
        dataEXTC0102.PropStrRiyoshaCd = Me.lblRiyoshaCd.Text
        dataEXTC0102.PropStrRiyoKana = Me.txtRiyoshaNmKana.Text
        dataEXTC0102.PropStrRiyoNm = Me.txtRiyoshaNm.Text
        dataEXTC0102.PropStrDaihyoNm = Me.txtDaihyoNm.Text
        dataEXTC0102.PropStrSekininBushoNm = Me.txtSekininBushoNm.Text
        dataEXTC0102.PropStrSekininNm = Me.cmbSekininNm.Text
        dataEXTC0102.PropStrSekininMail = Me.txtSekininMail.Text
        dataEXTC0102.PropStrAiteNm = Me.lblExasAiteNm.Text
        dataEXTC0102.PropStrAiteCd = Me.lblExasAite.Text
        dataEXTC0102.PropStrRiyoYubin1 = Me.txtRiyoPost1.Text
        dataEXTC0102.PropStrRiyoYubin2 = Me.txtRiyoPost2.Text
        dataEXTC0102.PropStrRiyoTodo = Me.cmbRiyoTodo.Text
        dataEXTC0102.PropStrRiyoShiku = Me.txtRiyoShiku.Text
        dataEXTC0102.PropStrRiyoBan = Me.txtRiyoBan.Text
        dataEXTC0102.PropStrRiyoBuild = Me.txtRiyoBuild.Text
        dataEXTC0102.PropStrRiyoTel11 = Me.txtRiyoTel1.Text
        dataEXTC0102.PropStrRiyoTel12 = Me.txtRiyoTel2.Text
        dataEXTC0102.PropStrRiyoTel13 = Me.txtRiyoTel3.Text
        dataEXTC0102.PropStrRiyoNaisen = Me.txtRiyoNaisen.Text
        dataEXTC0102.PropStrRiyoTel21 = Me.txtRiyoMobileTel1.Text
        'dataEXTC0102.PropStrRiyoTel22 = Me.txtRiyoMobileTel2.Text' 2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTC0102.PropStrRiyoTel23 = Me.txtRiyoMobileTel3.Text' 2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTC0102.PropStrRiyoFax11 = Me.txtRiyoFax1.Text
        dataEXTC0102.PropStrRiyoFax12 = Me.txtRiyoFax2.Text
        dataEXTC0102.PropStrRiyoFax13 = Me.txtRiyoFax3.Text
        dataEXTC0102.PropStrOnkyoNm = Me.txtOnkyoNm.Text
        dataEXTC0102.PropStrOnkyoTantoNm = Me.txtOnkyoTantoNm.Text
        dataEXTC0102.PropStrOnkyoTel11 = Me.txtOnkyoTel11.Text
        dataEXTC0102.PropStrOnkyoTel12 = Me.txtOnkyoTel12.Text
        dataEXTC0102.PropStrOnkyoTel13 = Me.txtOnkyoTel13.Text
        dataEXTC0102.PropStrOnkyoNaisen = Me.txtOnkyoNaisen.Text
        dataEXTC0102.PropStrOnkyoFax11 = Me.txtOnkyoFax11.Text
        dataEXTC0102.PropStrOnkyoFax12 = Me.txtOnkyoFax12.Text
        dataEXTC0102.PropStrOnkyoFax13 = Me.txtOnkyoFax13.Text
        dataEXTC0102.PropStrOnkyoMail = Me.txtOnkyoMail.Text

        '事務処理欄
        If Me.chkSendKbn.Checked = True Then
            dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_MITEI
        Else
            If Me.rdoSendKbn1.Checked = True Then
                dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_DL
            ElseIf Me.rdoSendKbn2.Checked = True Then
                dataEXTC0102.PropStrSendKbn = CommonDeclareEXT.SEND_KBN_KIBOU
            End If
        End If

        '特記事項
        dataEXTC0102.PropStrBiko = Me.txtBiko.Text

    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function inputCheckMain() As Boolean

        'トレースログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_START", Nothing, Nothing)

        Try
            '受付日
            If String.IsNullOrEmpty(Me.dtpKariuke.txtDate.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "受付日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfChar(Me.dtpKariuke.txtDate.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0004, "受付日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            'アーティスト名
            If String.IsNullOrEmpty(Me.txtShutuen.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "アーティスト名"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            'スタジオ
            ' 2015.12.09 UPD START↓ h.hagiwara
            'If Me.chkStudio.Checked Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0002, "利用スタジオ"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            If Me.rdoStudio1.Checked = False Then
                If Me.rdoStudio2.Checked = False Then
                    If Me.rdoStudio3.Checked = False Then
                        MsgBox(String.Format(CommonDeclareEXT.E0002, "利用スタジオ"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                End If
            End If
            ' 2015.12.09 UPD END↑ h.hagiwara
            'スプレッド
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
            Dim index As New Integer
            Dim lineCnt As New Integer

            index = 0
            lineCnt = 1
            If 0 = sheet.RowCount Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望日時"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                If sheet.Cells(index, 3).Value = False Then
                    If String.IsNullOrEmpty(sheet.Cells(index, 5).Value) Then
                        MsgBox(String.Format(CommonDeclareEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                    If String.IsNullOrEmpty(sheet.Cells(index, 6).Value) Then
                        MsgBox(String.Format(CommonDeclareEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                End If
                If commonValidate.IsHalfNmb(sheet.Cells(index, 5).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfNmb(sheet.Cells(index, 6).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                '利用形態
                If String.IsNullOrEmpty(sheet.Cells(index, 4).Value) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0002, "利用形態"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            Next

            '利用者名（カナ）
            If commonLogicEXT.IsFullKana(Me.txtRiyoshaNmKana.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0007, "利用者名（カナ）"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '責任者メールアドレス
            If mailRegexUtilities.IsValidEmail(Me.txtSekininMail.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0026, "責任者メールアドレス"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '郵便番号
            If commonValidate.IsHalfNmb(Me.txtRiyoPost1.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "郵便番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(Me.txtRiyoPost1.Text) = False Then
                If Len(Me.txtRiyoPost1.Text) <> 3 Then
                    MsgBox(String.Format(CommonDeclareEXT.E0010, "郵便番号1", 3), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            If commonValidate.IsHalfNmb(Me.txtRiyoPost2.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "郵便番号2"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(Me.txtRiyoPost2.Text) = False Then
                If Len(Me.txtRiyoPost2.Text) <> 4 Then
                    MsgBox(String.Format(CommonDeclareEXT.E0010, "郵便番号2", 4), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            '電話番号
            If commonValidate.IsHalfNmb(Me.txtRiyoTel1.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "電話番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtRiyoTel2.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "電話番号2"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtRiyoTel3.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "電話番号3"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '内線番号
            If commonValidate.IsHalfNmb(Me.txtRiyoNaisen.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "内線番号"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '携帯番号
            '2016.11.4 m.hayabuchi MOD Start 課題No.59
            'If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel1.Text) = False Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号1"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            'If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel2.Text) = False Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号2"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            'If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel3.Text) = False Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号3"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel1.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '2016.11.4 m.hayabuchi MOD End 課題No.59
            'FAX番号
            If commonValidate.IsHalfNmb(Me.txtRiyoFax1.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "FAX番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtRiyoFax2.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "FAX番号2"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtRiyoFax3.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "FAX番号3"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '音響会社電話番号
            If commonValidate.IsHalfNmb(Me.txtOnkyoTel11.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社電話番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtOnkyoTel12.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社電話番号2"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtOnkyoTel13.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社電話番号3"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '内線番号
            If commonValidate.IsHalfNmb(Me.txtOnkyoNaisen.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社内線番号"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            'FAX番号
            If commonValidate.IsHalfNmb(Me.txtOnkyoFax11.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社FAX番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtOnkyoFax12.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社FAX番号2"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtOnkyoFax13.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "音響会社FAX番号3"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If

            'トレースログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '異常終了
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック・条件
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>関連チェックを行う
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function inputCheckCondition(ByVal blnIsOfficial As Boolean) As Boolean

        'トレースログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_START", Nothing, Nothing)

        Try
            If blnIsOfficial Then
                If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2018, "利用者番号"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If Me.chkKashi.Checked Then
                    MsgBox(String.Format(CommonDeclareEXT.E2019, "貸出種別"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                ' 2015.12.09 UPD START↓ h.hagiwara
                'If Me.chkStudio.Checked Then
                '    MsgBox(String.Format(CommonDeclareEXT.E2019, "利用スタジオ"), MsgBoxStyle.Exclamation, "エラー")
                '    Return False
                'End If
                If Me.rdoStudio1.Checked = False Then
                    If Me.rdoStudio2.Checked = False Then
                        If Me.rdoStudio3.Checked = False Then
                            MsgBox(String.Format(CommonDeclareEXT.E2019, "利用スタジオ"), MsgBoxStyle.Exclamation, "エラー")
                            Return False
                        End If
                    End If
                End If
                ' 2015.12.09 UPD END↑ h.hagiwara
                If Me.chkOpe.Checked Then
                    MsgBox(String.Format(CommonDeclareEXT.E2019, "音響オペレータ"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            '電話番号
            If String.IsNullOrEmpty(Me.txtRiyoTel1.Text) = False _
                Or String.IsNullOrEmpty(Me.txtRiyoTel2.Text) = False _
                Or String.IsNullOrEmpty(Me.txtRiyoTel3.Text) = False Then
                If String.IsNullOrEmpty(Me.txtRiyoTel1.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtRiyoTel2.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtRiyoTel3.Text) = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2020, "電話番号"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            '携帯番号（必須チェック不要）
            ' 2016.11.2 m.hayabuchi MOD Start 課題No.59
            'If String.IsNullOrEmpty(Me.txtRiyoMobileTel1.Text) = False Then _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel2.Text) = False _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel3.Text) = False Then
            '    If String.IsNullOrEmpty(Me.txtRiyoMobileTel1.Text) = True Then _
            '        Or String.IsNullOrEmpty(Me.txtRiyoMobileTel2.Text) = True _
            '        Or String.IsNullOrEmpty(Me.txtRiyoMobileTel3.Text) = True Then
            'MsgBox(String.Format(CommonDeclareEXT.E2020, "携帯番号"), MsgBoxStyle.Exclamation, "エラー")
            'Return False
            'End If
            'End If
            ' 2016.11.2 m.hayabuchi MOD End 課題No.59
            'FAX番号
            If String.IsNullOrEmpty(Me.txtRiyoFax1.Text) = False _
                Or String.IsNullOrEmpty(Me.txtRiyoFax2.Text) = False _
                Or String.IsNullOrEmpty(Me.txtRiyoFax3.Text) = False Then
                If String.IsNullOrEmpty(Me.txtRiyoFax1.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtRiyoFax2.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtRiyoFax3.Text) = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2020, "FAX番号"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            '電話番号
            If String.IsNullOrEmpty(Me.txtOnkyoTel11.Text) = False _
                Or String.IsNullOrEmpty(Me.txtOnkyoTel12.Text) = False _
                Or String.IsNullOrEmpty(Me.txtOnkyoTel13.Text) = False Then
                If String.IsNullOrEmpty(Me.txtOnkyoTel11.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtOnkyoTel12.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtOnkyoTel13.Text) = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2020, "音響会社電話番号"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            'FAX番号
            If String.IsNullOrEmpty(Me.txtOnkyoFax11.Text) = False _
                Or String.IsNullOrEmpty(Me.txtOnkyoFax12.Text) = False _
                Or String.IsNullOrEmpty(Me.txtOnkyoFax13.Text) = False Then
                If String.IsNullOrEmpty(Me.txtOnkyoFax11.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtOnkyoFax12.Text) = True _
                    Or String.IsNullOrEmpty(Me.txtOnkyoFax13.Text) = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2020, "音響会社FAX番号"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            '利用日(SPREAD)
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
            Dim index As Integer = 0
            Dim lineCnt As Integer = 1
            Dim listRiyobi As New ArrayList
            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                '利用日時未設定
                ' 2015.12.08 DEL START↓ h.hagiwara 
                'If blnIsOfficial Then
                '    If sheet.Cells(index, 3).Value = True Then
                '        MsgBox(String.Format(CommonDeclareEXT.E2019, "利用日時"), MsgBoxStyle.Exclamation, "エラー")
                '        Return False
                '    End If
                'End If
                ' 2015.12.08 DEL END↑ h.hagiwara 
                '予約数超過
                If logicEXTZ0202.CheckRiyobiRegister(TOUROKU_KBN_KARI, dataRiyobi) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If listRiyobi.Contains(dataRiyobi.PropStrYoyakuDt) Then
                    MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                Else
                    listRiyobi.Add(dataRiyobi.PropStrYoyakuDt)
                End If
            Next

            'トレースログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '異常終了
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 予約取消ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim frm As New EXTZ0211
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If frm.PropStrDeleteKbn = TORIKESHI_KBN_CANCEL Then
            '論理削除
            '予約制御削除
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '更新
            dataEXTC0102.PropStrYoyakuSts = YOYAKU_STS_CANCEL_KARI
            If logicEXTC0102.RegYoyakuInfo(dataEXTC0102, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.PropStrDeleteKbn = TORIKESHI_KBN_DELETE Then
            '物理削除
            '予約制御削除
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTC0102.DeleteYoyakuList(dataEXTC0102.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTC0102.DeleteYoyaku(dataEXTC0102.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.DialogResult = System.Windows.Forms.DialogResult.OK Then
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' 利用者登録ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRiyoshaAdd_Click(sender As Object, e As EventArgs) Handles btnRiyoshaAdd.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim frm As New EXTM0202
        '利用者番号設定済でないかチェック
        If logicEXTC0102.ChkRiyoshaRegister(Me.txtRiyoshaNmKana.Text) = False Then
            MsgBox(CommonEXT.E2016, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '画面を表示
        With frm.dataEXTM0202
            .PropParamRiyoNm = Me.txtRiyoshaNm.Text
            .PropParamRiyoKana = Me.txtRiyoshaNmKana.Text
            .PropParamDaihyoNm = Me.txtDaihyoNm.Text
            .PropParamRiyoTel11 = Me.txtRiyoTel1.Text
            .PropParamRiyoTel12 = Me.txtRiyoTel2.Text
            .PropParamRiyoTel13 = Me.txtRiyoTel3.Text
            .PropParamRiyoNaisen = Me.txtRiyoNaisen.Text
            .PropParamRiyoFax11 = Me.txtRiyoFax1.Text
            .PropParamRiyoFax12 = Me.txtRiyoFax2.Text
            .PropParamRiyoFax13 = Me.txtRiyoFax3.Text
            .PropParamRiyoYubin1 = Me.txtRiyoPost1.Text
            .PropParamRiyoYubin2 = Me.txtRiyoPost2.Text
            .PropParamRiyoTodo = Me.cmbRiyoTodo.Text
            .PropParamRiyoShiku = Me.txtRiyoShiku.Text
            .PropParamRiyoBan = Me.txtRiyoBan.Text
            .PropParamRiyoBuild = Me.txtRiyoBuild.Text
        End With
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '利用者番号設定
            With frm.DataEXTM0202
                Me.lblRiyoshaCd.Text = .PropLblRiyo_cd.Text
                If frm.dataEXTM0202.PropRdoTujyo.Checked = True Then
                    dataEXTC0102.PropStrRiyoLvl = RIYOSHA_LV1
                ElseIf frm.dataEXTM0202.PropRdoChui.Checked = True Then
                    dataEXTC0102.PropStrRiyoLvl = RIYOSHA_LV2
                ElseIf frm.dataEXTM0202.PropRdoHuka.Checked = True Then
                    dataEXTC0102.PropStrRiyoLvl = RIYOSHA_LV3
                End If
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTC0102.PropStrRiyoLvl)
                Me.txtRiyoshaNm.Text = .PropTxtRiyo_nm.Text
                Me.txtRiyoshaNmKana.Text = .PropTxtRiyo_kana.Text
                Me.txtDaihyoNm.Text = .PropTxtDaihyo_nm.Text
                Me.txtRiyoTel1.Text = .PropTxtTel1.Text
                Me.txtRiyoTel2.Text = .PropTxtTel2.Text
                Me.txtRiyoTel3.Text = .PropTxtTel3.Text
                Me.txtRiyoNaisen.Text = .PropTxtNaisen.Text
                Me.txtRiyoFax1.Text = .PropTxtFax1.Text
                Me.txtRiyoFax2.Text = .PropTxtFax2.Text
                Me.txtRiyoFax3.Text = .PropTxtFax3.Text
                Me.txtRiyoPost1.Text = .PropTxtYubin1.Text
                Me.txtRiyoPost2.Text = .PropTxtYubin2.Text
                Me.cmbRiyoTodo.Text = .PropCmbTodo.Text
                Me.txtRiyoShiku.Text = .PropTxtShiku.Text
                Me.txtRiyoBan.Text = .PropTxtBanchi.Text
                Me.txtRiyoBuild.Text = .PropTxtBuild.Text
                Me.lblExasAite.Text = .PropLblAite_cd.Text
                Me.lblExasAiteNm.Text = .PropLblAite_nm.Text
            End With

            If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If
                Dim list As ArrayList = logicEXTC0102.GetSekininshaList(dataEXTC0102.PropStrRiyoshaCd)
                cmbSekininNm.Items.Clear()
                For i = 0 To list.Count - 1
                    cmbSekininNm.Items.Add(list(i))
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' 利用者検索ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRiyoshaSearch_Click(sender As Object, e As EventArgs) Handles btnRiyoshaSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim frm As New EXTM0201
        'パラメータ設定
        frm.DataEXTM0201.PropTxtRiyo_kana = Me.txtRiyoshaNmKana
        frm.dataEXTM0201.PropParamValue = "0"
        '画面を表示
        frm.ShowDialog()
        '利用者番号設定
        If String.IsNullOrEmpty(frm.DataEXTM0201.PropParamRiyoCd) = False Then
            With frm.DataEXTM0201
                Me.lblRiyoshaCd.Text = .PropParamRiyoCd
                dataEXTC0102.PropStrRiyoLvl = .PropParamRiyoLvl
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTC0102.PropStrRiyoLvl)
                Me.txtRiyoshaNm.Text = .PropParamRiyoNm
                Me.txtRiyoshaNmKana.Text = .PropParamRiyoKana
                Me.txtDaihyoNm.Text = .PropParamDaihyoNm
                Me.txtRiyoTel1.Text = .PropParamRiyoTel11
                Me.txtRiyoTel2.Text = .PropParamRiyoTel12
                Me.txtRiyoTel3.Text = .PropParamRiyoTel13
                Me.txtRiyoNaisen.Text = .PropParamRiyoNaisen
                Me.txtRiyoFax1.Text = .PropParamRiyoFax11
                Me.txtRiyoFax2.Text = .PropParamRiyoFax12
                Me.txtRiyoFax3.Text = .PropParamRiyoFax13
                Me.txtRiyoPost1.Text = .PropParamRiyoYubin1
                Me.txtRiyoPost2.Text = .PropParamRiyoYubin2
                Me.cmbRiyoTodo.Text = .PropParamRiyoTodo
                Me.txtRiyoShiku.Text = .PropParamRiyoShiku
                Me.txtRiyoBan.Text = .PropParamRiyoBan
                Me.txtRiyoBuild.Text = .PropParamRiyoBuild
                Me.lblExasAite.Text = .PropParamAiteCd
                Me.lblExasAiteNm.Text = .PropParamAiteNm
            End With

            ' 2016.11.04 ADD START m.hayabuchi        ' 責任者情報クリア（責任者名,メールアドレス,携帯電話番号）
            Me.cmbSekininNm.Text = ""
            Me.txtSekininMail.Text = ""
            Me.txtRiyoMobileTel1.Text = ""
            ' 2016.11.04 ADD END m.hayabuchi          ' 責任者情報クリア（責任者名,メールアドレス,携帯電話番号）

            '責任者再取得
            If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If
                Dim list As ArrayList = logicEXTC0102.GetSekininshaList(Me.lblRiyoshaCd.Text)
                cmbSekininNm.Items.Clear()
                ' 2016.08.09 UPD START e.watanabe　' ELSE追加（0件の場合）変更
                'For i = 0 To list.Count - 1
                'cmbSekininNm.Items.Add(list(i))
                'Next
                If list.Count > 0 Then
                    For i = 0 To list.Count - 1
                        cmbSekininNm.Items.Add(list(i))
                    Next
                    '0件の場合項目クリア
                Else
                    cmbSekininNm.Items.Add("")
                End If
                ' 2016.08.09 UPD END e.watanabe    ' ELSE追加（0件の場合）変更
            End If
        End If
    End Sub

    ''' <summary>
    ''' 戻るボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '予約受付制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        End If
        Me.Close()
    End Sub

    ''' <summary>
    ''' 閉じるボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTC0102_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '予約受付制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        End If
    End Sub

    ''' <summary>責任者名:コンボボックス選択時処理</summary>
    ''' <remarks>コンボボックスで選択された責任者のメールアドレス・携帯電話番号を入力する
    ''' <para>作成情報：2016/11/4 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub cmbSekininNm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSekininNm.SelectedIndexChanged

        '初期表示以外実行
        If strCmbLoadFlg = "1" Then

            'プロパティにセット
            With dataEXTC0102
                .PropStrSekininNm = Me.cmbSekininNm.Text   '画面.責任者名
                .PropStrRiyoshaCd = Me.lblRiyoshaCd.Text   '画面.利用者コード
            End With

            'SQL呼出
            If dataEXTC0102.PropStrSekininNm.Equals("") = False Then
                If dataEXTC0102.PropStrRiyoshaCd.Equals("") = False Then
                    If logicEXTC0102.GetSekininshaMailTel(dataEXTC0102) = False Then
                        MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                End If
            End If

            'テキストボックスにセット
            Me.txtSekininMail.Text = dataEXTC0102.PropStrSekininMail        '責任者.メールアドレス
            Me.txtRiyoMobileTel1.Text = dataEXTC0102.PropStrRiyoTel21       '責任者.携帯電話番号１


        End If

    End Sub

End Class
