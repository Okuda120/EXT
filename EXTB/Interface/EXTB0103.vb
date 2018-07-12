Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win
Imports FarPoint.Win.Spread.CellType
Imports EXTM
Imports GrapeCity.Win.Spread.InputMan.CellType
Imports FarPoint.Win.Spread
Imports System.IO
Imports System.Text
Imports System.Configuration

''' <summary>
''' EXTB0103
''' </summary>
''' <remarks>正式予約（シアター）画面
''' <para>作成情報：2015/08/20 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTB0103

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogicEXTB As New CommonLogicEXTB      '共通クラス
    Private mailRegexUtilities As New MailRegexUtilities
    Public extz0207 As New EXTZ0207
    '変数宣言
    Public dataCommon As New CommonDataEXT      '共通データクラス
    Public dataEXTB0102 As New DataEXTB0102     'データクラス
    Public dataEXTB0103 As New DataEXTB0103     'データクラス（印刷系）
    Public dataEXTZ0207 As New DataEXTZ0207     'データクラス（印刷系）
    Public logicEXTB0102 As New LogicEXTB0102   'ロジッククラス(仮予約と共通の処理はこちら)
    Public logicEXTB0103 As New LogicEXTB0103   'ロジッククラス
    Public logicEXTZ0202 As New LogicEXTZ0202   'ロジッククラス(予約受付制御関連)
    Private strFlashFlg As String               '請求情報更新ボタン押下で色が消えるフラグ
    Private strIraiFlg As String                '依頼完了
    Private strShoninFlg As String              '承認完了
    Private strFutaiFlg As String               '承認完了
    Private strPrintRiyoShoninFlg As String     '印刷利用承認
    Private strPrintFutaiTotalFlg As String     '印刷付帯
    Private strPrintExasFlg As String           'Exas出力
    Public dataEXTZ0214 As New DataEXTZ0214   'ロジッククラス(グループ請求関連)        ' 2016.01.20 ADD y.morooka グループ請求対応
    Private strCmbLoadFlg As String           'コンボボックス初期制御フラグ            ' 2016.08.08 ADD e.watanabe

    'SPREAD ボタン押下位置判別
    '利用日時
    Public Const FB_RIYOBI_DISP As Integer = 6
    '請求
    Public Const FB_BILLPAY_INPUT As Integer = 4
    Public Const FB_NYUKIN_INPUT As Integer = 13
    '付帯
    Public Const FB_FUTAI_EDIT As Integer = 1
    Public Const FB_FUTAI_COPY As Integer = 2
    Public Const FB_FUTAI_PRINT As Integer = 5

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTB0103_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'FLG初期化
        strFlashFlg = "0"
        strIraiFlg = "0"
        strShoninFlg = "0"
        strFutaiFlg = "0"
        strPrintRiyoShoninFlg = "0"
        strPrintFutaiTotalFlg = "0"
        strPrintExasFlg = "0"
        strCmbLoadFlg = "0"                   '2016.08.08 ADD e.watanabe
        '画面項目初期化
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        dataEXTB0102.PropStrShisetuKbn = SHISETU_KBN_THEATER
        'コンボボックスの内容設定
        commonLogicEXT.TodohukenLst(Me.cmbRiyoTodo)
        'メインデータ取得
        Me.lblYoyakuNo.Text = dataEXTB0102.PropStrYoyakuNo

        If logicEXTB0102.GetYoyakuData(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If logicEXTB0102.GetRiyobiData(dataEXTB0102) = True Then
            '利用日データ設定
            Dim comp As New CommonDataRiyobiCompareter()
            dataEXTB0102.PropListRiyobi.Sort(comp)
            SetSpreadDataRiyobi(dataEXTB0102.PropListRiyobi)
        End If
        '各種データセット取得=====================================================================
        '確認依頼/記録
        If logicEXTB0103.GetApproval(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            SetSpreadIraiRireki(dataEXTB0102.PropDsApproval.Tables("IRAI_RIREKI_TBL"))
            SetSpreadIraiKakunin(dataEXTB0102.PropDsApproval.Tables("CHECK_RIREKI_TBL"))
        End If
        'タイムスケジュール
        If logicEXTB0103.GetTimeSchedule(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            SetSpreadTimeSchedule(dataEXTB0102.PropDsTimeSch.Tables("TIME_SCHEDULE_TBL"))
        End If
        'チケット
        If logicEXTB0103.GetTicketInfo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            SetSpreadTicket(dataEXTB0102.PropDsTicket.Tables("TICKET_TBL"))
        End If
        '付帯
        If logicEXTB0103.GetFutai(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            '設定
            SetSpreadFutai(dataEXTB0102.PropHtFutai)
        End If
        '請求
        dataEXTB0102.propDblTax = logicEXTB0103.GetTaxfirstRiyobi(dataEXTB0102)
        If logicEXTB0103.GetBillReq(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            If SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")) = False Then
                '請求データが存在しない場合、初期データを作成する
                InitBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
                SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
            End If
        End If
        'プロジェクト情報
        If logicEXTB0103.GetExasRiyoryo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        'EXAS入金情報
        If logicEXTB0103.GetNyukin(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '現金・suica
        If logicEXTB0103.GetCashAndElectroMoney(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        SetSpreadAlsokMeisai(dataEXTB0102.PropDsCashElectro.Tables("ALSOK_MEISAI"))
        SetSpreadAlsokStore(dataEXTB0102.PropDsCashElectro.Tables("ALSOK_STORE"))
        '=========================================================================================
        'ヘッダ
        ' 2016.11.1 ADD m.hayabuchi Start 
        dataEXTB0102.PropStrKakuteiDt = dataEXTB0102.PropStrKariukeDt
        ' 2016.11.1 ADD m.hayabuchi End
        Me.dtpUke.txtDate.Text = dataEXTB0102.PropStrKakuteiDt
        '仮予約からの場合、正式予約に
        If dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_KARI Then
            Me.lblStatus.Text = commonLogicEXTB.GetYoyakuSts(YOYAKU_STS_SEISHIKI)
            Me.lblStatus.BackColor = commonLogicEXTB.GetYoyakuStsColor(YOYAKU_STS_SEISHIKI)
        Else
            Me.lblStatus.Text = commonLogicEXTB.GetYoyakuSts(dataEXTB0102.PropStrYoyakuSts)
            Me.lblStatus.BackColor = commonLogicEXTB.GetYoyakuStsColor(dataEXTB0102.PropStrYoyakuSts)
        End If
        If dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_KARI Then
            Me.rdoStatus1.Checked = True
        ElseIf dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_SEISHIKI Then
            Me.rdoStatus1.Checked = True
        ElseIf dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_SEISHIKI_COMP Then
            Me.rdoStatus2.Checked = True
        End If

        Me.lblAddUserCd.Text = dataEXTB0102.PropStrAddUserCd
        Me.lblAddUserNm.Text = dataEXTB0102.PropStrAddUserNm
        Me.lblAddUserDate.Text = dataEXTB0102.PropStrAddDt
        Me.lblUpUserCd.Text = dataEXTB0102.PropStrUpUserCd
        Me.lblUpUserNm.Text = dataEXTB0102.PropStrUpUserNm
        Me.lblUpUserDate.Text = dataEXTB0102.PropStrUpDt
        '催事
        Me.txtSaiji.Text = dataEXTB0102.PropStrSaijiNm
        Me.txtShutuen.Text = dataEXTB0102.PropStrShutsuenNm
        '貸し出し種別
        If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
            Me.rdoKashi1.Checked = True
        ElseIf dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
            Me.rdoKashi2.Checked = True
        ElseIf dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI Then
            Me.rdoKashi3.Checked = True
        End If
        '利用形状
        If dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_STAND Then
            Me.rdoRiyoType1.Checked = True
        ElseIf dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SEATING Then
            Me.rdoRiyoType2.Checked = True
        ElseIf dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MIX Then
            Me.rdoRiyoType3.Checked = True
        ElseIf dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SAIJI Then
            Me.rdoRiyoType4.Checked = True
        End If
        Me.txtTeiin.Text = dataEXTB0102.PropStrTeiin
        'ワンドリンク
        If dataEXTB0102.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_ARI Then
            Me.rdoOneDrink1.Checked = True
        ElseIf dataEXTB0102.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_NASHI Then
            Me.rdoOneDrink2.Checked = True
        End If
        '催事分類
        If dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MUSIC Then
            Me.rdoSaijiBunrui1.Checked = True
        ElseIf dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEKI Then
            Me.rdoSaijiBunrui2.Checked = True
        ElseIf dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEI Then
            Me.rdoSaijiBunrui3.Checked = True
        ElseIf dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_BUSINESS Then
            Me.rdoSaijiBunrui4.Checked = True
        ElseIf dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MOVIE Then
            Me.rdoSaijiBunrui5.Checked = True
        ElseIf dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ETC Then
            Me.rdoSaijiBunrui6.Checked = True
        End If
        '利用者情報
        Me.lblRiyoshaCd.Text = dataEXTB0102.PropStrRiyoshaCd
        Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTB0102.PropStrRiyoLvl)
        Me.txtRiyoshaNmKana.Text = dataEXTB0102.PropStrRiyoKana
        Me.txtRiyoshaNm.Text = dataEXTB0102.PropStrRiyoNm
        Me.txtDaihyoNm.Text = dataEXTB0102.PropStrDaihyoNm
        Me.txtSekininBushoNm.Text = dataEXTB0102.PropStrSekininBushoNm
        '責任者名コンボボックス作成
        cmbSekininNm.Items.Clear()                            'コンボボックス初期化  2016.11.04 m.hayabuchi ADD
        If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then
            Dim list As ArrayList = logicEXTB0102.GetSekininshaList(dataEXTB0102.PropStrRiyoshaCd)
            For i = 0 To list.Count - 1
                cmbSekininNm.Items.Add(list(i))
            Next
        End If
        Me.cmbSekininNm.Text = dataEXTB0102.PropStrSekininNm
        Me.txtSekininMail.Text = dataEXTB0102.PropStrSekininMail
        Me.lblExasAiteNm.Text = dataEXTB0102.PropStrAiteNm
        Me.lblExasAite.Text = dataEXTB0102.PropStrAiteCd
        Me.txtRiyoPost1.Text = dataEXTB0102.PropStrRiyoYubin1
        Me.txtRiyoPost2.Text = dataEXTB0102.PropStrRiyoYubin2
        Me.cmbRiyoTodo.Text = dataEXTB0102.PropStrRiyoTodo
        Me.txtRiyoShiku.Text = dataEXTB0102.PropStrRiyoShiku
        Me.txtRiyoBan.Text = dataEXTB0102.PropStrRiyoBan
        Me.txtRiyoBuild.Text = dataEXTB0102.PropStrRiyoBuild
        Me.txtRiyoTel1.Text = dataEXTB0102.PropStrRiyoTel11
        Me.txtRiyoTel2.Text = dataEXTB0102.PropStrRiyoTel12
        Me.txtRiyoTel3.Text = dataEXTB0102.PropStrRiyoTel13
        Me.txtRiyoNaisen.Text = dataEXTB0102.PropStrRiyoNaisen
        Me.txtRiyoMobileTel1.Text = dataEXTB0102.PropStrRiyoTel21
        'Me.txtRiyoMobileTel2.Text = dataEXTB0102.PropStrRiyoTel22 '2016.11.2 m.hayabuchi DEL 課題No.59
        'Me.txtRiyoMobileTel3.Text = dataEXTB0102.PropStrRiyoTel23 '2016.11.2 m.hayabuchi DEL 課題No.59
        Me.txtRiyoFax1.Text = dataEXTB0102.PropStrRiyoFax11
        Me.txtRiyoFax2.Text = dataEXTB0102.PropStrRiyoFax12
        Me.txtRiyoFax3.Text = dataEXTB0102.PropStrRiyoFax13
        'チケット
        If dataEXTB0102.PropStrTicketEnterKbn = CommonDeclareEXT.TICKET_ENTER_YURYO Then
            Me.rdoTicket1.Checked = True
        ElseIf dataEXTB0102.PropStrTicketEnterKbn = CommonDeclareEXT.TICKET_ENTER_MURYO Then
            Me.rdoTicket2.Checked = True
        End If
        If dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_KOMI Then
            Me.rdoTicketDrink1.Checked = True
        ElseIf dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_BETSU Then
            Me.rdoTicketDrink2.Checked = True
        ElseIf dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_NASHI Then
            Me.rdoTicketDrink3.Checked = True
        End If
        'HP
        If dataEXTB0102.PropStrHpKeisai = CommonDeclareEXT.HP_OK Then
            Me.rdoEtr2.Checked = True
        ElseIf dataEXTB0102.PropStrHpKeisai = CommonDeclareEXT.HP_NG Then
            Me.rdoEtr1.Checked = True
        End If
        Me.dtpInfoDisc.txtDate.Text = dataEXTB0102.PropStrJohoKokaiDt
        Me.txtDiscTime1.Text = dataEXTB0102.PropStrJohoKokaiTime
        Me.dtpDisc.txtDate.Text = dataEXTB0102.PropStrKokaiDt
        Me.txtDiscTime2.Text = dataEXTB0102.PropStrKokaiTime
        '付帯
        If dataEXTB0102.PropStrFinputSts = CommonDeclareEXT.F_INPUT_KAKUTEI Then
            Me.rdoFutaiSts2.Checked = True
        Else
            Me.rdoFutaiSts1.Checked = True
        End If
        '特記事項
        Me.txtBiko.Text = dataEXTB0102.PropStrBiko
        Me.txtComment.Text = dataEXTB0102.PropStrRiyoCom

        Me.btnCancel.Enabled = True

        Me.pnlSaijiBunrui.SuspendLayout()
        Me.pnlOneDrink.SuspendLayout()
        Me.pnlRiyoType.SuspendLayout()
        Me.pnlKashi.SuspendLayout()

        'スクロールバーを必要な場合のみ表示させます
        'fpRiyobi.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                              ' 2015.12.21 UPD h.hagiwara
        fpRiyobi.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                            ' 2015.12.21 UPD h.hagiwara
        fpRiyobi.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        'ボタン設定
        statusBtnSetting()

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
    ''' 【画面表示用】利用日時リストの値を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="dataList"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadDataRiyobi(ByVal dataList As ArrayList)
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        sheet.RowCount = dataList.Count
        sheet.ColumnCount = 10
        Dim index As New Integer
        Dim lineCnt As New Integer
        index = 0
        lineCnt = 1
        Dim gcCellTime As FarPoint.Win.Spread.Cell
        '時間設定
        Dim maskcell As New FarPoint.Win.Spread.CellType.MaskCellType()
        maskcell.Mask = "99:99"
        maskcell.NullDisplay = "     "
        'Cell設定
        Dim numCellTanka As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellTanka.MaximumValue = 99999999
        numCellTanka.ShowSeparator = True
        numCellTanka.DecimalPlaces = 0
        Dim numCellBai As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellBai.MaximumValue = 999
        numCellBai.ShowSeparator = True
        numCellBai.DecimalPlaces = 2
        Dim numCellSu As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellSu.MaximumValue = 999
        numCellSu.ShowSeparator = True
        numCellSu.DecimalPlaces = 0
        Dim numCellRyokin As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellRyokin.MaximumValue = 9999999999
        numCellRyokin.ShowSeparator = True
        numCellRyokin.DecimalPlaces = 0

        '参照ボタン
        Dim btnTankaDisp As New ButtonCellType With {.Text = "参照"}
        For Each dataRiyobi As CommonDataRiyobi In dataList
            If dataRiyobi.PropBlnSelect Is Nothing = False And dataRiyobi.PropBlnSelect = True Then
                sheet.Cells(index, 0).Value = True
            Else
                sheet.Cells(index, 0).Value = False
            End If
            gcCellTime = sheet.Cells(index, 1)
            gcCellTime.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
            gcCellTime.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
            sheet.Cells(index, 1).Value = lineCnt.ToString
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = dataRiyobi.PropStrYoyakuDtDisp
            sheet.Cells(index, 2).Locked = True
            sheet.Cells(index, 3).CellType = maskcell
            sheet.Cells(index, 4).CellType = maskcell
            If dataRiyobi.PropStrMiteiFlg = "0" Then
                sheet.Cells(index, 3).Value = dataRiyobi.PropStrStartTime
                sheet.Cells(index, 4).Value = dataRiyobi.PropStrEndTime
            Else
                sheet.Cells(index, 3).Value = Nothing
                sheet.Cells(index, 4).Value = Nothing
            End If
            sheet.Cells(index, 5).Value = dataRiyobi.PropIntTanka
            sheet.Cells(index, 5).CellType = numCellTanka
            sheet.Cells(index, 6).CellType = btnTankaDisp
            sheet.Cells(index, 7).Value = dataRiyobi.PropDblBairitu
            sheet.Cells(index, 7).CellType = numCellBai
            If dataRiyobi.PropDblBairitu Is Nothing Or dataRiyobi.PropDblBairitu = 0 Then
                sheet.Cells(index, 7).Value = 1
            End If
            sheet.Cells(index, 8).Value = dataRiyobi.PropIntSu
            sheet.Cells(index, 8).CellType = numCellSu
            sheet.Cells(index, 9).Value = dataRiyobi.PropIntRiyoKin
            sheet.Cells(index, 9).CellType = numCellRyokin

            index = index + 1
            lineCnt = lineCnt + 1
        Next
    End Sub

    ''' <summary>
    ''' 【画面→DB用】利用日SPREADをDataに設定する
    ''' </summary>
    ''' <param name="dataList"></param>
    ''' <remarks></remarks>
    Private Sub SetRiyobiSpreadData(ByRef dataList As ArrayList)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        sheet.RowCount = dataList.Count
        sheet.ColumnCount = 10
        Dim index As New Integer
        Dim lineCnt As New Integer

        index = 0
        lineCnt = 1
        For Each dataRiyobi As CommonDataRiyobi In dataList
            dataRiyobi.PropBlnSelect = sheet.Cells(index, 0).Value
            dataRiyobi.PropStrStartTime = sheet.Cells(index, 3).Value
            dataRiyobi.PropStrEndTime = sheet.Cells(index, 4).Value
            ' 2015.12.08 UPD START↓ h.hagiwara 時間未定も更新可能対応
            'dataRiyobi.PropStrMiteiFlg = "0"
            If sheet.Cells(index, 3).Value <> "" And sheet.Cells(index, 4).Value <> "" Then
                dataRiyobi.PropStrMiteiFlg = "0"
            Else
                dataRiyobi.PropStrMiteiFlg = "1"
            End If
            ' 2015.12.08 UPD END↑ h.hagiwara 時間未定も更新可能対応
            dataRiyobi.PropIntTanka = sheet.Cells(index, 5).Value
            dataRiyobi.PropDblBairitu = sheet.Cells(index, 7).Value
            dataRiyobi.PropIntSu = sheet.Cells(index, 8).Value
            dataRiyobi.PropIntRiyoKin = sheet.Cells(index, 9).Value
            index = index + 1
            lineCnt = lineCnt + 1
        Next
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

        Dim frm As New EXTZ0202
        SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
        '「利用日追加」画面を表示
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        Dim rowCntBefore As Integer = sheet.RowCount '件数カウントBefore
        frm.PropStrTorokuKbn = TOUROKU_KBN_SEISHIKI
        frm.PropLstRiyobi = dataEXTB0102.PropListRiyobi
        frm.PropStrYoyakuNo = dataEXTB0102.PropStrYoyakuNo
        frm.PropStrShisetuKbn = SHISETU_KBN_THEATER
        frm.PropStrStudioKbn = STUDIO_MITEI
        frm.ShowDialog()

        Dim comp As New CommonDataRiyobiCompareter()
        dataEXTB0102.PropListRiyobi.Sort(comp)
        SetSpreadDataRiyobi(dataEXTB0102.PropListRiyobi)
        Dim rowCntAfter As Integer = sheet.RowCount '件数カウントAfter
        If rowCntAfter <> rowCntBefore = True Then
            '請求情報変更フラグ設定
            editFlgSetting()
        End If
    End Sub

    ''' <summary>
    ''' 利用日削除ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnYoyakuDel_Click(sender As Object, e As EventArgs) Handles btnYoyakuDel.Click
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        Dim i As Integer = 0
        Dim rowCnt As Integer = sheet.RowCount
        Dim lstRiyobi As ArrayList = dataEXTB0102.PropListRiyobi
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
        SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
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
                '請求情報変更フラグ設定
                editFlgSetting()
                j = j + 1
            End If
            i = i + 1
        Loop
        SetSpreadDataRiyobi(dataEXTB0102.PropListRiyobi)
    End Sub

    ''' <summary>
    ''' SPREAD利用日時一覧、利用料金自動計算
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub ChangeEventHandler(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fpRiyobi.Change
        ' 2015.12.21 UPD START↓ h.hagiwara
        'If 5 <> e.Column And 7 <> e.Column Then
        '    Exit Sub
        'End If
        If 5 <> e.Column And 7 <> e.Column And 9 <> e.Column Then
            Exit Sub
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        Dim tanka As Integer
        Dim bairitu As Double
        Dim ryokin As Decimal
        If commonValidate.IsHalfNmb(sheet.Cells(e.Row, 5).Value) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0003, "単価"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If commonValidate.IsHalfChar(sheet.Cells(e.Row, 7).Value) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0004, "倍率"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            Dim dVal As Double
            If Double.TryParse(sheet.Cells(e.Row, 7).Value, dVal) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0004, "倍率"), MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
        End If
        If 9 <> e.Column Then                                        ' 2015.12.21 ADD h.hagiwara
            tanka = sheet.Cells(e.Row, 5).Value
            bairitu = sheet.Cells(e.Row, 7).Value
            ryokin = tanka * bairitu
            'sheet.Cells(e.Row, 9).Value = Integer.Parse(Math.Round(ryokin))                      ' 2015.12.21 UPD h.hagiwara
            sheet.Cells(e.Row, 9).Value = Long.Parse(Math.Round(ryokin))                          ' 2015.12.21 UPD h.hagiwara
        End If                                                      ' 2015.12.21 ADD h.hagiwara
        '請求情報変更フラグ設定
        editFlgSetting()
    End Sub


    ''' <summary>
    ''' スプレッド利用日時一覧のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnTankaDisp_Click(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fpRiyobi.ButtonClicked
        '参照ボタンが押された場合
        If e.Column = FB_RIYOBI_DISP Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet                    ' 2015.12.15 ADD h.hagiwara

            ' 2015.12.16 ADD START↓ h.hagiwara 遷移前マスタ存在チェック追加
            Dim strCheck As Date = sheet.Cells(e.Row, 2).Value
            dataEXTB0103.PropStrClickedRiyobi = strCheck.ToString(CommonDeclareEXT.FMT_DATE)
            If logicEXTB0103.GetRyokinBairituMst(dataEXTB0103) = False Then
                'メッセージを出力 
                MsgBox(E2047, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            ' 2015.12.16 ADD END↑ h.hagiwara 遷移前マスタ存在チェック追加

            Dim frm As New EXTZ0212
            SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
            '「利用日追加」画面を表示
            'Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet                  ' 2015.12.15 DEL h.hagiwara
            frm.dataEXTZ0212.PropLstRiyobi = dataEXTB0102.PropListRiyobi
            frm.dataEXTZ0212.PropStrShisetu = CommonDeclareEXT.SHISETU_KBN_THEATER
            '押下行日付、選択行日付の設定
            Dim dt As Date = sheet.Cells(e.Row, 2).Value
            frm.dataEXTZ0212.PropStrSelectRiyobi = dt.ToString(CommonDeclareEXT.FMT_DATE)
            frm.ShowDialog()
            SetSpreadDataRiyobi(dataEXTB0102.PropListRiyobi)
            '請求情報変更フラグ設定
            If frm.dataEXTZ0212.PropBlnChangeFlg = True Then
                editFlgSetting()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 【画面表示用】確認依頼履歴の値を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadIraiRireki(ByVal table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbIrai.ActiveSheet
        sheet.RowCount = table.Rows.Count
        sheet.ColumnCount = 3
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("user_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("irai_dt")
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = row("com")
            sheet.Cells(index, 2).Locked = True
            index = index + 1
        Loop

        'スクロールバーを必要な場合のみ表示させます
        fbIrai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbIrai.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 【画面表示用】確認記録リストの値を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadIraiKakunin(ByVal table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbRecord.ActiveSheet
        sheet.RowCount = table.Rows.Count
        sheet.ColumnCount = 4
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("user_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("check_dt")
            sheet.Cells(index, 1).Locked = True
            If "1" = row("check_sts") Then
                sheet.Cells(index, 2).Value = "承認"
            ElseIf "2" = row("check_sts") Then
                sheet.Cells(index, 2).Value = "差し戻し"
            End If
            sheet.Cells(index, 2).Locked = True
            sheet.Cells(index, 3).Value = row("com")
            sheet.Cells(index, 3).Locked = True
            index = index + 1
        Loop

        'スクロールバーを必要な場合のみ表示させます
        fbRecord.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbRecord.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 【画面表示用】タイムスケジュールを画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadTimeSchedule(ByVal table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTimeSchedule.ActiveSheet
        Dim index As New Integer
        Dim row As DataRow
        Dim btnDel As New ButtonCellType With {.Text = "削除"}         ' 2015.12.03 ADD h.hagiwara

        sheet.RowCount = table.Rows.Count
        'sheet.ColumnCount = 3                       ' 2015.12.03 UPD h.hagiwara
        sheet.ColumnCount = 4                        ' 2015.12.03 UPD h.hagiwara
        index = 0

        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("shurui")
            sheet.Cells(index, 1).Value = row("time")
            sheet.Cells(index, 2).Value = row("biko")
            sheet.Cells(index, 3).CellType = btnDel                   ' 2015.12.03 ADD h.hagiwara
            index = index + 1
        Loop

        'スクロールバーを必要な場合のみ表示させます
        fbTimeSchedule.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbTimeSchedule.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 【DataSet格納】画面入力したタイムスケジュールを格納する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetTimeScheduleSpread(ByRef table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTimeSchedule.ActiveSheet
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            row("shurui") = sheet.Cells(index, 0).Value
            row("time") = sheet.Cells(index, 1).Value
            row("biko") = sheet.Cells(index, 2).Value
            index = index + 1
        Loop
    End Sub

    ''' <summary>
    ''' 【画面表示用】チケット情報を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadTicket(ByVal table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTicket.ActiveSheet
        Dim index As New Integer
        Dim row As DataRow
        Dim btnDel As New ButtonCellType With {.Text = "削除"}         ' 2015.12.03 ADD h.hagiwara
        Dim numCellRyokin As New FarPoint.Win.Spread.CellType.NumberCellType()                  ' 2015.12.08 ADD h.hagiwara
        'numCellRyokin.MaximumValue = 9999999999                                                ' 2015.12.08 ADD h.hagiwara
        numCellRyokin.MaximumValue = 99999999                                                   ' 2015.12.18 UPD h.hagiwara
        numCellRyokin.ShowSeparator = True                                                      ' 2015.12.08 ADD h.hagiwara
        numCellRyokin.DecimalPlaces = 0                                                         ' 2015.12.08 ADD h.hagiwara

        sheet.RowCount = table.Rows.Count
        'sheet.ColumnCount = 2                        ' 2015.12.03 UPD h.hagiwara
        sheet.ColumnCount = 3                         ' 2015.12.03 UPD h.hagiwara
        index = 0

        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("ticket_nm")
            sheet.Cells(index, 1).Value = commonLogicEXT.DbNullToNothing(row, "ticket_kin")
            sheet.Cells(index, 1).CellType = numCellRyokin                                     ' 2015.12.08 ADD h.hagiwara
            sheet.Cells(index, 2).CellType = btnDel                                            ' 2015.12.03 ADD h.hagiwara
            index = index + 1
        Loop

        'スクロールバーを必要な場合のみ表示させます
        fbTicket.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbTicket.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 【DataSet格納】画面入力したチケット情報を格納する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub SetTicketSpread(ByRef table As DataTable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTicket.ActiveSheet
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            row("ticket_nm") = sheet.Cells(index, 0).Value
            row("ticket_kin") = commonLogicEXT.convInsNum(sheet.Cells(index, 1).Value)
            index = index + 1
        Loop
    End Sub

    ''' <summary>
    ''' 【画面表示用】付帯情報を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="ht"></param>
    ''' <remarks></remarks>
    Private Sub SetSpreadFutai(ByVal ht As Hashtable)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        sheet.RowCount = ht.Count
        sheet.ColumnCount = 7
        sheet.Columns(6).Visible = False
        'lock
        sheet.Columns(0).Locked = True
        sheet.Columns(3).Locked = True
        sheet.Columns(4).Locked = True
        Dim index As New Integer
        Dim row As DataRow
        Dim table As DataTable
        index = 0
        Dim btnEdit As New ButtonCellType With {.Text = "編集"}
        Dim btnCopy As New ButtonCellType With {.Text = "コピー"}
        Dim btnPrint As New ButtonCellType With {.Text = "印刷"}
        Dim dt As DateTime
        For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
            table = ht(dataRiyobi.PropStrYoyakuDt)
            row = table.Rows(0)
            dt = row("yoyaku_dt")
            sheet.Cells(index, 0).Value = dt.ToString(CommonDeclareEXT.FMT_DATE_DISP)
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).CellType = btnEdit
            sheet.Cells(index, 2).CellType = btnCopy
            sheet.Cells(index, 3).Value = row("fuzoku_nm")
            sheet.Cells(index, 4).Value = row("total_fuzoku_kin")
            sheet.Cells(index, 5).CellType = btnPrint
            sheet.Cells(index, 6).Value = dt.ToString(CommonDeclareEXT.FMT_DATE)
            index = index + 1
        Next

        'スクロールバーを必要な場合のみ表示させます
        fbFutai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbFutai.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' SPREAD付帯一覧のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFutai_Click(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbFutai.ButtonClicked
        If e.Column = FB_FUTAI_EDIT Then
            'SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「付帯入力」画面を表示
            Dim frm As New EXTZ0209
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet

            ' 2015.12.16 ADD START↓ h.hagiwara 遷移前マスタ存在チェック追加
            Dim strCheck As Date = sheet.Cells(e.Row, 0).Value
            dataEXTB0103.PropStrClickedRiyobi = strCheck.ToString(CommonDeclareEXT.FMT_DATE)
            If logicEXTB0103.GetFutaiSetubiInfMst(dataEXTB0103) = False Then
                'メッセージを出力 
                MsgBox(E2048, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            ' 2015.12.16 ADD END↑ h.hagiwara 遷移前マスタ存在チェック追加

            frm.dataEXTZ0209.PropStrYoyakuNo = dataEXTB0102.PropStrYoyakuNo
            'frm.dataEXTZ0209.PropStrSaijiNm = dataEXTB0102.PropStrSaijiNm                     ' 2015.12.08 UPD h.hagiwara 
            frm.dataEXTZ0209.PropStrSaijiNm = Me.txtSaiji.Text                                 ' 2015.12.08 UPD h.hagiwara 
            frm.dataEXTZ0209.PropStrShisetu = CommonDeclareEXT.SHISETU_KBN_THEATER
            '押下行日付、選択行日付の設定
            frm.dataEXTZ0209.PropStrRiyobiDisp = sheet.Cells(e.Row, 0).Value
            frm.dataEXTZ0209.PropStrRiyobi = sheet.Cells(e.Row, 6).Value
            frm.dataEXTZ0209.PropFutaiDetailTable = dataEXTB0102.PropHtFutaiDetail(sheet.Cells(e.Row, 6).Value)

            ' 2015.12.21 ADD START↓ h.hagiwara
            Dim lngAll As Long = 0
            Dim tblTemp As DataTable
            For i = 0 To sheet.RowCount - 1
                If i <> e.Row Then
                    tblTemp = dataEXTB0102.PropHtFutai(sheet.Cells(i, 6).Value)
                    lngAll += tblTemp.Rows(0)("total_fuzoku_kin")
                End If
            Next
            frm.dataEXTZ0209.PropLngTotal = lngAll
            ' 2015.12.21 ADD END↑ h.hagiwara

            frm.ShowDialog()
            If frm.PropBlnChangeFlg Then
                Dim futaiTable As DataTable = dataEXTB0102.PropHtFutai(frm.dataEXTZ0209.PropStrRiyobi)
                futaiTable.Rows(0)("total_fuzoku_kin") = frm.dataEXTZ0209.PropIntTotal
                futaiTable.Rows(0)("fuzoku_nm") = frm.dataEXTZ0209.PropStrFutaiTotalNm
                ' 2015.11.13 ADD START↓ h.hagiwara
                If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
                    futaiTable.Rows(0)("tax_kin") = 0
                Else
                    futaiTable.Rows(0)("tax_kin") = frm.dataEXTZ0209.PropIntTotalTax
                End If
                ' 2015.11.13 ADD END↑ h.hagiwara
                dataEXTB0102.PropHtFutaiDetail(sheet.Cells(e.Row, 6).Value) = frm.dataEXTZ0209.PropFutaiDetailTable
                strFutaiFlg = "1"
                '請求情報変更フラグ設定
                editFlgSetting()
                SetSpreadFutai(dataEXTB0102.PropHtFutai)
            End If
        ElseIf e.Column = FB_FUTAI_COPY Then
            Dim frm As New EXTZ0210
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
            If dataEXTB0102.PropHtFutaiDetail.Count = 1 Then
                MsgBox("コピー対象日がありません")
                Exit Sub
            End If

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim aryRiyobi(dataEXTB0102.PropHtFutaiDetail.Count - 2) As String
            Dim aryIndex As Integer = 0
            frm.PropStrSelectedDateDisp = sheet.Cells(e.Row, 0).Value
            frm.PropStrSelectedDate = sheet.Cells(e.Row, 6).Value
            For Each key As String In dataEXTB0102.PropHtFutaiDetail.Keys
                If frm.PropStrSelectedDate <> key Then
                    aryRiyobi(aryIndex) = key
                    aryIndex = aryIndex + 1
                End If
            Next
            frm.PropAryRiyobi = aryRiyobi
            '「付帯コピー」画面を表示
            frm.ShowDialog()
            'コピー処理
            If frm.PropBlnChangeFlg Then
                Dim dtCopyFutaiHeader As DataTable
                Dim dtCopyFutaiDetail As DataTable
                For Each tgDate In frm.PropLstSelectedRiyobi
                    dtCopyFutaiHeader = dataEXTB0102.PropHtFutai(frm.PropStrSelectedDate).Copy()
                    dtCopyFutaiDetail = dataEXTB0102.PropHtFutaiDetail(frm.PropStrSelectedDate).Copy()
                    Dim fromRow As DataRow = dtCopyFutaiHeader.Rows(0)
                    Dim toRow As DataRow = dataEXTB0102.PropHtFutai(tgDate).Rows(0)
                    toRow("fuzoku_nm") = fromRow("fuzoku_nm")
                    toRow("total_fuzoku_kin") = fromRow("total_fuzoku_kin")
                    toRow("tax_kin") = fromRow("tax_kin")                                                         ' 2015.11.13 ADD h.hagiwara
                    For i = 0 To dtCopyFutaiDetail.Rows.Count - 1
                        dtCopyFutaiDetail.Rows(i).Item("yoyaku_dt") = tgDate
                    Next
                    dataEXTB0102.PropHtFutaiDetail(tgDate) = dtCopyFutaiDetail
                Next
                strFutaiFlg = "1"
                '請求情報変更フラグ設定
                editFlgSetting()
                SetSpreadFutai(dataEXTB0102.PropHtFutai)
            End If
        ElseIf e.Column = FB_FUTAI_PRINT Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            dataEXTB0103.PropStrClickedIncident = Me.fbFutai.ActiveSheet.Cells(e.Row, 0).Value
            '予約番号
            dataEXTB0103.PropStrReserveNo_Output = Me.lblYoyakuNo.Text
            '日付
            dataEXTB0103.PropStrCalculateDay_Output = Me.fpRiyobi.ActiveSheet.Cells(0, 2).Value
            dataEXTB0103.PropStrCalculateDay_Output = dataEXTB0103.PropStrCalculateDay_Output.Substring(0, 10)
            'クリックされた日付
            dataEXTB0103.PropStrUseDays_Output = Me.fbFutai.ActiveSheet.Cells(e.Row, 0).Value
            dataEXTB0103.PropStrUseDays_Output = dataEXTB0103.PropStrUseDays_Output.Substring(0, 10)
            '貸出区分
            dataEXTB0103.PropStrRentalClass_Output = dataEXTB0102.PropStrKashiKind
            If MsgBox(EXT_C0009, MsgBoxStyle.OkCancel, TITLE_INFO) = vbOK Then
                If logicEXTB0103.PrintUseDetailsMain(dataEXTB0103) = False Then
                    MsgBox(puErrMsg)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 【画面表示用】請求情報を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function SetSpreadBillPay(ByVal table As DataTable) As Boolean
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        sheet.RowCount = table.Rows.Count
        sheet.ColumnCount = 25
        'Dim total As New Integer                           ' 2015.12.21 UPD h.hagiwara
        Dim total As New Long                               ' 2015.12.21 UPD h.hagiwara
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Dim btnSeikyu As New ButtonCellType With {.Text = "入力"}
        Dim btnNyukin As New ButtonCellType With {.Text = "入力"}
        'lock
        sheet.Columns(1).Locked = True
        sheet.Columns(2).Locked = True
        sheet.Columns(3).Locked = True
        sheet.Columns(5).Locked = True
        sheet.Columns(6).Locked = True
        sheet.Columns(7).Locked = True
        sheet.Columns(8).Locked = True
        sheet.Columns(9).Locked = True
        sheet.Columns(10).Locked = True
        sheet.Columns(11).Locked = True
        sheet.Columns(12).Locked = True
        sheet.Columns(14).Locked = True
        '非表示カラム設定
        sheet.Columns(15).Visible = False
        sheet.Columns(16).Visible = False
        sheet.Columns(17).Visible = False
        sheet.Columns(18).Visible = False
        sheet.Columns(19).Visible = False
        sheet.Columns(20).Visible = False
        sheet.Columns(21).Visible = False
        sheet.Columns(22).Visible = False
        sheet.Columns(23).Visible = False
        sheet.Columns(24).Visible = False

        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = False                                             ' 2015.12.16 UPD h.hagiwara
            If row("seikyu_input_flg") = "1" Then '請求入力
                sheet.Cells(index, 1).Value = "○"
            Else
                sheet.Cells(index, 1).Value = ""
            End If
            If row("seikyu_irai_flg") = "1" Then 'EXAS請求依頼
                sheet.Cells(index, 2).Value = "○"
            Else
                sheet.Cells(index, 2).Value = ""
            End If
            If row("nyukin_input_flg") = "1" Then '入金入力
                sheet.Cells(index, 3).Value = "○"
            Else
                sheet.Cells(index, 3).Value = ""
            End If
            sheet.Cells(index, 4).CellType = btnSeikyu
            sheet.Cells(index, 5).Value = row("kakutei_kin")
            sheet.Cells(index, 6).Value = row("chosei_kin")
            sheet.Cells(index, 7).Value = row("tax_kin")
            row("seikyu_kin") = row("kakutei_kin") + row("chosei_kin") + row("tax_kin")                       ' 2016.01.26 ADD h.hagiwara
            sheet.Cells(index, 8).Value = row("seikyu_kin")
            total = total + row("seikyu_kin")
            sheet.Cells(index, 9).Value = row("seikyu_dt")
            sheet.Cells(index, 10).Value = row("seikyu_irai_no")
            sheet.Cells(index, 11).Value = row("nyukin_yotei_dt")
            sheet.Cells(index, 12).Value = row("seikyu_naiyo_nm")
            sheet.Cells(index, 13).CellType = btnNyukin
            sheet.Cells(index, 14).Value = row("nyukin_dt")
            '非表示項目
            sheet.Cells(index, 15).Value = row("seq")
            sheet.Cells(index, 16).Value = row("shokei")
            sheet.Cells(index, 17).Value = row("nyukin_kbn")
            sheet.Cells(index, 18).Value = row("nyukin_kin")
            sheet.Cells(index, 19).Value = row("nyukin_link_no")
            sheet.Cells(index, 20).Value = row("aite_cd")
            sheet.Cells(index, 21).Value = row("seikyu_naiyo")
            sheet.Cells(index, 22).Value = row("seikyu_title1")
            sheet.Cells(index, 23).Value = row("seikyu_title2")
            sheet.Cells(index, 24).Value = row("aite_nm")
            index = index + 1
        Loop
        dataEXTB0102.propIntTtl = total
        'スクロールバーを必要な場合のみ表示させます
        fbBillPay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbBillPay.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        If table.Rows.Count = 0 Then
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 請求情報の初期表示行を作成
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Sub InitBillPay(ByRef table As DataTable)
        'SPREAD
        Dim index As New Integer
        Dim btnSeikyu As New ButtonCellType With {.Text = "入力"}
        Dim btnNyukin As New ButtonCellType With {.Text = "入力"}
        Dim row1 As DataRow = table.NewRow
        Dim row2 As DataRow = table.NewRow
        Dim dtSeikyu As DateTime = System.DateTime.Now
        Dim dtNyukin1 As DateTime = DateTime.Parse(dtSeikyu.ToString("yyyy/MM/") & "01")
        ' Dim dtNyukin As DateTime                        ' 2016.03.03 DEL h.hagiwara
        ' Dim dtNyukinwork As DateTime                    ' 2016.03.03 DEL h.hagiwara               
        Dim title1 As String = ""
        ' 2015.12.08 UPD START↓ h.hagiwara
        'Dim title2 As String = dataEXTB0102.PropStrSaijiNm   
        Dim title2 As String = ""
        If dataEXTB0102.PropStrSaijiNm.ToString.Length > 25 Then
            title2 = dataEXTB0102.PropStrSaijiNm.ToString.Substring(0, 25)
        Else
            title2 = dataEXTB0102.PropStrSaijiNm
        End If
        ' 2015.12.08 UPD END↑ h.hagiwara
        '利用料
        'Dim kakuteiKin As Integer = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102)                  ' 2015.12.21 UPD h.hagiwara
        Dim kakuteiKin As Long = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102)                      ' 2015.12.21 UPD h.hagiwara
        'Dim taxKin As Integer = Integer.Parse(Math.Round((kakuteiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))           ' 2015.12.21 UPD h.hagiwara
        Dim taxKin As Long = Integer.Parse(Math.Round((kakuteiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))               ' 2015.12.21 UPD h.hagiwara
        'Dim seikyuKin As Integer = kakuteiKin + taxKin                                           ' 2015.12.21 UPD h.hagiwara
        Dim seikyuKin As Long = kakuteiKin + taxKin                                               ' 2015.12.21 UPD h.hagiwara

        table.Rows.Clear()                                                              ' 2015.12.18 ADD h.hagiwara
        ' 2016.03.03 DEL START↓ h.hagiwara
        '入金予定日の取得
        'dtNyukinwork = dtNyukin1.AddMonths(2)
        'dtNyukinwork = dtNyukinwork.AddDays(-1)
        'dtNyukin = logicEXTB0103.WeekDayCheck(dtNyukinwork)
        ' 2016.03.03 DEL END↑ h.hagiwara

        row1("yoyaku_no") = DBNull.Value
        row1("seq") = DBNull.Value
        row1("seikyu_input_flg") = "0"
        row1("nyukin_input_flg") = "0"
        row1("seikyu_irai_flg") = "0"
        row1("kakutei_kin") = kakuteiKin
        row1("chosei_kin") = "0"
        row1("tax_kin") = taxKin
        row1("seikyu_kin") = seikyuKin
        'row1("seikyu_dt") = dtSeikyu.ToString(CommonDeclareEXT.FMT_DATE)                               ' 2016.03.03 UPD h.hagiwara
        row1("seikyu_dt") = Nothing                                                                     ' 2016.03.03 UPD h.hagiwara
        row1("seikyu_irai_no") = DBNull.Value
        'row1("nyukin_yotei_dt") = dtNyukin.ToString(CommonDeclareEXT.FMT_DATE)                         ' 2016.03.03 UPD h.hagiwara
        row1("nyukin_yotei_dt") = Nothing                                                               ' 2016.03.03 UPD h.hagiwara
        row1("seikyu_naiyo_nm") = "利用料"
        row1("nyukin_dt") = DBNull.Value
        row1("seq") = DBNull.Value
        row1("shokei") = kakuteiKin
        row1("nyukin_kbn") = DBNull.Value
        row1("nyukin_kin") = DBNull.Value
        row1("nyukin_link_no") = DBNull.Value
        row1("aite_cd") = dataEXTB0102.PropStrAiteCd
        row1("aite_nm") = dataEXTB0102.PropStrAiteNm
        row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO
        logicEXTB0103.SetSeikyuTitle(dataEXTB0102, title1, SEIKYU_NAIYOU_RIYO)
        row1("seikyu_title1") = title1
        row1("seikyu_title2") = title2
        row1("get_seikyu_input_flg") = "0"                                  ' 2016.02.03 ADD h.hagiwara
        '付帯設備
        kakuteiKin = logicEXTB0103.GetFutairyo(dataEXTB0102)
        taxKin = logicEXTB0103.GetFutairyoTax(dataEXTB0102)                 ' 2015.11.13 UPD h.hagiwara
        seikyuKin = kakuteiKin + taxKin
        title1 = ""
        ' 2015.12.08 UPD START↓ h.hagiwara
        'title2 = dataEXTB0102.PropStrSaijiNm
        If dataEXTB0102.PropStrSaijiNm.ToString.Length > 25 Then
            title2 = dataEXTB0102.PropStrSaijiNm.ToString.Substring(0, 25)
        Else
            title2 = dataEXTB0102.PropStrSaijiNm
        End If
        ' 2015.12.08 UPD END↑ h.hagiwara
        row2("yoyaku_no") = DBNull.Value
        row2("seq") = DBNull.Value
        row2("seikyu_input_flg") = "0"
        row2("nyukin_input_flg") = "0"
        row2("seikyu_irai_flg") = "0"
        row2("kakutei_kin") = kakuteiKin                                 ' 2015.12.21 UPD h.hagiwara
        row2("chosei_kin") = "0"
        row2("tax_kin") = taxKin
        row2("seikyu_kin") = seikyuKin                                   ' 2015.12.21 UPD h.hagiwara
        'row2("seikyu_dt") = dtSeikyu.ToString(CommonDeclareEXT.FMT_DATE)                               ' 2016.03.03 UPD h.hagiwara
        row2("seikyu_dt") = Nothing                                                                     ' 2016.03.03 UPD h.hagiwara
        row2("seikyu_irai_no") = DBNull.Value
        'row2("nyukin_yotei_dt") = dtNyukin.ToString(CommonDeclareEXT.FMT_DATE)                         ' 2016.03.03 UPD h.hagiwara
        row2("nyukin_yotei_dt") = Nothing                                                               ' 2016.03.03 UPD h.hagiwara
        row2("seikyu_naiyo_nm") = "付帯設備"
        row2("nyukin_dt") = DBNull.Value
        row2("seq") = DBNull.Value
        row2("shokei") = kakuteiKin                                      ' 2015.12.21 UPD h.hagiwara
        row2("nyukin_kbn") = DBNull.Value
        row2("nyukin_kin") = DBNull.Value
        row2("nyukin_link_no") = DBNull.Value
        row2("aite_cd") = dataEXTB0102.PropStrAiteCd
        row2("aite_nm") = dataEXTB0102.PropStrAiteNm
        row2("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI
        logicEXTB0103.SetSeikyuTitle(dataEXTB0102, title1, SEIKYU_NAIYOU_FUTAI)
        row2("seikyu_title1") = title1
        row2("seikyu_title2") = title2
        row2("get_seikyu_input_flg") = "0"                                  ' 2016.02.03 ADD h.hagiwara

        '追加
        table.Rows.Add(row1)
        table.Rows.Add(row2)
    End Sub

    ''' <summary>
    ''' 請求情報の再計算
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function CalcBillPay(ByRef table As DataTable) As Boolean
        'Private Sub CalcBillPay(ByRef table As DataTable)                                          ' 2015.12.21 UPD h.hagiwara  
        'SPREAD
        Dim index As New Integer
        Dim btnSeikyu As New ButtonCellType With {.Text = "入力"}
        Dim btnNyukin As New ButtonCellType With {.Text = "入力"}

        'Dim intRiyoTtl As Integer = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102)                             ' 2015.12.21 UPD h.hagiwara
        'Dim intFutaiTtl As Integer = logicEXTB0103.GetFutairyo(dataEXTB0102)                                ' 2015.12.21 UPD h.hagiwara
        Dim intRiyoTtl As Long = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102)                                 ' 2015.12.21 UPD h.hagiwara
        Dim intFutaiTtl As Long = logicEXTB0103.GetFutairyo(dataEXTB0102)                                    ' 2015.12.21 UPD h.hagiwara
        'Dim intRiyoFutaiiTtl As Integer = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102) + logicEXTB0103.GetFutairyo(dataEXTB0102)  ' 2015.11.13 UPD h.hagiwara ' 2015.12.21 UPD h.hagiwara
        Dim intRiyoFutaiiTtl As Long = logicEXTB0103.GetRiyoKinTotal(dataEXTB0102) + logicEXTB0103.GetFutairyo(dataEXTB0102)      ' 2015.12.21 UPD h.hagiwara
        'Dim intFutaiTax As Integer = logicEXTB0103.GetFutairyoTax(dataEXTB0102)                                                   ' 2015.11.13 ADD h.hagiwara 
        Dim intFutaiTax As Long = logicEXTB0103.GetFutairyoTax(dataEXTB0102)                                                       ' 2015.12.21 UPD h.hagiwara 

        ' 2015.12.21 ADD START↓ h.hagiwara 請求金額桁判定
        Dim lngall As Long
        lngall = intRiyoFutaiiTtl + intFutaiTax
        If lngall.ToString.Length > 11 Then
            puErrMsg = String.Format(CommonDeclareEXT.E2049, "予約番号の請求金額")
            Return False
        End If
        ' 2015.12.21 ADD END↑ h.hagiwara 請求金額桁判定

        Dim row As DataRow
        '合計の計算
        '請求入力済の合計値確認
        Do While index < table.Rows.Count
            row = table.Rows(index)
            If row("seikyu_input_flg") = "1" Then
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    intRiyoFutaiiTtl -= row("kakutei_kin")
                ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Then
                    intRiyoTtl -= row("kakutei_kin")
                ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Then
                    intFutaiTtl -= row("kakutei_kin")
                    intFutaiTax -= row("tax_kin")                                                        ' 2015.12.17 ADD h.hagiwara
                End If
            End If
            index += 1
        Loop
        index = 0
        '再計算
        '請求・付帯各1行まで再計算、請求未入力の行がない場合、行を作成する
        Do While index < table.Rows.Count
            row = table.Rows(index)
            If row("seikyu_input_flg") = "1" Then
                '対象外
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                '利用料付帯設備
                ' 2015.12.21 UPD START↓ h.hagiwara
                'Dim kakuteiKin As Integer = intRiyoFutaiiTtl                                 
                'Dim shokeiKin As Integer = kakuteiKin + row("chosei_kin")                   
                'Dim taxKin As Integer = Integer.Parse(Math.Round((shokeiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                'Dim seikyuKin As Integer = shokeiKin + taxKin
                Dim kakuteiKin As Long = intRiyoFutaiiTtl
                Dim shokeiKin As Long = kakuteiKin + row("chosei_kin")
                Dim taxKin As Long = Long.Parse(Math.Round((shokeiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                Dim seikyuKin As Long = shokeiKin + taxKin
                ' 2015.12.21 UPD END↑ h.hagiwara
                row("kakutei_kin") = kakuteiKin
                row("shokei") = shokeiKin
                row("tax_kin") = taxKin
                row("seikyu_kin") = seikyuKin
                intRiyoFutaiiTtl -= kakuteiKin
                intRiyoTtl = 0
                intFutaiTtl = 0
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Then
                '利用料
                ' 2015.12.21 UPD START↓ h.hagiwara
                'Dim kakuteiKin As Integer = intRiyoTtl                                
                'Dim shokeiKin As Integer = kakuteiKin + row("chosei_kin")              
                'Dim taxKin As Integer = Integer.Parse(Math.Round((shokeiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                'Dim seikyuKin As Integer = shokeiKin + taxKin                         
                ' 2015.12.21 UPD END↑ h.hagiwara
                Dim kakuteiKin As Long = intRiyoTtl
                Dim shokeiKin As Long = kakuteiKin + row("chosei_kin")
                Dim taxKin As Long = Long.Parse(Math.Round((shokeiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                Dim seikyuKin As Long = shokeiKin + taxKin
                row("kakutei_kin") = kakuteiKin
                row("shokei") = shokeiKin
                row("tax_kin") = taxKin
                row("seikyu_kin") = seikyuKin
                intRiyoTtl -= kakuteiKin
                intRiyoFutaiiTtl -= kakuteiKin
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Then
                '付帯設備
                ' 2015.12.21 UPD START↓ h.hagiwara
                'Dim kakuteiKin As Integer = intFutaiTtl
                'Dim shokeiKin As Integer = kakuteiKin + row("chosei_kin")
                '' Dim taxKin As Integer = Integer.Parse(Math.Round((shokeiKin * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))   ' 2015.11.13 UPD h.hagiwara
                'Dim taxKin As Integer = intFutaiTax                                                                                         ' 2015.11.13 UPD h.hagiwara
                'Dim seikyuKin As Integer = shokeiKin + taxKin
                Dim kakuteiKin As Long = intFutaiTtl
                Dim shokeiKin As Long = kakuteiKin + row("chosei_kin")
                Dim taxKin As Long = intFutaiTax                                                                                         ' 2015.11.13 UPD h.hagiwara
                Dim seikyuKin As Long = shokeiKin + taxKin
                ' 2015.12.21 UPD END↑ h.hagiwara
                row("kakutei_kin") = kakuteiKin
                row("shokei") = shokeiKin
                row("tax_kin") = taxKin
                row("seikyu_kin") = seikyuKin
                intFutaiTtl -= kakuteiKin
                intRiyoFutaiiTtl -= kakuteiKin
                intFutaiTax -= intFutaiTax                                        ' 2015.12.17 ADD h.hagiwara
            End If
            index += 1
        Loop
        ' 2015.12.16 MOVE START↓ h.hagiwara
        ''不要行削除(正規行は対象としない)
        index = 0
        Dim delCnt As New Integer
        Dim htR As Hashtable = dataEXTB0102.PropHtExasPro
        Dim htF As Hashtable = dataEXTB0102.PropHtExasProFutai
        Do While index < table.Rows.Count
            row = table.Rows(index)
            Dim htKey As String = index.ToString
            If row("seikyu_input_flg") = "1" Then
                '対象外
                index += 1
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI And row("seikyu_kin") = 0 And index > 0 Then
                table.Rows.Remove(row)
                htR.Remove(htKey)
                htEditIndex(htR, index)
                htF.Remove(htKey)
                htEditIndex(htF, index)
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO And row("seikyu_kin") = 0 And index > 0 Then
                table.Rows.Remove(row)
                htEditIndex(htR, index)
            ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI And row("seikyu_kin") = 0 And index > 1 Then
                table.Rows.Remove(row)
                htEditIndex(htF, index)
            Else
                index += 1
            End If
            If index = table.Rows.Count Then
                Exit Do
            End If
        Loop
        ' 2015.12.16 MOVE END↑ h.hagiwara
        '請求依頼作成
        '複製もとを取得
        index = 0
        Dim cpTable As DataTable = table.Copy()
        Dim cpRow As DataRow
        Dim newRow As DataRow
        Dim isSet As Boolean
        If intRiyoFutaiiTtl <> 0 Then
            Do While index < cpTable.Rows.Count
                cpRow = cpTable.Rows(index)
                If cpRow("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    isSet = True
                    Exit Do
                End If
                index += 1
            Loop
            If isSet = False Then
                '該当行なしのためスキップ
            Else
                newRow = table.NewRow()
                cpRow = cpTable.Rows(index)
                newRow("seq") = DBNull.Value
                newRow("seikyu_input_flg") = "0"
                newRow("nyukin_input_flg") = "0"
                newRow("seikyu_irai_flg") = "0"
                newRow("kakutei_kin") = intRiyoFutaiiTtl
                newRow("chosei_kin") = "0"
                newRow("tax_kin") = Integer.Parse(Math.Round((intRiyoFutaiiTtl * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                newRow("seikyu_kin") = newRow("kakutei_kin") + newRow("tax_kin")
                newRow("seikyu_dt") = cpRow("seikyu_dt")
                newRow("seikyu_irai_no") = DBNull.Value
                newRow("nyukin_yotei_dt") = cpRow("nyukin_yotei_dt")
                newRow("seikyu_naiyo_nm") = "利用料+付帯設備"
                newRow("nyukin_dt") = DBNull.Value
                newRow("seq") = DBNull.Value
                newRow("shokei") = intRiyoFutaiiTtl
                newRow("nyukin_kbn") = DBNull.Value
                newRow("nyukin_kin") = DBNull.Value
                newRow("nyukin_link_no") = DBNull.Value
                newRow("aite_cd") = cpRow("aite_cd")
                newRow("aite_nm") = cpRow("aite_nm")
                newRow("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI
                newRow("seikyu_title1") = cpRow("seikyu_title1")
                newRow("seikyu_title2") = cpRow("seikyu_title2")
                newRow("get_seikyu_input_flg") = "0"                                       ' 2016.02.18 ADD h.hagiwara 不具合対応
                table.Rows.Add(newRow)
                'プロジェクト設定
                Dim htKey As String
                If IsDBNull(cpRow("seikyu_irai_no")) Then
                    htKey = index.ToString
                Else
                    htKey = cpRow("seikyu_irai_no")
                End If
                Dim prjTable As DataTable = dataEXTB0102.PropHtExasPro(htKey).Copy()
                Dim prjIndex As New Integer
                Dim prjRow As DataRow
                Do While prjIndex < prjTable.Rows.Count
                    prjRow = prjTable.Rows(prjIndex)
                    prjRow("keijo_kin") = 0
                    prjRow("tax_kin") = 0
                    prjRow("seq") = DBNull.Value
                    prjRow("seikyu_irai_no") = DBNull.Value
                    prjRow("seikyu_irai_seq") = DBNull.Value
                    prjRow("content_cd") = DBNull.Value
                    prjRow("content_uchi_cd") = DBNull.Value
                    prjRow("event_nm") = ""
                    prjRow("content_uchi_nm") = ""
                    prjIndex += 1
                Loop
                dataEXTB0102.PropHtExasPro.Add((table.Rows.Count - 1).ToString, prjTable)
                'プロジェクト設定
                prjTable = dataEXTB0102.PropHtExasProFutai(htKey).Copy()
                prjIndex = 0
                Do While prjIndex < prjTable.Rows.Count
                    prjRow = prjTable.Rows(prjIndex)
                    prjRow("keijo_kin") = 0
                    prjRow("tax_kin") = 0
                    prjRow("seq") = DBNull.Value
                    prjRow("seikyu_irai_no") = DBNull.Value
                    prjRow("seikyu_irai_seq") = DBNull.Value
                    prjRow("content_cd") = DBNull.Value
                    prjRow("content_uchi_cd") = DBNull.Value
                    prjRow("event_nm") = ""
                    prjRow("content_uchi_nm") = ""
                    prjIndex += 1
                Loop
                dataEXTB0102.PropHtExasProFutai.Add((table.Rows.Count - 1).ToString, prjTable)
                intRiyoTtl = 0
                intFutaiTtl = 0
            End If
        End If
        index = 0
        isSet = False
        If intRiyoTtl <> 0 Then
            Do While index < cpTable.Rows.Count
                cpRow = cpTable.Rows(index)
                If cpRow("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Then
                    isSet = True
                    Exit Do
                End If
                index += 1
            Loop
            If isSet = False Then
                '該当行なしのためスキップ
            Else
                newRow = table.NewRow()
                cpRow = cpTable.Rows(index)
                newRow("seq") = DBNull.Value
                newRow("seikyu_input_flg") = "0"
                newRow("nyukin_input_flg") = "0"
                newRow("seikyu_irai_flg") = "0"
                newRow("kakutei_kin") = intRiyoTtl
                newRow("chosei_kin") = "0"
                newRow("tax_kin") = Integer.Parse(Math.Round((intRiyoTtl * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                newRow("seikyu_kin") = newRow("kakutei_kin") + newRow("tax_kin")
                newRow("seikyu_dt") = cpRow("seikyu_dt")
                newRow("seikyu_irai_no") = DBNull.Value
                newRow("nyukin_yotei_dt") = cpRow("nyukin_yotei_dt")
                newRow("seikyu_naiyo_nm") = "利用料"
                newRow("nyukin_dt") = DBNull.Value
                newRow("seq") = DBNull.Value
                newRow("shokei") = intRiyoTtl
                newRow("nyukin_kbn") = DBNull.Value
                newRow("nyukin_kin") = DBNull.Value
                newRow("nyukin_link_no") = DBNull.Value
                newRow("aite_cd") = cpRow("aite_cd")
                newRow("aite_nm") = cpRow("aite_nm")
                newRow("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO
                newRow("seikyu_title1") = cpRow("seikyu_title1")
                newRow("seikyu_title2") = cpRow("seikyu_title2")
                newRow("get_seikyu_input_flg") = "0"                                       ' 2016.02.18 ADD h.hagiwara 不具合対応
                table.Rows.Add(newRow)
                'プロジェクト設定
                Dim htKey As String
                If IsDBNull(cpRow("seikyu_irai_no")) Then
                    htKey = index.ToString
                Else
                    htKey = cpRow("seikyu_irai_no")
                End If
                Dim prjTable As DataTable = dataEXTB0102.PropHtExasPro(htKey).Copy()
                Dim prjIndex As New Integer
                Dim prjRow As DataRow
                Do While prjIndex < prjTable.Rows.Count
                    prjRow = prjTable.Rows(prjIndex)
                    prjRow("keijo_kin") = 0
                    prjRow("tax_kin") = 0
                    prjRow("seq") = DBNull.Value
                    prjRow("seikyu_irai_no") = DBNull.Value
                    prjRow("seikyu_irai_seq") = DBNull.Value
                    prjRow("content_cd") = DBNull.Value
                    prjRow("content_uchi_cd") = DBNull.Value
                    prjRow("event_nm") = ""
                    prjRow("content_uchi_nm") = ""
                    prjIndex += 1
                Loop
                dataEXTB0102.PropHtExasPro.Add((table.Rows.Count - 1).ToString, prjTable)
            End If
        End If
        index = 0
        isSet = False
        If intFutaiTtl <> 0 Then
            Do While index < cpTable.Rows.Count
                cpRow = cpTable.Rows(index)
                If cpRow("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Then
                    isSet = True
                    Exit Do
                End If
                index += 1
            Loop
            If isSet = False Then
                '該当行なしのためスキップ
            Else
                cpRow = cpTable.Rows(index)
                newRow = table.NewRow()
                newRow("seq") = DBNull.Value
                newRow("seikyu_input_flg") = "0"
                newRow("nyukin_input_flg") = "0"
                newRow("seikyu_irai_flg") = "0"
                newRow("kakutei_kin") = intFutaiTtl
                newRow("chosei_kin") = "0"
                newRow("tax_kin") = Integer.Parse(Math.Round((intFutaiTtl * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                newRow("seikyu_kin") = newRow("kakutei_kin") + newRow("tax_kin")
                newRow("seikyu_dt") = cpRow("seikyu_dt")
                newRow("seikyu_irai_no") = DBNull.Value
                newRow("nyukin_yotei_dt") = cpRow("nyukin_yotei_dt")
                newRow("seikyu_naiyo_nm") = "付帯設備"
                newRow("nyukin_dt") = DBNull.Value
                newRow("seq") = DBNull.Value
                newRow("shokei") = intFutaiTtl
                newRow("nyukin_kbn") = DBNull.Value
                newRow("nyukin_kin") = DBNull.Value
                newRow("nyukin_link_no") = DBNull.Value
                newRow("aite_cd") = cpRow("aite_cd")
                newRow("aite_nm") = cpRow("aite_nm")
                newRow("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI
                newRow("seikyu_title1") = cpRow("seikyu_title1")
                newRow("seikyu_title2") = cpRow("seikyu_title2")
                newRow("get_seikyu_input_flg") = "0"                                       ' 2016.02.18 ADD h.hagiwara 不具合対応
                table.Rows.Add(newRow)
                'プロジェクト設定
                Dim htKey As String
                If IsDBNull(cpRow("seikyu_irai_no")) Then
                    htKey = index.ToString
                Else
                    htKey = cpRow("seikyu_irai_no")
                End If
                Dim prjTable As DataTable = dataEXTB0102.PropHtExasProFutai(htKey).Copy()
                Dim prjIndex As New Integer
                Dim prjRow As DataRow
                Do While prjIndex < prjTable.Rows.Count
                    prjRow = prjTable.Rows(prjIndex)
                    prjRow("keijo_kin") = 0
                    prjRow("tax_kin") = 0
                    prjRow("seq") = DBNull.Value
                    prjRow("seikyu_irai_no") = DBNull.Value
                    prjRow("seikyu_irai_seq") = DBNull.Value
                    prjRow("content_cd") = DBNull.Value
                    prjRow("content_uchi_cd") = DBNull.Value
                    prjRow("event_nm") = ""
                    prjRow("content_uchi_nm") = ""
                    prjIndex += 1
                Loop
                dataEXTB0102.PropHtExasProFutai.Add((table.Rows.Count - 1).ToString, prjTable)
            End If
        End If
        ' 2015.12.16 MOVE START↓ h.hagiwara
        ''不要行削除(正規行は対象としない)
        'index = 0
        'Dim delCnt As New Integer
        'Dim htR As Hashtable = dataEXTB0102.PropHtExasPro
        'Dim htF As Hashtable = dataEXTB0102.PropHtExasProFutai
        'Do While index < table.Rows.Count
        '    row = table.Rows(index)
        '    Dim htKey As String = index.ToString
        '    If row("seikyu_input_flg") = "1" Then
        '        '対象外
        '        index += 1
        '    ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI And row("seikyu_kin") = 0 And index > 0 Then
        '        table.Rows.Remove(row)
        '        htR.Remove(htKey)
        '        htEditIndex(htR, index)
        '        htF.Remove(htKey)
        '        htEditIndex(htF, index)
        '    ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO And row("seikyu_kin") = 0 And index > 0 Then
        '        table.Rows.Remove(row)
        '        htEditIndex(htR, index)
        '    ElseIf row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI And row("seikyu_kin") = 0 And index > 1 Then
        '        table.Rows.Remove(row)
        '        htEditIndex(htF, index)
        '    Else
        '        index += 1
        '    End If
        '    If index = table.Rows.Count Then
        '        Exit Do
        '    End If
        'Loop
        ' 2015.12.16 MOVE END↑ h.hagiwara
        ' END Sub

        Return True                                       ' 2015.12.21 UPD h.hagiwara

    End Function                                          ' 2015.12.21 UPD h.hagiwara  

    Private Sub htEditIndex(ByRef ht As Hashtable, ByVal index As Integer)
        Dim htKey As String
        Dim innerIndex As Integer
        Dim prjTable As DataTable
        For Each key As String In ht.Keys
            If innerIndex >= index Then
                htKey = innerIndex.ToString
                If key = htKey Then
                    prjTable = ht(key).Copy()
                    ht.Remove(key)
                    htKey = (innerIndex - 1).ToString
                    ht.Add(htKey, prjTable)
                End If
            End If
            innerIndex += 1
        Next
    End Sub

    ''' <summary>
    ''' SPREAD請求一覧のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBillPay_Click(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbBillPay.ButtonClicked
        If e.Column = FB_BILLPAY_INPUT Then
            Dim frm As New EXTZ0203
            '「請求入力」画面を表示
            '押下行の取得
            ' 2016.03.03 ADD START↓ h.hagiwara
            Dim dtSeikyu As DateTime = System.DateTime.Now
            Dim dtNyukin1 As DateTime = DateTime.Parse(dtSeikyu.ToString("yyyy/MM/") & "01")
            Dim dtNyukin As DateTime
            Dim dtNyukinwork As DateTime
            ' 2016.03.03 ADD END↑ h.hagiwara
            Dim row As DataRow = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL").Rows(e.Row)
            frm.dataEXTZ0203.propStrYoyakuNo = dataEXTB0102.PropStrYoyakuNo
            frm.dataEXTZ0203.propStrSeq = commonLogicEXT.DbNullToNothing(row, "seq")
            frm.dataEXTZ0203.propStrSeikyuIraiNo = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_no")
            ' 2016.03.03 UPD START↓ h.hagiwara
            'frm.dataEXTZ0203.propStrSeikyuDt = commonLogicEXT.DbNullToNothing(row, "seikyu_dt")
            'frm.dataEXTZ0203.propStrNyukinYoteiDt = commonLogicEXT.DbNullToNothing(row, "nyukin_yotei_dt")
            If IsDBNull(row("seikyu_dt")) = True Then
                dtNyukinwork = dtNyukin1.AddMonths(2)
                dtNyukinwork = dtNyukinwork.AddDays(-1)
                dtNyukin = logicEXTB0103.WeekDayCheck(dtNyukinwork)
                frm.dataEXTZ0203.propStrSeikyuDt = dtSeikyu.ToString(CommonDeclareEXT.FMT_DATE)
                frm.dataEXTZ0203.propStrNyukinYoteiDt = dtNyukin.ToString(CommonDeclareEXT.FMT_DATE)
            Else
                frm.dataEXTZ0203.propStrSeikyuDt = commonLogicEXT.DbNullToNothing(row, "seikyu_dt")
                frm.dataEXTZ0203.propStrNyukinYoteiDt = commonLogicEXT.DbNullToNothing(row, "nyukin_yotei_dt")
            End If
            ' 2016.03.03 UPD END↑ h.hagiwara
            frm.dataEXTZ0203.propStrKakuteiKin = commonLogicEXT.DbNullToNothing(row, "kakutei_kin")
            frm.dataEXTZ0203.propStrChoseiKin = commonLogicEXT.DbNullToNothing(row, "chosei_kin")
            frm.dataEXTZ0203.propStrShokei = commonLogicEXT.DbNullToNothing(row, "shokei")
            frm.dataEXTZ0203.propStrTaxKin = commonLogicEXT.DbNullToNothing(row, "tax_kin")
            frm.dataEXTZ0203.propStrSeikyuKin = commonLogicEXT.DbNullToNothing(row, "seikyu_kin")
            frm.dataEXTZ0203.propStrSeikyuNaiyo = commonLogicEXT.DbNullToNothing(row, "seikyu_naiyo")
            frm.dataEXTZ0203.propStrSeikyuNaiyoNm = commonLogicEXT.DbNullToNothing(row, "seikyu_naiyo_nm")
            frm.dataEXTZ0203.propStrAiteCd = commonLogicEXT.DbNullToNothing(row, "aite_cd")
            frm.dataEXTZ0203.propStrAiteNm = commonLogicEXT.DbNullToNothing(row, "aite_nm")
            frm.dataEXTZ0203.propStrNyukinKbn = commonLogicEXT.DbNullToNothing(row, "nyukin_kbn")
            frm.dataEXTZ0203.propStrSeikyuTitle1 = commonLogicEXT.DbNullToNothing(row, "seikyu_title1")
            ' 2015.12.07 UPD START↓ h.hagiwara タイトル変更時の対応
            'frm.dataEXTZ0203.propStrSeikyuTitle2 = commonLogicEXT.DbNullToNothing(row, "seikyu_title2")
            If String.IsNullOrEmpty(Me.txtSaiji.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "催事名"), MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            If commonLogicEXT.DbNullToNothing(row, "seikyu_input_flg") Is Nothing Then
                If Me.txtSaiji.Text.Length > 25 Then
                    frm.dataEXTZ0203.propStrSeikyuTitle2 = Me.txtSaiji.Text.Substring(0, 25)
                Else
                    frm.dataEXTZ0203.propStrSeikyuTitle2 = Me.txtSaiji.Text
                End If
            ElseIf commonLogicEXT.DbNullToNothing(row, "seikyu_input_flg") = "0" Then
                If Me.txtSaiji.Text.Length > 25 Then
                    frm.dataEXTZ0203.propStrSeikyuTitle2 = Me.txtSaiji.Text.Substring(0, 25)
                Else
                    frm.dataEXTZ0203.propStrSeikyuTitle2 = Me.txtSaiji.Text
                End If
            Else
                frm.dataEXTZ0203.propStrSeikyuTitle2 = commonLogicEXT.DbNullToNothing(row, "seikyu_title2")
            End If
            ' 2015.12.07 UPD END↑ h.hagiwara タイトル変更時の対応
            frm.dataEXTZ0203.propStrNyukinDt = commonLogicEXT.DbNullToNothing(row, "nyukin_dt")
            frm.dataEXTZ0203.propStrNyukinKin = commonLogicEXT.DbNullToNothing(row, "nyukin_kin")
            frm.dataEXTZ0203.propStrSeikyuInputFlg = commonLogicEXT.DbNullToNothing(row, "seikyu_input_flg")
            frm.dataEXTZ0203.propStrSeikyuIraiFlg = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_flg")
            frm.dataEXTZ0203.propStrNyukinInputFlg = commonLogicEXT.DbNullToNothing(row, "nyukin_input_flg")
            frm.dataEXTZ0203.propStrNyukinLinkNo = commonLogicEXT.DbNullToNothing(row, "nyukin_link_no")
            frm.dataEXTZ0203.propStrGetSeikyuInputFlg = commonLogicEXT.DbNullToNothing(row, "get_seikyu_input_flg")                 ' 2016.02.03 ADD h.hagiwara
            frm.dataEXTZ0203.propDblTax = dataEXTB0102.propDblTax
            frm.dataEXTZ0203.PropDsBillReq = dataEXTB0102.PropDsBillReq
            frm.dataEXTZ0203.propIntEditLine = e.Row

            '相手先
            If inputCheckSeikyuInfoInput() = False Then
                MsgBox(CommonDeclareEXT.E2041, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            '利用料
            If SEIKYU_NAIYOU_RIYO = frm.dataEXTZ0203.propStrSeikyuNaiyo Or SEIKYU_NAIYOU_RIYOFUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Then
                Dim htExasRiyoryo As Hashtable = dataEXTB0102.PropHtExasPro
                Dim dtExasRiyoryo As DataTable
                If frm.dataEXTZ0203.propStrSeikyuIraiNo Is Nothing Then
                    dtExasRiyoryo = htExasRiyoryo(e.Row.ToString)
                ElseIf htExasRiyoryo.ContainsKey(frm.dataEXTZ0203.propStrSeikyuIraiNo) = True Then
                    dtExasRiyoryo = htExasRiyoryo(frm.dataEXTZ0203.propStrSeikyuIraiNo)
                Else
                    dtExasRiyoryo = htExasRiyoryo(e.Row.ToString)
                End If
                frm.dataEXTZ0203.PropDtExasRiyoryo = dtExasRiyoryo
            End If
            '付帯
            If SEIKYU_NAIYOU_FUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Or SEIKYU_NAIYOU_RIYOFUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Then
                Dim htExasFutai As Hashtable = dataEXTB0102.PropHtExasProFutai
                Dim dtExasFutai As DataTable
                If frm.dataEXTZ0203.propStrSeikyuIraiNo Is Nothing Then
                    dtExasFutai = htExasFutai(e.Row.ToString)
                ElseIf htExasFutai.ContainsKey(frm.dataEXTZ0203.propStrSeikyuIraiNo) = True Then
                    dtExasFutai = htExasFutai(frm.dataEXTZ0203.propStrSeikyuIraiNo)
                Else
                    dtExasFutai = htExasFutai(e.Row.ToString)
                End If
                frm.dataEXTZ0203.PropDtExasFutai = dtExasFutai
            End If
            frm.ShowDialog()
            SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
            If SEIKYU_NAIYOU_RIYO = frm.dataEXTZ0203.propStrSeikyuNaiyo Or SEIKYU_NAIYOU_RIYOFUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Then
                Dim htExasRiyoryo As Hashtable = dataEXTB0102.PropHtExasPro
                If frm.dataEXTZ0203.propStrSeikyuIraiNo Is Nothing Then
                    htExasRiyoryo(e.Row.ToString) = frm.dataEXTZ0203.PropDtExasRiyoryo
                ElseIf htExasRiyoryo.ContainsKey(frm.dataEXTZ0203.propStrSeikyuIraiNo) = True Then
                    htExasRiyoryo(frm.dataEXTZ0203.propStrSeikyuIraiNo) = frm.dataEXTZ0203.PropDtExasRiyoryo
                Else
                    htExasRiyoryo(e.Row.ToString) = frm.dataEXTZ0203.PropDtExasRiyoryo
                End If
            End If
            '付帯
            If SEIKYU_NAIYOU_FUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Or SEIKYU_NAIYOU_RIYOFUTAI = frm.dataEXTZ0203.propStrSeikyuNaiyo Then
                Dim htExasFutai As Hashtable = dataEXTB0102.PropHtExasProFutai
                If frm.dataEXTZ0203.propStrSeikyuIraiNo Is Nothing Then
                    htExasFutai(e.Row.ToString) = frm.dataEXTZ0203.PropDtExasFutai
                ElseIf htExasFutai.ContainsKey(frm.dataEXTZ0203.propStrSeikyuIraiNo) = True Then
                    htExasFutai(frm.dataEXTZ0203.propStrSeikyuIraiNo) = frm.dataEXTZ0203.PropDtExasFutai
                Else
                    htExasFutai(e.Row.ToString) = frm.dataEXTZ0203.PropDtExasFutai
                End If
            End If
        ElseIf e.Column = FB_NYUKIN_INPUT Then
            Try
                Dim frm As New EXTZ0207
                '「入金入力」画面を表示
                '押下行の取得
                Dim row As DataRow = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL").Rows(e.Row)
                frm.dataEXTZ0207.propStrYoyakuNo = dataEXTB0102.PropStrYoyakuNo
                frm.dataEXTZ0207.propStrSeq = commonLogicEXT.DbNullToNothing(row, "seq")
                frm.dataEXTZ0207.propStrSeikyuIraiNo = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_no")
                frm.dataEXTZ0207.propStrSeikyuDt = commonLogicEXT.DbNullToNothing(row, "seikyu_dt")
                frm.dataEXTZ0207.propStrNyukinYoteiDt = commonLogicEXT.DbNullToNothing(row, "nyukin_yotei_dt")
                frm.dataEXTZ0207.propStrKakuteiKin = commonLogicEXT.DbNullToNothing(row, "kakutei_kin")
                frm.dataEXTZ0207.propStrChoseiKin = commonLogicEXT.DbNullToNothing(row, "chosei_kin")
                frm.dataEXTZ0207.propStrShokei = commonLogicEXT.DbNullToNothing(row, "shokei")
                frm.dataEXTZ0207.propStrTaxKin = commonLogicEXT.DbNullToNothing(row, "tax_kin")
                frm.dataEXTZ0207.propStrSeikyuKin = commonLogicEXT.DbNullToNothing(row, "seikyu_kin")
                frm.dataEXTZ0207.propStrSeikyuNaiyo = commonLogicEXT.DbNullToNothing(row, "seikyu_naiyo")
                frm.dataEXTZ0207.propStrSeikyuNaiyoNm = commonLogicEXT.DbNullToNothing(row, "seikyu_naiyo_nm")
                frm.dataEXTZ0207.propStrAiteCd = commonLogicEXT.DbNullToNothing(row, "aite_cd")
                frm.dataEXTZ0207.propStrAiteNm = commonLogicEXT.DbNullToNothing(row, "aite_nm")
                frm.dataEXTZ0207.propStrNyukinKbn = commonLogicEXT.DbNullToNothing(row, "nyukin_kbn")
                frm.dataEXTZ0207.propStrSeikyuTitle1 = commonLogicEXT.DbNullToNothing(row, "seikyu_title1")
                frm.dataEXTZ0207.propStrSeikyuTitle2 = commonLogicEXT.DbNullToNothing(row, "seikyu_title2")
                frm.dataEXTZ0207.propStrNyukinDt = commonLogicEXT.DbNullToNothing(row, "nyukin_dt")
                frm.dataEXTZ0207.propStrNyukinKin = commonLogicEXT.DbNullToNothing(row, "nyukin_kin")
                frm.dataEXTZ0207.propStrSeikyuInputFlg = commonLogicEXT.DbNullToNothing(row, "seikyu_input_flg")
                frm.dataEXTZ0207.propStrSeikyuIraiFlg = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_flg")
                frm.dataEXTZ0207.propStrNyukinInputFlg = commonLogicEXT.DbNullToNothing(row, "nyukin_input_flg")
                frm.dataEXTZ0207.propStrNyukinLinkNo = commonLogicEXT.DbNullToNothing(row, "nyukin_link_no")
                frm.dataEXTZ0207.propStrGetSeikyuInputFlg = commonLogicEXT.DbNullToNothing(row, "get_seikyu_input_flg")                ' 2016.02.03 ADD h.hagiwara
                frm.dataEXTZ0207.propDblTax = dataEXTB0102.propDblTax
                frm.dataEXTZ0207.PropDsBillReq = dataEXTB0102.PropDsBillReq
                frm.dataEXTZ0207.propIntEditLine = e.Row
                frm.dataEXTZ0207.PropRowNyukin = dataEXTB0102.PropDtExasNyukin.Rows(e.Row)
                frm.ShowDialog()
                If frm.dataEXTZ0207.PropBlnChangeFlg = True Then
                    '入金情報変更フラグ設定
                    Dim editRow As DataRow = frm.dataEXTZ0207.PropRowNyukin
                    row("nyukin_dt") = frm.dataEXTZ0207.propStrNyukinDt
                    row("nyukin_kin") = frm.dataEXTZ0207.propStrNyukinKin
                    row("nyukin_input_flg") = frm.dataEXTZ0207.propStrNyukinInputFlg
                    row("nyukin_link_no") = commonLogicEXT.DbNothingToNull(frm.dataEXTZ0207.propStrNyukinLinkNo)
                    SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 依頼追加のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnIrai_Click(sender As Object, e As EventArgs) Handles btnIrai.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0213
        frm.PropStrIraiOrKiroku = SHONIN_IRAI
        frm.PropDsApproval = dataEXTB0102.PropDsApproval
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        SetSpreadIraiRireki(dataEXTB0102.PropDsApproval.Tables("IRAI_RIREKI_TBL"))
        If frm.PropBlnChangeFlg = True Then
            strIraiFlg = "1"
        End If
        logicEXTB0103.RegRirekiKiroku(dataEXTB0102)
    End Sub

    ''' <summary>
    ''' 記録追加のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShonin_Click(sender As Object, e As EventArgs) Handles btnShonin.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0213
        frm.PropStrIraiOrKiroku = SHONIN_KIROKU
        frm.PropDsApproval = dataEXTB0102.PropDsApproval
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        SetSpreadIraiKakunin(dataEXTB0102.PropDsApproval.Tables("CHECK_RIREKI_TBL"))
        If frm.PropBlnChangeFlg = True Then
            strShoninFlg = "1"
        End If
        logicEXTB0103.RegRirekiKiroku(dataEXTB0102)
    End Sub

    ''' <summary>
    ''' 行追加のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnTimeSchAdd_Click(sender As Object, e As EventArgs) Handles btnTimeSchAdd.Click
        Dim index As New Integer
        Dim table As DataTable
        Dim row As DataRow
        table = dataEXTB0102.PropDsTimeSch.Tables("TIME_SCHEDULE_TBL")
        row = table.NewRow
        row("seq") = DBNull.Value
        row("time") = DBNull.Value
        row("biko") = DBNull.Value
        table.Rows.Add(row)
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTimeSchedule.ActiveSheet
        sheet.RowCount = sheet.RowCount + 1
    End Sub

    ''' <summary>
    ''' 行追加のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnTicketAdd_Click(sender As Object, e As EventArgs) Handles btnTicketAdd.Click
        Dim index As New Integer
        Dim table As DataTable
        Dim numCellRyokin As New FarPoint.Win.Spread.CellType.NumberCellType()                  ' 2015.12.08 ADD h.hagiwara
        'numCellRyokin.MaximumValue = 9999999999                                                ' 2015.12.08 ADD h.hagiwara
        numCellRyokin.MaximumValue = 99999999                                                   ' 2015.12.18 UPD h.hagiwara
        numCellRyokin.ShowSeparator = True                                                      ' 2015.12.08 ADD h.hagiwara
        numCellRyokin.DecimalPlaces = 0                                                         ' 2015.12.08 ADD h.hagiwara

        Dim row As DataRow
        table = dataEXTB0102.PropDsTicket.Tables("TICKET_TBL")
        row = table.NewRow
        row("seq") = DBNull.Value
        row("ticket_nm") = DBNull.Value
        row("ticket_kin") = DBNull.Value
        table.Rows.Add(row)
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTicket.ActiveSheet
        sheet.RowCount = sheet.RowCount + 1
        sheet.Cells(sheet.RowCount - 1, 1).CellType = numCellRyokin                             ' 2015.12.08 ADD h.hagiwara

    End Sub

    ''' <summary>
    ''' 利用形状定員設定処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdoRiyoType_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRiyoType1.CheckedChanged, rdoRiyoType2.CheckedChanged,
                                                                                        rdoRiyoType3.CheckedChanged, rdoRiyoType4.CheckedChanged
        If Me.rdoRiyoType1.Checked = True Then
            Me.txtTeiin.Text = CommonDeclareEXT.PropIntTeiinA.ToString
        ElseIf Me.rdoRiyoType2.Checked = True Then
            Me.txtTeiin.Text = CommonDeclareEXT.PropIntTeiinB.ToString
        End If
    End Sub

    ''' <summary>
    ''' 請求情報変更フラグ設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub editFlgSetting()
        '点滅
        Me.timFlashLabel.Start()
        strFlashFlg = "1"
        'Disable処理
        Me.btnYoyakuAdd.Enabled = False
        Me.btnYoyakuDel.Enabled = False
        Me.btnRiyoshaSearch.Enabled = False
        Me.btnSplit.Enabled = False
        Me.btnBond.Enabled = False
        Me.btnClaimCancel.Enabled = False
        Me.btnPayCancel.Enabled = False
        Me.btnPrintRiyoShonin.Enabled = False
        Me.btnPrintFutaiTotal.Enabled = False
        Me.btnPrintExas.Enabled = False
        Me.btnRegister.Enabled = False
        Me.btnCancel.Enabled = False
        'Me.btnBack.Enabled = False                          ' 2015.12.02 DEL h.hagiwara
        Me.btnIrai.Enabled = False
        Me.btnShonin.Enabled = False
        Me.btnTimeSchAdd.Enabled = False
        Me.btnTicketAdd.Enabled = False
    End Sub

    ''' <summary>
    ''' 請求分割
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSplit_Click(sender As Object, e As EventArgs) Handles btnSplit.Click
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim index As New Integer
        Dim blnSelected As Boolean = False
        Dim table1 As DataTable
        Dim table2 As DataTable
        Dim row1 As DataRow
        Dim row2 As DataRow
        Dim strBillypayflg As String = ""                                                ' 2015.12.11 ADD h.hagiwara
        Dim intCnt As Integer = 0                                                        ' 2015.12.11 ADD h.hagiwara

        index = 0
        Dim frm As New EXTZ0206
        ' 2015.12.11 UPD START↓ h.hagiwara 請求情報入力済の判定追加
        'Do While index < sheet.Rows.Count
        '    If sheet.Cells(index, 0).Value = True Then
        '        frm.dataEXTZ0206.PropStrFromNaiyo = sheet.Cells(index, 12).Value
        '        frm.dataEXTZ0206.PropIntFromKakutei = sheet.Cells(index, 5).Value
        '        frm.dataEXTZ0206.PropStrFromName = sheet.Cells(index, 24).Value
        '        blnSelected = True
        '    End If
        '    index = index + 1
        'Loop
        For i = 0 To sheet.Rows.Count - 1
            If sheet.Cells(i, 0).Value = True Then
                frm.dataEXTZ0206.PropStrFromNaiyo = sheet.Cells(i, 12).Value
                frm.dataEXTZ0206.PropIntFromKakutei = sheet.Cells(i, 5).Value
                frm.dataEXTZ0206.PropStrFromName = sheet.Cells(i, 24).Value
                blnSelected = True
                strBillypayflg = sheet.Cells(i, 1).Value
                intCnt += 1
                index = i
            End If
        Next
        ' 2015.12.11 UPD END↑ h.hagiwara 請求情報入力済の判定追加

        '2015.12.11 ADD START↓ h.hagiwara 請求情報入力済の判定追加
        If intCnt > 1 Then
            MsgBox(String.Format(CommonEXT.E2045, "分割対象の請求"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        ElseIf intCnt = 1 Then
            If strBillypayflg <> "" Then
                MsgBox(String.Format(CommonEXT.E2046, "分割処理"), MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
        End If
        '2015.12.11 ADD END↑ h.hagiwara 請求情報入力済の判定追加

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If blnSelected = True Then
            frm.ShowDialog()
            If frm.dataEXTZ0206.PropBlnChangeFlg Then
                table1 = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
                table2 = table1.Copy
                row1 = table1.Rows(index)
                row2 = table2.Rows(index)
                '分割元の行への設定
                row1("kakutei_kin") = frm.dataEXTZ0206.PropIntToKakutei1
                row1("tax_kin") = Integer.Parse(Math.Round((row1("kakutei_kin") * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                row1("shokei") = row1("kakutei_kin") + row1("chosei_kin")
                row1("seikyu_kin") = row1("shokei") + row1("tax_kin")
                '分割先の行を分割元行を複製して作成する
                row2("kakutei_kin") = frm.dataEXTZ0206.PropIntToKakutei2
                row2("chosei_kin") = 0
                row2("tax_kin") = Integer.Parse(Math.Round((row2("kakutei_kin") * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
                row2("shokei") = row2("kakutei_kin") + row2("chosei_kin")
                row2("seikyu_kin") = row2("shokei") + row2("tax_kin")
                row2("seq") = DBNull.Value
                row2("seikyu_irai_no") = DBNull.Value
                row2("nyukin_dt") = DBNull.Value
                row2("nyukin_kin") = DBNull.Value
                row2("nyukin_link_no") = DBNull.Value
                '各種情報の複製
                Dim htRiyoryo As Hashtable
                ' 2015.12.16 UPD START↓ h.hagiwara 分割時のEXASプロジェクト情報上の利用計上額､消費税額を再設定
                'If row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Or row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                '    htRiyoryo = dataEXTB0102.PropHtExasPro
                '    Dim dtRiyoryo As DataTable
                '    If IsDBNull(row1("seikyu_irai_no")) = False Then
                '        dtRiyoryo = htRiyoryo(row1("seikyu_irai_no").ToString)
                '    Else
                '        dtRiyoryo = htRiyoryo(index.ToString)
                '    End If
                '    htRiyoryo.Add((table1.Rows.Count).ToString, prjCopy(dtRiyoryo))
                '    dataEXTB0102.PropHtExasPro = htRiyoryo.Clone
                'End If
                If row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Then
                    htRiyoryo = dataEXTB0102.PropHtExasPro
                    Dim dtRiyoryo As DataTable
                    If IsDBNull(row1("seikyu_irai_no")) = False Then
                        dtRiyoryo = htRiyoryo(row1("seikyu_irai_no").ToString)
                    Else
                        dtRiyoryo = htRiyoryo(index.ToString)
                    End If
                    prjCopy2(dtRiyoryo, frm.dataEXTZ0206.PropIntToKakutei1, "1")
                    htRiyoryo.Add((table1.Rows.Count).ToString, prjCopy2(dtRiyoryo, frm.dataEXTZ0206.PropIntToKakutei2, "2"))
                    dataEXTB0102.PropHtExasPro = htRiyoryo.Clone
                End If
                If row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    htRiyoryo = dataEXTB0102.PropHtExasPro
                    Dim dtRiyoryo As DataTable
                    If IsDBNull(row1("seikyu_irai_no")) = False Then
                        dtRiyoryo = htRiyoryo(row1("seikyu_irai_no").ToString)
                    Else
                        dtRiyoryo = htRiyoryo(index.ToString)
                    End If
                    htRiyoryo.Add((table1.Rows.Count).ToString, prjCopy(dtRiyoryo))
                    dataEXTB0102.PropHtExasPro = htRiyoryo.Clone
                End If
                ' 2015.12.16 UPD END↑ h.hagiwara 分割時のEXASプロジェクト情報上の利用計上額､消費税額を再設定
                Dim htFutai As Hashtable
                If row1("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Or row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    htFutai = dataEXTB0102.PropHtExasProFutai
                    Dim dtFutai As DataTable
                    If IsDBNull(row1("seikyu_irai_no")) = False Then
                        dtFutai = htFutai(row1("seikyu_irai_no").ToString)
                    Else
                        dtFutai = htFutai(index.ToString)
                    End If
                    htFutai.Add((table1.Rows.Count).ToString, prjCopy(dtFutai))
                    dataEXTB0102.PropHtExasProFutai = htFutai
                End If
                Dim dtNyukin As DataTable = dataEXTB0102.PropDtExasNyukin
                Dim dtNyukinTemp As DataTable = dtNyukin.Copy
                Dim nyuRow As DataRow = dtNyukinTemp.Rows(index)
                If nyuRow("line_no") = index Then
                    nyuRow("line_no") = dtNyukin.Rows.Count + 1
                    ' 2015.12.16 ADD START↓ h.hagiwara 
                    nyuRow("nyukin_yotei_dt") = DBNull.Value
                    nyuRow("nyukin_dt") = DBNull.Value
                    nyuRow("seikyu_kin") = DBNull.Value
                    nyuRow("seikyu_dt") = DBNull.Value
                    nyuRow("input_dt") = DBNull.Value
                    nyuRow("sekikyu_no") = DBNull.Value
                    nyuRow("seikyu_irai_no") = DBNull.Value
                    nyuRow("seikyu_kin") = DBNull.Value
                    nyuRow("seikyu_kin1") = DBNull.Value
                    nyuRow("nyukin_link_no") = DBNull.Value
                    ' 2015.12.16 ADD END↑ h.hagiwara 
                    dtNyukin.ImportRow(nyuRow)
                End If
                table1.ImportRow(row2)
                SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
            End If
        Else
            MsgBox(String.Format(CommonEXT.E0002, "分割対象の請求"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' PROJECTコピー
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function prjCopy(ByVal table As DataTable) As DataTable
        Dim newTable As DataTable = table.Copy()
        Dim index As New Integer
        Dim row As DataRow
        Do While index < table.Rows.Count
            'row = table.Rows(index)                                                ' 2015.12.16 UPD h.hagiwara
            row = newTable.Rows(index)                                              ' 2015.12.16 UPD h.hagiwara
            row("seq") = DBNull.Value
            row("seikyu_irai_no") = DBNull.Value
            row("seikyu_irai_seq") = DBNull.Value
            index = index + 1
        Loop
        Return newTable
    End Function

    ''' <summary>
    ''' PROJECTコピー(利用料のみ分割した場合の処理）    ' 2015.12.16 ADD h.hagiwara
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function prjCopy2(ByVal table As DataTable, ByVal kakutei As Object, ByRef str1 As String) As DataTable
        Dim newTable As DataTable = table.Copy()
        Dim index As New Integer
        Dim row As DataRow
        Do While index < table.Rows.Count
            If str1 = "1" Then
                row = table.Rows(index)
            Else
                row = newTable.Rows(index)
                row("seq") = DBNull.Value
                row("seikyu_irai_no") = DBNull.Value
                row("seikyu_irai_seq") = DBNull.Value
            End If
            row("keijo_kin") = 0
            row("tax_kin") = 0
            If str1 = "1" And index = 0 Then
                row("keijo_kin") = kakutei
                row("tax_kin") = Integer.Parse(Math.Round((kakutei * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
            ElseIf str1 = "2" And index = table.Rows.Count - 1 Then
                row("keijo_kin") = kakutei
                row("tax_kin") = Integer.Parse(Math.Round((kakutei * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))
            End If
            index = index + 1
        Loop
        Return newTable
    End Function

    ''' <summary>
    ''' 請求結合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBond_Click(sender As Object, e As EventArgs) Handles btnBond.Click
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim index As New Integer
        Dim index1 As New Integer '1つめのチェック行
        Dim index2 As New Integer '2つめのチェック行
        Dim blnSelected1 As Boolean = False
        Dim blnSelected2 As Boolean = False
        Dim table As DataTable
        Dim row1 As DataRow
        Dim row2 As DataRow
        Dim key1 As String
        Dim key2 As String
        Dim intBillypaycnt As Integer                                                 ' 2015.12.11 ADD h.hagiwara
        Dim dtTemp1 As DataTable                                                      ' 2015.12.16 ADD h.hagiwara
        Dim exasrow1 As DataRow                                                       ' 2015.12.16 ADD h.hagiwara
        Dim exasrow2 As DataRow                                                       ' 2015.12.16 ADD h.hagiwara

        intBillypaycnt = 0                                                            ' 2015.12.11 ADD h.hagiwara
        index = 0
        Dim frm As New EXTZ0206
        Do While index < sheet.Rows.Count
            If sheet.Cells(index, 0).Value = True Then
                If blnSelected1 = False Then
                    index1 = index
                    blnSelected1 = True
                    If sheet.Cells(index, 1).Value <> "" Then                             ' 2015.12.11 ADD h.hagiwara
                        intBillypaycnt += 1                                               ' 2015.12.11 ADD h.hagiwara
                    End If                                                                ' 2015.12.11 ADD h.hagiwara
                Else
                    index2 = index
                    blnSelected2 = True
                    If sheet.Cells(index, 1).Value <> "" Then                             ' 2015.12.11 ADD h.hagiwara
                        intBillypaycnt += 1                                               ' 2015.12.11 ADD h.hagiwara
                    End If                                                                ' 2015.12.11 ADD h.hagiwara
                    Exit Do
                End If
            End If
            index = index + 1
        Loop

        '2015.12.11 ADD START↓ h.hagiwara 請求情報入力済の判定追加
        If intBillypaycnt > 0 Then
            MsgBox(String.Format(CommonEXT.E2046, "結合処理"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '2015.12.11 ADD END↑ h.hagiwara 請求情報入力済の判定追加

        If blnSelected2 = True Then
            table = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
            row1 = table.Rows(index1)
            row2 = table.Rows(index2)
            row1("kakutei_kin") = row1("kakutei_kin") + row2("kakutei_kin")
            row1("chosei_kin") = row1("chosei_kin") + row2("chosei_kin")
            row1("shokei") = row1("kakutei_kin") + row1("chosei_kin")
            ' row1("tax_kin") = Integer.Parse(Math.Round((row1("shokei") * dataEXTB0102.propDblTax), MidpointRounding.AwayFromZero))    ' 2015/11/13 UPD h.hagiwara
            row1("tax_kin") = row1("tax_kin") + row2("tax_kin")                                                                         ' 2015/11/13 UPD h.hagiwara
            row1("seikyu_kin") = row1("shokei") + row1("tax_kin")
            If IsDBNull(row1("seikyu_irai_no")) = False Then
                key1 = row1("seikyu_irai_no").ToString
            Else
                key1 = index1.ToString
            End If
            If IsDBNull(row2("seikyu_irai_no")) = False Then
                key2 = row2("seikyu_irai_no").ToString
            Else
                key2 = index2.ToString
            End If
            If (row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO And row2("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO) Then
                Dim dtTemp As DataTable = dataEXTB0102.PropHtExasPro(key2)
                dataEXTB0102.PropHtExasPro.Remove(key2)
                ' 2015.12.16 ADD START↓ h.hagiwara EXAS請求情報の内容も統合
                dtTemp1 = dataEXTB0102.PropHtExasPro(key1)
                For i = 0 To dtTemp1.Rows.Count - 1
                    exasrow1 = dtTemp1.Rows(i)
                    For j = 0 To dtTemp.Rows.Count - 1
                        exasrow2 = dtTemp.Rows(j)
                        If exasrow1("riyo_ym") = exasrow2("riyo_ym") And _
                           exasrow1("kamoku_cd") = exasrow2("kamoku_cd") And _
                           exasrow1("saimoku_cd") = exasrow2("saimoku_cd") And _
                           exasrow1("uchi_cd") = exasrow2("uchi_cd") And _
                           exasrow1("shosai_cd") = exasrow2("shosai_cd") Then
                            exasrow1("keijo_kin") = exasrow1("keijo_kin") + exasrow2("keijo_kin")
                            exasrow1("tax_kin") = exasrow1("tax_kin") + exasrow2("tax_kin")
                        End If
                    Next
                Next
                dataEXTB0102.PropHtExasPro.Remove(key1)
                dataEXTB0102.PropHtExasPro.Add(key1, dtTemp1)
                ' 2015.12.16 ADD END↑ h.hagiwara EXAS請求情報の内容も統合
            ElseIf (row1("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI And row2("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI) Then
                Dim dtTemp As DataTable = dataEXTB0102.PropHtExasProFutai(key2)
                dataEXTB0102.PropHtExasProFutai.Remove(key2)
            ElseIf (row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO And row2("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI) Then
                '請求内容の編集
                row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI
                row1("seikyu_naiyo_nm") = "利用料+付帯設備"
                logicEXTB0103.SetSeikyuTitle(dataEXTB0102, row1("seikyu_title1"), SEIKYU_NAIYOU_FUTAI)
                Dim dtTemp As DataTable = dataEXTB0102.PropHtExasProFutai(key2)
                dataEXTB0102.PropHtExasProFutai.Remove(key1)
                dataEXTB0102.PropHtExasProFutai.Remove(key2)
                dataEXTB0102.PropHtExasProFutai.Add(key1, dtTemp)
            ElseIf (row2("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO And row1("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI) Then
                '請求内容の編集
                row1("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI
                row1("seikyu_naiyo_nm") = "利用料+付帯設備"
                logicEXTB0103.SetSeikyuTitle(dataEXTB0102, row1("seikyu_title1"), SEIKYU_NAIYOU_FUTAI)
                Dim dtTemp As DataTable = dataEXTB0102.PropHtExasProFutai(key1)
                dataEXTB0102.PropHtExasProFutai.Remove(key1)
                dataEXTB0102.PropHtExasProFutai.Remove(key2)
                dataEXTB0102.PropHtExasProFutai.Add(key1, dtTemp)
                ' 2016.02.18 ADD START↓ h.hagiwara 利用料のハッシュキーを変更
                Dim dtTempPro As DataTable = dataEXTB0102.PropHtExasPro(key2)
                dataEXTB0102.PropHtExasPro.Remove(key2)
                dataEXTB0102.PropHtExasPro.Add(key1, dtTempPro)
                ' 2016.02.18 ADD END↑ h.hagiwara 利用料のハッシュキーを変更
            End If
            table.Rows.Remove(row2)
            SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        Else
            MsgBox(String.Format(CommonEXT.E0002, "結合対象の請求２行"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' 【画面表示用】ALSOK入金情報を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function SetSpreadAlsokMeisai(ByVal table As DataTable) As Boolean
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbCash.ActiveSheet
        sheet.RowCount = table.Rows.Count
        sheet.ColumnCount = 5
        Dim index As New Integer
        Dim lngTtl As New Long
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("deposit_dt")
            sheet.Cells(index, 1).Value = row("regi_nm")
            sheet.Cells(index, 2).Value = row("tenpo_nm")
            sheet.Cells(index, 3).Value = row("deposit_amount")
            index = index + 1
            lngTtl += row("deposit_amount")
        Loop
        Me.lblCashTotal.Text = String.Format("{0:#,##0}", lngTtl)

        'スクロールバーを必要な場合のみ表示させます
        fbBillPay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbBillPay.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        If table.Rows.Count = 0 Then
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 【画面表示用】ALSOK入金情報を画面表示するSPREADに設定する
    ''' </summary>
    ''' <param name="table"></param>
    ''' <remarks></remarks>
    Private Function SetSpreadAlsokStore(ByVal table As DataTable) As Boolean
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbStore.ActiveSheet
        sheet.RowCount = table.Rows.Count
        sheet.ColumnCount = 3
        Dim index As New Integer
        Dim row As DataRow
        index = 0

        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Value = row("regi_nm")
            sheet.Cells(index, 1).Value = row("tenpo_nm")
            sheet.Cells(index, 2).Value = row("deposit_amount")
            index = index + 1
        Loop

        'スクロールバーを必要な場合のみ表示させます
        fbBillPay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbBillPay.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        If table.Rows.Count = 0 Then
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 更新ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        '画面入力情報をDataに格納
        convertData(dataEXTB0102)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '利用日データ
        SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
        SetTimeScheduleSpread(dataEXTB0102.PropDsTimeSch.Tables("TIME_SCHEDULE_TBL"))
        SetTicketSpread(dataEXTB0102.PropDsTicket.Tables("TICKET_TBL"))
        '入力チェック
        If inputCheckMain() = False Then
            Exit Sub
        End If
        If inputCheckCondition() = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0002, "予約情報"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        'ステータス設定
        If dataEXTB0102.PropStrYoyakuSts > YOYAKU_STS_SEISHIKI_COMP Then                                             ' 2015.12.10 ADD h.hagiwara 取消情報はステータス変更しない
        Else                                                                                                         ' 2015.12.10 ADD h.hagiwara 取消情報はステータス変更しない
            If Me.rdoStatus1.Checked = True Then
                dataEXTB0102.PropStrYoyakuSts = CommonDeclareEXT.YOYAKU_STS_SEISHIKI
            ElseIf Me.rdoStatus2.Checked = True Then
                dataEXTB0102.PropStrYoyakuSts = CommonDeclareEXT.YOYAKU_STS_SEISHIKI_COMP
            End If
        End If                                                                                                       ' 2015.12.10 ADD h.hagiwara 取消情報はステータス変更しない
        '予約制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表削除
        If logicEXTB0102.DeleteYoyakuList(dataEXTB0102.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTB0102.InsertYoyakuList(dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約情報更新
        If logicEXTB0103.RegYoyakuDetail(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTB0103_Load(sender, e)
    End Sub

    ''' <summary>
    ''' 画面入力情報をDataに格納
    ''' </summary>
    ''' <remarks>画面入力情報をDataに格納
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertData(ByRef dataEXTB0102 As DataEXTB0102)
        'ヘッダ
        '受付USERCDが空の場合設定する
        If String.IsNullOrEmpty(dataEXTB0102.PropStrKakuUsercd) = True Then
            dataEXTB0102.PropStrKakuUsercd = CommonDeclareEXT.PropComStrUserId
        End If
        dataEXTB0102.PropStrKakuteiDt = Me.dtpUke.txtDate.Text

        '催事
        dataEXTB0102.PropStrSaijiNm = Me.txtSaiji.Text
        dataEXTB0102.PropStrShutsuenNm = Me.txtShutuen.Text

        If Me.rdoKashi1.Checked = True Then
            dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
        ElseIf Me.rdoKashi2.Checked = True Then
            dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
        ElseIf Me.rdoKashi3.Checked = True Then
            dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI
        End If
        '利用形状
        If Me.rdoRiyoType1.Checked = True Then
            dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_STAND
        ElseIf Me.rdoRiyoType2.Checked = True Then
            dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SEATING
        ElseIf Me.rdoRiyoType3.Checked = True Then
            dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MIX
        ElseIf Me.rdoRiyoType4.Checked = True Then
            dataEXTB0102.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SAIJI
        End If
        dataEXTB0102.PropStrTeiin = Me.txtTeiin.Text
        'ワンドリンク
        If Me.rdoOneDrink1.Checked = True Then
            dataEXTB0102.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_ARI
        ElseIf Me.rdoOneDrink2.Checked = True Then
            dataEXTB0102.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_NASHI
        End If
        '催事分類
        If Me.rdoSaijiBunrui1.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MUSIC
        ElseIf Me.rdoSaijiBunrui2.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEKI
        ElseIf Me.rdoSaijiBunrui3.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEI
        ElseIf Me.rdoSaijiBunrui4.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_BUSINESS
        ElseIf Me.rdoSaijiBunrui5.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MOVIE
        ElseIf Me.rdoSaijiBunrui6.Checked = True Then
            dataEXTB0102.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ETC
        End If
        '利用者情報
        dataEXTB0102.PropStrRiyoshaCd = Me.lblRiyoshaCd.Text
        dataEXTB0102.PropStrRiyoKana = Me.txtRiyoshaNmKana.Text
        dataEXTB0102.PropStrRiyoNm = Me.txtRiyoshaNm.Text
        dataEXTB0102.PropStrDaihyoNm = Me.txtDaihyoNm.Text
        dataEXTB0102.PropStrSekininBushoNm = Me.txtSekininBushoNm.Text
        dataEXTB0102.PropStrSekininNm = Me.cmbSekininNm.Text
        dataEXTB0102.PropStrSekininMail = Me.txtSekininMail.Text
        dataEXTB0102.PropStrAiteNm = Me.lblExasAiteNm.Text
        dataEXTB0102.PropStrAiteCd = Me.lblExasAite.Text
        dataEXTB0102.PropStrRiyoYubin1 = Me.txtRiyoPost1.Text
        dataEXTB0102.PropStrRiyoYubin2 = Me.txtRiyoPost2.Text
        dataEXTB0102.PropStrRiyoTodo = Me.cmbRiyoTodo.Text
        dataEXTB0102.PropStrRiyoShiku = Me.txtRiyoShiku.Text
        dataEXTB0102.PropStrRiyoBan = Me.txtRiyoBan.Text
        dataEXTB0102.PropStrRiyoBuild = Me.txtRiyoBuild.Text
        dataEXTB0102.PropStrRiyoTel11 = Me.txtRiyoTel1.Text
        dataEXTB0102.PropStrRiyoTel12 = Me.txtRiyoTel2.Text
        dataEXTB0102.PropStrRiyoTel13 = Me.txtRiyoTel3.Text
        dataEXTB0102.PropStrRiyoNaisen = Me.txtRiyoNaisen.Text
        dataEXTB0102.PropStrRiyoTel21 = Me.txtRiyoMobileTel1.Text
        'dataEXTB0102.PropStrRiyoTel22 = Me.txtRiyoMobileTel2.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTB0102.PropStrRiyoTel23 = Me.txtRiyoMobileTel3.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTB0102.PropStrRiyoFax11 = Me.txtRiyoFax1.Text
        dataEXTB0102.PropStrRiyoFax12 = Me.txtRiyoFax2.Text
        dataEXTB0102.PropStrRiyoFax13 = Me.txtRiyoFax3.Text

        '特記事項
        dataEXTB0102.PropStrBiko = Me.txtBiko.Text
        dataEXTB0102.PropStrRiyoCom = Me.txtComment.Text

        'チケット
        If Me.rdoTicket1.Checked Then
            dataEXTB0102.PropStrTicketEnterKbn = CommonDeclareEXT.TICKET_ENTER_YURYO
        ElseIf Me.rdoTicket2.Checked Then
            dataEXTB0102.PropStrTicketEnterKbn = CommonDeclareEXT.TICKET_ENTER_MURYO
        End If
        If Me.rdoTicketDrink1.Checked Then
            dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_KOMI
        ElseIf Me.rdoTicketDrink2.Checked Then
            dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_BETSU
        ElseIf Me.rdoTicketDrink3.Checked Then
            dataEXTB0102.PropStrTicketDrinkKbn = CommonDeclareEXT.TICKET_DRINK_NASHI
        End If
        'HP
        If Me.rdoEtr1.Checked Then
            dataEXTB0102.PropStrHpKeisai = CommonDeclareEXT.HP_OK
        ElseIf Me.rdoEtr2.Checked Then
            dataEXTB0102.PropStrHpKeisai = CommonDeclareEXT.HP_NG
        End If
        dataEXTB0102.PropStrJohoKokaiDt = Me.dtpInfoDisc.txtDate.Text
        dataEXTB0102.PropStrJohoKokaiTime = Me.txtDiscTime1.Text
        dataEXTB0102.PropStrKokaiDt = Me.dtpDisc.txtDate.Text
        dataEXTB0102.PropStrKokaiTime = Me.txtDiscTime2.Text
        '付帯ステータス
        If Me.rdoFutaiSts1.Checked Then
            dataEXTB0102.PropStrFinputSts = F_INPUT_MIKAKU
        ElseIf Me.rdoFutaiSts2.Checked Then
            dataEXTB0102.PropStrFinputSts = F_INPUT_KAKUTEI
        End If
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
            If String.IsNullOrEmpty(Me.dtpUke.txtDate.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "受付日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfChar(Me.dtpUke.txtDate.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0004, "受付日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '催事名
            If String.IsNullOrEmpty(Me.txtSaiji.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "催事名"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '出演者名
            ' 2016.01.05 DEL START↓ h.hagiwara
            'If String.IsNullOrEmpty(Me.txtShutuen.Text) Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0001, "出演者名"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            ' 2016.01.05 DEL END↑ h.hagiwara
            '定員
            If commonValidate.IsHalfNmb(Me.txtTeiin.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "定員"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
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
            For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
                ' 2015.12.08 DEL START↓ h.hagiwara 
                'If String.IsNullOrEmpty(sheet.Cells(index, 3).Value) Then
                '    MsgBox(String.Format(CommonDeclareEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                '    Return False
                'End If
                ' 2015.12.08 DEL END↑ h.hagiwara 
                If commonValidate.IsHalfNmb(sheet.Cells(index, 3).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                ' 2015.12.08 DEL START↓ h.hagiwara 
                'If String.IsNullOrEmpty(sheet.Cells(index, 4).Value) Then
                '    MsgBox(String.Format(CommonDeclareEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                '    Return False
                'End If
                ' 2015.12.08 DEL END↑ h.hagiwara 
                If commonValidate.IsHalfNmb(sheet.Cells(index, 4).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If String.IsNullOrEmpty(sheet.Cells(index, 5).Value) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "単価"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfNmb(sheet.Cells(index, 5).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "単価"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If String.IsNullOrEmpty(sheet.Cells(index, 7).Value) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "倍率"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfChar(sheet.Cells(index, 7).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0004, "倍率"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If String.IsNullOrEmpty(sheet.Cells(index, 8).Value) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "数量"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfNmb(sheet.Cells(index, 8).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "数量"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If String.IsNullOrEmpty(sheet.Cells(index, 9).Value) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "料金"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfNmb(sheet.Cells(index, 9).Value) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E0003, "料金"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                index = index + 1
            Next
            index = 0
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
            ' 2016.11.2 m.hayabuchi MOD Start 課題No.59
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
            ' 2016.11.2 m.hayabuchi MOD End 課題No.59
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

            If String.IsNullOrEmpty(Me.txtDiscTime1.Text) = False Then
                If Len(Me.txtDiscTime1.Text) <> 4 Then
                    MsgBox(String.Format(CommonDeclareEXT.E0010, "情報公開　公開時間", 4), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If
            If String.IsNullOrEmpty(Me.txtDiscTime2.Text) = False Then
                If Len(Me.txtDiscTime2.Text) <> 4 Then
                    MsgBox(String.Format(CommonDeclareEXT.E0010, "公開日　公開時間", 4), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
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
    Public Function inputCheckCondition() As Boolean

        'トレースログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_START", Nothing, Nothing)

        Try
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
            ' 2016.11.2 m.hayabuchi DEL Start 課題No.59
            'If String.IsNullOrEmpty(Me.txtRiyoMobileTel1.Text) = False Then _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel2.Text) = False _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel3.Text) = False Then
            'If String.IsNullOrEmpty(Me.txtRiyoMobileTel1.Text) = True Then _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel2.Text) = True _
            '    Or String.IsNullOrEmpty(Me.txtRiyoMobileTel3.Text) = True Then
            'MsgBox(String.Format(CommonDeclareEXT.E2020, "携帯番号"), MsgBoxStyle.Exclamation, "エラー")
            'Return False
            'End If
            'End If
            ' 2016.11.2 m.hayabuchi DEL End 課題No.59
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
            '利用日(SPREAD)
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
            Dim index As Integer = 0
            Dim lineCnt As Integer = 1
            For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
                ' 2015.12.08 ADD START↓ h.hagiwara 
                If String.IsNullOrEmpty(sheet.Cells(index, 3).Value) Then
                    If String.IsNullOrEmpty(sheet.Cells(index, 4).Value) Then
                    Else
                        MsgBox(String.Format(CommonDeclareEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                Else
                    If String.IsNullOrEmpty(sheet.Cells(index, 4).Value) Then
                        MsgBox(String.Format(CommonDeclareEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                End If
                ' 2015.12.08 ADD END↑ h.hagiwara 
                '予約数超過
                If logicEXTZ0202.CheckRiyobiRegister(TOUROKU_KBN_KARI, dataRiyobi) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                    Return False
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
    ''' ボタン点滅用タイマーイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub timFlashLabel_Tick(sender As Object, e As EventArgs) Handles timFlashLabel.Tick
        timFlashLabel.Enabled = True
        timFlashLabel.Interval = 500
        If btnBillInfoUpdate.BackColor = Color.Orange Then
            btnBillInfoUpdate.BackColor = Color.Red
        ElseIf btnBillInfoUpdate.BackColor = Color.Red Then
            btnBillInfoUpdate.BackColor = Color.Orange
        End If
    End Sub

    ''' <summary>
    ''' 請求情報更新ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBillInfoUpdate_Click(sender As Object, e As EventArgs) Handles btnBillInfoUpdate.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '更新設定(曜日追加)
        SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
        '付帯
        If logicEXTB0103.UpFutai(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            '設定
            SetSpreadFutai(dataEXTB0102.PropHtFutai)
        End If
        '請求再計算
        ' 2015.12.21 UPD START↓ h.hagiwara
        'CalcBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        If CalcBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara
        SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        'プロジェクト情報
        If logicEXTB0103.CalcExasRiyoryo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        'EXAS入金情報
        If logicEXTB0103.GetNyukin(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '点滅解除
        If strFlashFlg = "1" Then
            Me.timFlashLabel.Stop()
            strFlashFlg = "0"
        End If
        'Enable処理
        Me.btnYoyakuAdd.Enabled = True
        Me.btnYoyakuDel.Enabled = True
        Me.btnRiyoshaSearch.Enabled = True
        Me.btnSplit.Enabled = True
        Me.btnBond.Enabled = True
        Me.btnRegister.Enabled = True
        Me.btnCancel.Enabled = True
        Me.btnBack.Enabled = True
        Me.btnTimeSchAdd.Enabled = True
        Me.btnTicketAdd.Enabled = True
        statusBtnSetting()
    End Sub

    ''' <summary>
    ''' 仮予約ボタン活性状態設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub statusBtnSetting()
        If dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_KARI Then
            'Disable処理
            Me.btnPrintRiyoShonin.Enabled = False
            Me.btnPrintFutaiTotal.Enabled = False
            Me.btnPrintExas.Enabled = False
            Me.btnClaimCancel.Enabled = False
            Me.btnPayCancel.Enabled = False
            Me.btnIrai.Enabled = False
            Me.btnShonin.Enabled = False
        Else
            Me.btnPrintRiyoShonin.Enabled = True
            Me.btnPrintFutaiTotal.Enabled = True
            Me.btnPrintExas.Enabled = True
            Me.btnClaimCancel.Enabled = True
            Me.btnPayCancel.Enabled = True
            Me.btnIrai.Enabled = True
            Me.btnShonin.Enabled = True
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

        ' 2015.12.02 DEL START↓ h.hagiwara
        'If strPrintRiyoShoninFlg = "1" Or strPrintFutaiTotalFlg = "1" Or strPrintExasFlg = "1" Then
        '    MsgBox(CommonEXT.E2025, MsgBoxStyle.Exclamation, "エラー")
        '    Exit Sub
        'End If
        ' 2015.12.02 DEL END↑ h.hagiwara
        '予約受付制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        End If
        Me.Close()
    End Sub

    ''' <summary>
    ''' 閉じるButton処理
    ''' 未更新時無効
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTB0103_FormClosing(ByVal sender As Object, _
            ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2015.12.02 DEL START↓ h.hagiwara
        'If strPrintRiyoShoninFlg = "1" Or strPrintFutaiTotalFlg = "1" Or strPrintExasFlg = "1" Then
        '    If e.CloseReason = CloseReason.UserClosing Then
        '        MsgBox(CommonEXT.E2025, MsgBoxStyle.Exclamation, "エラー")
        '        e.Cancel = True
        '    End If
        'End If
        ' 2015.12.02 DEL END↑ h.hagiwara
        '予約受付制御削除
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        End If
    End Sub

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
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '更新
            dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_CANCEL_SEISHIKI
            'If logicEXTB0102.RegYoyakuInfo(dataEXTB0102, True) = False Then                                       ' 2015.12.10 UPD h.hagiwara
            If logicEXTB0103.RegYoyakuInfo(dataEXTB0102) = False Then                                              ' 2015.12.10 UPD h.hagiwara
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.PropStrDeleteKbn = TORIKESHI_KBN_DELETE Then
            '物理削除
            '予約制御削除
            If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTB0102.DeleteYoyakuList(dataEXTB0102.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTB0102.DeleteYoyaku(dataEXTB0102.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.DialogResult = System.Windows.Forms.DialogResult.OK Then
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' 請求取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnClaimCancel_Click(sender As Object, e As EventArgs) Handles btnClaimCancel.Click
        '確認処理
        If MsgBox(CommonEXT.C0007, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim dtBillPay As DataTable = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
        Dim index As Integer = 0
        Dim row As DataRow
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim isIraiSumi As Boolean = False
        Do While index < sheet.RowCount
            If sheet.Cells(index, 0).Value = True Then
                row = dtBillPay.Rows(index)
                If row("nyukin_input_flg") = "1" Then
                    MsgBox(CommonEXT.E2040, MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                If row("seikyu_irai_flg") = "1" Then
                    isIraiSumi = True
                End If
            End If
            index = index + 1
        Loop
        If isIraiSumi = True Then
            If MsgBox(CommonEXT.I0005, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
                Return
            End If
        End If
        index = 0
        Do While index < sheet.RowCount
            If sheet.Cells(index, 0).Value = True Then
                row = dtBillPay.Rows(index)
                row("seikyu_input_flg") = "0"
                row("seikyu_irai_flg") = "0"
                row("get_seikyu_input_flg") = "0"
                row("seikyu_dt") = Nothing                                        ' 2016.03.03 ADD h.hagiwara
                row("nyukin_yotei_dt") = Nothing                                  ' 2016.03.03 ADD h.hagiwara
                If logicEXTB0103.CancellSeikyu(row) = False Then
                    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
            End If
            index = index + 1
        Loop

        MsgBox(String.Format(CommonDeclareEXT.I0002, "請求取消"), MsgBoxStyle.Information, "完了")

        SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        ' 2015.12.22 ADD START↓ h.hagiwara
        'プロジェクト情報
        If logicEXTB0103.GetExasRiyoryo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.22 ADD END↑ h.hagiwara

    End Sub

    ''' <summary>
    ''' 入金取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPayCancel_Click(sender As Object, e As EventArgs) Handles btnPayCancel.Click
        '確認処理
        If MsgBox(CommonEXT.C0006, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim dtBillPay As DataTable = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
        Dim index As Integer = 0
        Dim row As DataRow
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Do While index < sheet.RowCount
            If sheet.Cells(index, 0).Value = True Then
                row = dtBillPay.Rows(index)
                row("nyukin_input_flg") = "0"
                row("nyukin_link_no") = DBNull.Value                  ' 2016.03.03 ADD h.hagiwara
                row("nyukin_dt") = Nothing                            ' 2016.03.03 ADD h.hagiwara
                row("nyukin_kin") = DBNull.Value                      ' 2016.09.09 ADD m.hayabuchi 不具合対応
                If logicEXTB0103.CancellNyukin(row) = False Then
                    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If

                '2016.09.21 m.hayabuchi add START 不具合対応
                If logicEXTB0103.GetNyukin(dataEXTB0102) = False Then
                    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                Dim frm As New EXTZ0207
                Dim dtRowNyukin As DataRow = dataEXTB0102.PropDtExasNyukin.Rows(index)

                frm.SetNyukinSpread(dtRowNyukin)
                '2016.09.21 m.hayabuchi add END 不具合対応

            End If
            index = index + 1
        Loop
        SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))

        MsgBox(String.Format(CommonDeclareEXT.I0002, "入金取消"), MsgBoxStyle.Information, "完了")
    End Sub

    ''' <summary>
    '''  [利用承認書印刷]ボタンクリック時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>利用承認書の印刷を行う
    ''' <para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnPrintRiyoShonin_Click(sender As Object, e As EventArgs) Handles btnPrintRiyoShonin.Click
        strPrintRiyoShoninFlg = "1"

        '印刷
        If MsgBox(EXT_C0008, MsgBoxStyle.OkCancel, "利用承認書印刷") = vbOK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            ' 消費税取得基準日(利用日）
            '予約番号
            dataEXTB0103.PropStrReserveNo_Output = Me.lblYoyakuNo.Text
            dataEXTB0103.PropStrCalculateDay_Output = Me.fpRiyobi.ActiveSheet.Cells(0, 2).Value
            dataEXTB0103.PropStrCalculateDay_Output = dataEXTB0103.PropStrCalculateDay_Output.Substring(0, 10)
            If logicEXTB0103.PrintCertificatesMain(dataEXTB0103) = False Then
                MsgBox(puErrMsg)
            End If
        End If

    End Sub

    ''' <summary>
    '''  [付帯利用明細(合算)印刷]ボタンクリック時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>利用承認書の印刷を行う
    ''' <para>作成情報：2015/09/01 h.mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnPrintFutaiTotal_Click(sender As Object, e As EventArgs) Handles btnPrintFutaiTotal.Click
        strPrintFutaiTotalFlg = "1"

        '予約番号
        dataEXTB0103.PropStrReserveNo_Output = Me.lblYoyakuNo.Text
        '日付
        dataEXTB0103.PropStrCalculateDay_Output = Me.fpRiyobi.ActiveSheet.Cells(0, 2).Value
        dataEXTB0103.PropStrCalculateDay_Output = dataEXTB0103.PropStrCalculateDay_Output.Substring(0, 10)
        '貸出区分
        dataEXTB0103.PropStrRentalClass_Output = dataEXTB0102.PropStrKashiKind

        If MsgBox(EXT_C0009, MsgBoxStyle.OkCancel, TITLE_INFO) = vbOK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            If logicEXTB0103.PrintUseDetailsMain(dataEXTB0103) = False Then
                MsgBox(puErrMsg)
            End If
        End If

    End Sub

    ''' <summary>
    ''' 出力：EXAS
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrintExas_Click(sender As Object, e As EventArgs) Handles btnPrintExas.Click
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim index As New Integer
        Dim isChecked As Boolean = False
        Dim strContHead As String = "PA0"
        Dim isSetSeikyu As Boolean = True        ' 2015.12.02 ADD h.hagiwara 請求情報未入力＆請求依頼番号未発番はEXAS請求データを出力不可とする
        Dim checkCount As New Integer            ' 2016.01.15 ADD y.morooka グループ請求対応　チェックボックス複数チェック判定
        Dim frm As New EXTZ0214                  ' 2016.01.20 ADD y.morooka グループ請求対応
        Dim isSetAitesaki As Boolean = True      ' 2016.01.25 ADD y.morooka グループ請求対応 相手先チェック

        Do While sheet.RowCount > index
            If sheet.Cells(index, 0).Value = True Then
                isChecked = True
                checkCount = checkCount + 1         '2016.01.15 ADD y.morooka グループ請求対応　チェックボックス複数チェック判定
                ' 2015.12.02 ADD START↓ h.hagiwara 請求情報未入力＆請求依頼番号未発番はEXAS請求データを出力不可とする
                If IsDBNull(sheet.Cells(index, 1).Value) Or IsDBNull(sheet.Cells(index, 10).Value) Then
                    isSetSeikyu = False
                ElseIf sheet.Cells(index, 1).Value = "" Or sheet.Cells(index, 10).Value = "" Then
                    isSetSeikyu = False
                End If
                ' 2015.12.02 ADD END↑ h.hagiwara 請求情報未入力＆請求依頼番号未発番はEXAS請求データを出力不可とする
            End If
            index = index + 1
        Loop
        If isChecked = False Then
            MsgBox(String.Format(CommonEXT.E0002, "対象の請求"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2015.12.02 ADD START↓ h.hagiwara 請求情報未入力＆請求依頼番号未発番はEXAS請求データを出力不可とする
        If isSetSeikyu = False Then
            MsgBox(CommonEXT.E2043, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.02 ADD END↑ h.hagiwara 請求情報未入力＆請求依頼番号未発番はEXAS請求データを出力不可とする
        ' 2016.01.15 ADD START↓ y.morooka グループ請求対応　EXAS相手先がグループの場合
        If IsDBNull(dataEXTB0102.PropStrAiteCd) Then
            isSetAitesaki = False
        ElseIf dataEXTB0102.PropStrAiteCd = "" Then
            isSetAitesaki = False
        End If
        If isSetAitesaki = False Then
            MsgBox(CommonEXT.E2050, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If isSetAitesaki = True Then
            If PropGSeikyusaki Like "*" & dataEXTB0102.PropStrAiteCd & "*" Then
                'チェックボックス複数チェック判定
                If checkCount > 1 Then
                    Dim result As DialogResult = MsgBox(CommonEXT.C0015, MsgBoxStyle.YesNo, "Check")
                    If result = MsgBoxResult.No Then Exit Sub
                End If
                '子画面呼び出し
                'パラメータ設定
                frm.dataEXTZ0214.PropStrAitesakiCD = dataEXTB0102.PropStrAiteCd
                frm.dataEXTZ0214.PropStrAitesakiNm = dataEXTB0102.PropStrAiteNm

                '画面を表示
                frm.ShowDialog()

                'グループ請求情報画面の戻るボタン押下時、Exit Subする
                If frm.dataEXTZ0214.PropBlnGSeikyuCloseFlg = True Then
                    Exit Sub
                End If
            End If
        End If
        ' 2016.01.15 ADD END↑ y.morooka グループ請求対応　EXAS相手先がグループの場合
        index = 0

        Dim sfd As New SaveFileDialog()
        'sfd.FileName = "EXAS請求依頼_" + dataEXTB0102.PropStrYoyakuNo + ".csv"                                                     ' 2015.11.26 UPD h.hagiwara
        sfd.FileName = "EXAS請求依頼_" & dataEXTB0102.PropStrYoyakuNo & "_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".csv"      ' 2015.11.26 UPD h.hagiwara
        'ディレクトリ
        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        sfd.InitialDirectory = stCurrentDir
        sfd.Filter = "CSVファイル|*.csv|すべてのファイル(*.*)|*.*"
        sfd.FilterIndex = 1
        sfd.Title = "EXAS請求依頼の出力先を選択してください"
        sfd.RestoreDirectory = True

        If sfd.ShowDialog() = DialogResult.OK Then
            Dim sw As New StreamWriter(sfd.FileName, False, System.Text.Encoding.GetEncoding("Shift-JIS"))
            Dim sbHeader As StringBuilder
            Dim sbDetail As StringBuilder
            Dim table1 As DataTable = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
            ' 2016.08.12 ADD START↓ m.hayabuchi 代行処理対応
            Dim StrTantoNm = ""
            Dim StrTantoBusho = ""

            '代行フラグを確認
            If CommonDeclareEXT.PropStrDaikoFlg.Equals("1") Then
                '代行処理の場合
                StrTantoNm = CommonDeclareEXT.PropStrDaikoTanto
                StrTantoBusho = CommonDeclareEXT.PropStrDaikoBusho
            End If
            ' 2016.08.12 ADD END↑ m.hayabuchi 代行処理対応

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            ' 見出し出力
            sw.WriteLine(CSV_HEADER)

            Do While sheet.RowCount > index
                '請求
                Dim row As DataRow
                Dim ht As Hashtable
                Dim htKey As String
                Dim table2 As DataTable
                Dim dIndex As New Integer

                'If sheet.Cells(index, 0).Value = True Then                                                             ' 2015.12.16 UPD h.hagiwara
                If sheet.Cells(index, 0).Value = True And sheet.Cells(index, 8).Value <> 0 Then                         ' 2015.12.16 UPD h.hagiwara
                    row = table1.Rows(index)
                    'Header
                    sbHeader = New StringBuilder
                    '請求日
                    appendHdr(sbHeader, row("seikyu_dt").ToString.Replace("/", ""))
                    '入金予定日
                    appendHdr(sbHeader, row("nyukin_yotei_dt").ToString.Replace("/", ""))
                    '当社担当者コード
                    'appendHdr(sbHeader, "")    ' 2016.08.12 m.hayabuchi mod 代行処理対応
                    appendHdr(sbHeader, StrTantoNm)
                    '当社担当者所属部署コード
                    'appendHdr(sbHeader, "")    ' 2016.08.12 m.hayabuchi mod 代行処理対応
                    appendHdr(sbHeader, StrTantoBusho)
                    '当社部署
                    appendHdr(sbHeader, "")
                    '担当者Tｅｌ
                    appendHdr(sbHeader, "")
                    '請求先コード
                    appendHdr(sbHeader, row("aite_cd"))
                    dataEXTB0103.PropStrAitecd = row("aite_cd")
                    dataEXTB0103.PropStrPostno = ""
                    dataEXTB0103.PropStrAddr1 = ""
                    dataEXTB0103.PropStrAddr2 = ""
                    dataEXTB0103.PropStrAitenm = ""
                    If logicEXTB0103.GetAitesakiInf(dataEXTB0103) = False Then
                        '請求書郵便番号
                        appendHdr(sbHeader, Me.txtRiyoPost1.Text + Me.txtRiyoPost2.Text)
                        '請求書住所１
                        appendHdr(sbHeader, Me.cmbRiyoTodo.Text + Me.txtRiyoShiku.Text + Me.txtRiyoBan.Text)
                        '請求書住所２
                        appendHdr(sbHeader, Me.txtRiyoBuild.Text)
                        '請求先名称
                        appendHdr(sbHeader, row("aite_nm"))
                    Else
                        '請求書郵便番号
                        appendHdr(sbHeader, dataEXTB0103.PropStrPostno)
                        '請求書住所１
                        appendHdr(sbHeader, dataEXTB0103.PropStrAddr1)
                        '請求書住所２
                        appendHdr(sbHeader, dataEXTB0103.PropStrAddr2)
                        '請求先名称
                        appendHdr(sbHeader, dataEXTB0103.PropStrAitenm)
                    End If
                    'G請求先部署コード
                    ' 2016.01.20 MOD START↓ y.morooka グループ請求対応　EXAS相手先がグループの場合
                    'appendHdr(sbHeader, "")
                    If frm.dataEXTZ0214.PropBlnGSeikyuFlg = False Then
                        appendHdr(sbHeader, "")
                    Else
                        appendHdr(sbHeader, frm.dataEXTZ0214.PropStrSeikyusakiBusyoCD)
                    End If
                    ' 2016.01.20 MOD END↑ y.morooka グループ請求対応　EXAS相手先がグループの場合
                    '請求書部署名
                    appendHdr(sbHeader, "")
                    'G請求書担当者コード
                    ' 2016.01.20 MOD START↓ y.morooka グループ請求対応　EXAS相手先がグループの場合
                    'appendHdr(sbHeader, "")
                    If frm.dataEXTZ0214.PropBlnGSeikyuFlg = False Then
                        appendHdr(sbHeader, "")
                    Else
                        appendHdr(sbHeader, frm.dataEXTZ0214.PropStrSeikyusyoTantoCD)
                    End If
                    ' 2016.01.20 MOD END↑ y.morooka グループ請求対応　EXAS相手先がグループの場合
                    '請求書担当者名
                    appendHdr(sbHeader, "")
                    '経理連絡欄
                    appendHdr(sbHeader, "")
                    'タイトル１
                    appendHdr(sbHeader, row("seikyu_title1"))
                    'タイトル２
                    appendHdr(sbHeader, row("seikyu_title2"))
                    '納品書備考
                    appendHdr(sbHeader, "")
                    '明細
                    '利用料
                    If row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                        ht = dataEXTB0102.PropHtExasPro
                        If IsDBNull(row("seikyu_irai_no")) Then
                            htKey = index.ToString
                        Else
                            htKey = row("seikyu_irai_no")
                        End If
                        table2 = ht(htKey)
                        Dim dRow As DataRow
                        dIndex = 0
                        Do While table2.Rows.Count > dIndex
                            sbDetail = New StringBuilder
                            dRow = table2.Rows(dIndex)
                            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "EXAS請求依頼 利用料[" + dRow("riyo_ym") + "]", Nothing, Nothing)

                            If dRow("keijo_kin") <> 0 Then                                             ' 2015.12.16 ADD h.hagiwara

                                '処理区分
                                appendDtl(sbDetail, "2")
                                '番組コード
                                appendDtl(sbDetail, "")
                                '番組シーケンス
                                appendDtl(sbDetail, "")
                                'プロジェクトコード
                                appendDtl(sbDetail, strContHead & dRow("content_cd"))
                                'プロジェクトシーケンス
                                appendDtl(sbDetail, dRow("content_uchi_cd"))
                                '目的コード
                                appendDtl(sbDetail, "")
                                '預り先コード
                                appendDtl(sbDetail, "")
                                '計上部署コード
                                appendDtl(sbDetail, "")
                                'コンテンツ識別区分
                                appendDtl(sbDetail, "")
                                'コンテンツコード
                                appendDtl(sbDetail, "")
                                'コンテンツ内訳コード
                                appendDtl(sbDetail, "")
                                '予算外案件コード
                                appendDtl(sbDetail, "")
                                '勘定科目コード
                                appendDtl(sbDetail, dRow("kamoku_cd"))
                                '細目コード
                                appendDtl(sbDetail, dRow("saimoku_cd"))
                                '内訳コード
                                appendDtl(sbDetail, dRow("uchi_cd"))
                                '詳細コード
                                appendDtl(sbDetail, dRow("shosai_cd"))
                                '借方勘定科目コード
                                appendDtl(sbDetail, dRow("karikamoku_cd"))
                                '借方細目コード
                                appendDtl(sbDetail, dRow("kari_saimoku_cd"))
                                '借方内訳コード
                                appendDtl(sbDetail, dRow("kari_uchi_cd"))
                                '借方詳細コード
                                appendDtl(sbDetail, dRow("kari_shosai_cd"))
                                '発生月自
                                appendDtl(sbDetail, dRow("riyo_ym"))
                                '発生月至
                                appendDtl(sbDetail, "")
                                'セグメントコード
                                appendDtl(sbDetail, "")
                                '入力摘要１
                                appendDtl(sbDetail, dRow("tekiyo1"))
                                '入力摘要２
                                appendDtl(sbDetail, dRow("tekiyo2") & "[" & row("seikyu_irai_no") & "]")
                                '単価
                                appendDtl(sbDetail, dRow("keijo_kin"))
                                '数量
                                appendDtl(sbDetail, "1")
                                '消費税額
                                appendDtl(sbDetail, dRow("tax_kin"))
                                '消費税区分
                                appendDtl(sbDetail, "")
                                '消費税率
                                appendDtl(sbDetail, Double.Parse(dataEXTB0102.propDblTax) * 100)
                                '外税内税区分
                                appendDtl(sbDetail, "")
                                'G請求内容コード
                                ' 2016.01.20 MOD START↓ y.morooka グループ請求対応　EXAS相手先がグループの場合
                                'appendDtl(sbDetail, "")
                                If frm.dataEXTZ0214.PropBlnGSeikyuFlg = False Then
                                    appendDtl(sbDetail, "")
                                Else
                                    appendDtl(sbDetail, frm.dataEXTZ0214.PropStrSeikyuNaiyoCD)
                                End If
                                ' 2016.01.20 MOD END↑ y.morooka グループ請求対応　EXAS相手先がグループの場合
                                'G_セグメントコード
                                appendDtl(sbDetail, "")
                                'Gコンテンツ識別区分
                                appendDtl(sbDetail, "")
                                'Gコンテンツコード
                                appendDtl(sbDetail, "")
                                'Gコンテンツ内訳コード
                                appendDtl(sbDetail, "")

                                sw.WriteLine(sbHeader.ToString + sbDetail.ToString)                    ' 2015.12.16 ADD h.hagiwara
                            End If                                                                     ' 2015.12.16 ADD h.hagiwara

                            dIndex = dIndex + 1
                            'sw.WriteLine(sbHeader.ToString + sbDetail.ToString)                       ' 2015.12.16 DEL h.hagiwara
                        Loop
                    End If
                    '付帯
                    If row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                        ht = dataEXTB0102.PropHtExasProFutai
                        If IsDBNull(row("seikyu_irai_no")) Then
                            htKey = index.ToString
                        Else
                            htKey = row("seikyu_irai_no")
                        End If
                        table2 = ht(htKey)
                        Dim dRow As DataRow
                        dIndex = 0
                        Do While table2.Rows.Count > dIndex
                            sbDetail = New StringBuilder
                            dRow = table2.Rows(dIndex)
                            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "EXAS請求依頼 付帯設備使用料[" + dRow("shukei_grp") + "]", Nothing, Nothing)

                            If dRow("keijo_kin") <> 0 Then                                             ' 2015.12.16 ADD h.hagiwara

                                '処理区分
                                appendDtl(sbDetail, "2")
                                '番組コード
                                appendDtl(sbDetail, "")
                                '番組シーケンス
                                appendDtl(sbDetail, "")
                                'プロジェクトコード
                                appendDtl(sbDetail, strContHead & dRow("content_cd"))
                                'プロジェクトシーケンス
                                appendDtl(sbDetail, dRow("content_uchi_cd"))
                                '目的コード
                                appendDtl(sbDetail, "")
                                '預り先コード
                                appendDtl(sbDetail, "")
                                '計上部署コード
                                appendDtl(sbDetail, "")
                                'コンテンツ識別区分
                                appendDtl(sbDetail, "")
                                'コンテンツコード
                                appendDtl(sbDetail, "")
                                'コンテンツ内訳コード
                                appendDtl(sbDetail, "")
                                '予算外案件コード
                                appendDtl(sbDetail, "")
                                '勘定科目コード
                                appendDtl(sbDetail, dRow("kamoku_cd"))
                                '細目コード
                                appendDtl(sbDetail, dRow("saimoku_cd"))
                                '内訳コード
                                appendDtl(sbDetail, dRow("uchi_cd"))
                                '詳細コード
                                appendDtl(sbDetail, dRow("shosai_cd"))
                                '借方勘定科目コード
                                appendDtl(sbDetail, dRow("karikamoku_cd"))
                                '借方細目コード
                                appendDtl(sbDetail, dRow("kari_saimoku_cd"))
                                '借方内訳コード
                                appendDtl(sbDetail, dRow("kari_uchi_cd"))
                                '借方詳細コード
                                appendDtl(sbDetail, dRow("kari_shosai_cd"))
                                '発生月自
                                appendDtl(sbDetail, dRow("riyo_ym"))
                                '発生月至
                                appendDtl(sbDetail, "")
                                'セグメントコード
                                appendDtl(sbDetail, "")
                                '入力摘要１
                                appendDtl(sbDetail, dRow("tekiyo1"))
                                '入力摘要２
                                appendDtl(sbDetail, dRow("tekiyo2") & "[" & row("seikyu_irai_no") & "]")
                                '単価
                                appendDtl(sbDetail, dRow("keijo_kin"))
                                '数量
                                appendDtl(sbDetail, "1")
                                ' 税抜区分の取得
                                dataEXTB0103.PropStrCalculateDay_Output = Me.fpRiyobi.ActiveSheet.Cells(0, 2).Value
                                dataEXTB0103.PropStrCalculateDay_Output = dataEXTB0103.PropStrCalculateDay_Output.Substring(0, 10)
                                dataEXTB0103.PropStrGrpKey = dRow("shukei_grp")
                                If logicEXTB0103.GetNotaxflg(dataEXTB0103) = "1" Then
                                    '消費税額
                                    appendDtl(sbDetail, "0")
                                    '消費税区分
                                    appendDtl(sbDetail, "10")
                                    '消費税率
                                    appendDtl(sbDetail, "")
                                    '外税内税区分
                                    appendDtl(sbDetail, "2")
                                Else
                                    '消費税額
                                    appendDtl(sbDetail, dRow("tax_kin"))
                                    '消費税区分
                                    appendDtl(sbDetail, "")
                                    '消費税率
                                    appendDtl(sbDetail, Double.Parse(dataEXTB0102.propDblTax) * 100)
                                    '外税内税区分
                                    appendDtl(sbDetail, "")
                                End If
                                'G請求内容コード
                                ' 2016.01.20 MOD START↓ y.morooka グループ請求対応　EXAS相手先がグループの場合
                                'appendDtl(sbDetail, "")
                                If frm.dataEXTZ0214.PropBlnGSeikyuFlg = False Then
                                    appendDtl(sbDetail, "")
                                Else
                                    appendDtl(sbDetail, frm.dataEXTZ0214.PropStrSeikyuNaiyoCD)
                                End If
                                ' 2016.01.20 MOD END↑ y.morooka グループ請求対応　EXAS相手先がグループの場合
                                'G_セグメントコード
                                appendDtl(sbDetail, "")
                                'Gコンテンツ識別区分
                                appendDtl(sbDetail, "")
                                'Gコンテンツコード
                                appendDtl(sbDetail, "")
                                'Gコンテンツ内訳コード
                                appendDtl(sbDetail, "")

                                sw.WriteLine(sbHeader.ToString + sbDetail.ToString)                        ' 2015.12.16 ADD h.hagiwara

                            End If                                                                         ' 2015.12.16 ADD h.hagiwara
                            dIndex = dIndex + 1
                            'sw.WriteLine(sbHeader.ToString + sbDetail.ToString)                           ' 2015.12.16 DEL h.hagiwara
                        Loop
                    End If
                    row("seikyu_irai_flg") = "1"
                End If
                index = index + 1
            Loop
            sw.Close()
            SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
            strPrintExasFlg = "1"
            MsgBox(String.Format(CommonDeclareEXT.I0002, "EXAS用請求依頼ファイルの出力"), MsgBoxStyle.Information, "完了")
        End If
    End Sub

    ''' <summary>
    ''' StringBuilder[Header]追記
    ''' </summary>
    ''' <param name="sb"></param>
    ''' <param name="addObj"></param>
    ''' <remarks></remarks>
    Private Sub appendHdr(ByRef sb As StringBuilder, ByVal addObj As Object)
        append(sb, addObj, False)
    End Sub

    ''' <summary>
    ''' StringBuilder[Detail]追記
    ''' </summary>
    ''' <param name="sb"></param>
    ''' <param name="addObj"></param>
    ''' <remarks></remarks>
    Private Sub appendDtl(ByRef sb As StringBuilder, ByVal addObj As Object)
        append(sb, addObj, True)
    End Sub

    ''' <summary>
    ''' StringBuilder追記処理
    ''' </summary>
    ''' <param name="sb"></param>
    ''' <param name="addObj"></param>
    ''' <param name="isDtl"></param>
    ''' <remarks></remarks>
    Private Sub append(ByRef sb As StringBuilder, ByVal addObj As Object, ByVal isDtl As Boolean)
        Dim sep As String = ","
        Dim quote As String = """"
        If IsDBNull(addObj) Then
            addObj = ""
        ElseIf String.IsNullOrEmpty(addObj) Then
            addObj = ""
        End If

        If sb.Length > 0 Or isDtl Then
            'sb.Append(sep + quote + addObj.ToString + quote)           ' 2015.11.13 UPD h.hagiwara
            sb.Append(sep + addObj.ToString)                            ' 2015.11.13 UPD h.hagiwara
        Else
            'sb.Append(quote + addObj.ToString + quote)                 ' 2015.11.13 UPD h.hagiwara
            sb.Append(addObj.ToString)                                  ' 2015.11.13 UPD h.hagiwara
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
        frm.DataEXTM0201.PropParamValue = "0"
        '画面を表示
        frm.ShowDialog()
        '利用者番号設定
        If String.IsNullOrEmpty(frm.DataEXTM0201.PropParamRiyoCd) = False Then
            With frm.DataEXTM0201
                Me.lblRiyoshaCd.Text = .PropParamRiyoCd
                dataEXTB0102.PropStrRiyoLvl = .PropParamRiyoLvl
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTB0102.PropStrRiyoLvl)
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

            ' 2016.11.04 ADD START m.hayabuchi         ' 責任者情報クリア（責任者名,メールアドレス,携帯電話番号）
            Me.cmbSekininNm.Text = ""
            Me.txtSekininMail.Text = ""
            Me.txtRiyoMobileTel1.Text = ""
            ' 2016.11.04 ADD END m.hayabuchi           ' 責任者情報クリア（責任者名,メールアドレス,携帯電話番号）

            '責任者再取得
            If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                Dim list As ArrayList = logicEXTB0102.GetSekininshaList(Me.lblRiyoshaCd.Text)
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

            '請求入金情報の相手先CDと相手先名を再設定
            For i As Integer = 0 To dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL").Rows.count - 1
                dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL").Rows(i)("aite_cd") = frm.DataEXTM0201.PropParamAiteCd
                dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL").Rows(i)("aite_nm") = frm.DataEXTM0201.PropParamAiteNm
            Next
        End If
    End Sub

    ''' <summary>
    ''' 請求情報入力ボタン入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>

    Public Function inputCheckSeikyuInfoInput() As Boolean
        'トレースログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '相手先CD
        If Me.lblExasAite.Text = "" Then
            Return False
        End If

        'トレースログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True
    End Function

    ''' <summary></summary>
    ''' <remarks>Spreadタイムスケジュールのボタンイベント
    ''' <para>作成情報：2015.12.03 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub fbTimeSchedule_ButtonClicked(sender As Object, e As EditorNotifyEventArgs) Handles fbTimeSchedule.ButtonClicked

        If e.Column = 3 Then
            'リストから削除
            Me.fbTimeSchedule.Sheets(0).RemoveRows(e.Row, 1)
            Dim table As DataTable
            table = dataEXTB0102.PropDsTimeSch.Tables("TIME_SCHEDULE_TBL")
            table.Rows.RemoveAt(e.Row)
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>Spreadチケットのボタンイベント
    ''' <para>作成情報：2015.12.03 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub fbTicket_ButtonClicked(sender As Object, e As EditorNotifyEventArgs) Handles fbTicket.ButtonClicked

        If e.Column = 2 Then
            'リストから削除
            Me.fbTicket.Sheets(0).RemoveRows(e.Row, 1)
            Dim table As DataTable
            table = dataEXTB0102.PropDsTicket.Tables("TICKET_TBL")
            table.Rows.RemoveAt(e.Row)
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>貸出種別：一般貸出選択処理
    ''' <para>作成情報：2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub rdoKashi1_Click(sender As Object, e As EventArgs) Handles rdoKashi1.Click

        ' 一般貸出に変更された場合の確認
        If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
            ' 社内利用⇒一般貸出に変更は請求情報の入力有無でメッセージを表示
            If logicEXTB0103.GetBillpayFlg(dataEXTB0102) = True Then
                'メッセージを出力 
                MsgBox(E2044, MsgBoxStyle.Exclamation, "エラー")
                Me.rdoKashi2.Checked = True
                Exit Sub
            Else
                'メッセージを出力 
                MsgBox(C0014, MsgBoxStyle.Information, "正式予約登録／詳細画面")
                dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
                Exit Sub
            End If
        ElseIf dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI Then
            ' 特例⇒一般貸出に変更は特になし
            dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
            Exit Sub
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>貸出種別：社内利用選択処理
    ''' <para>作成情報：2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub rdoKashi2_Click(sender As Object, e As EventArgs) Handles rdoKashi2.Click

        ' 社内利用に変更された場合の確認
        If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
            ' 請求情報の入力有無でメッセージを表示
            If logicEXTB0103.GetBillpayFlg(dataEXTB0102) = True Then
                'メッセージを出力 
                MsgBox(E2044, MsgBoxStyle.Exclamation, "エラー")
                Me.rdoKashi1.Checked = True
                Exit Sub
            Else
                'メッセージを出力 
                MsgBox(C0014, MsgBoxStyle.Information, "正式予約登録／詳細画面")
                dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
                Exit Sub
            End If
        ElseIf dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI Then
            ' 請求情報の入力有無でメッセージを表示
            If logicEXTB0103.GetBillpayFlg(dataEXTB0102) = True Then
                'メッセージを出力 
                MsgBox(E2044, MsgBoxStyle.Exclamation, "エラー")
                Me.rdoKashi3.Checked = True
                Exit Sub
            Else
                'メッセージを出力 
                MsgBox(C0014, MsgBoxStyle.Information, "正式予約登録／詳細画面")
                dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
                Exit Sub
            End If
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>貸出種別：特例選択処理
    ''' <para>作成情報：2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub rdoKashi3_Click(sender As Object, e As EventArgs) Handles rdoKashi3.Click

        ' 特例に変更された場合の確認
        If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
            ' 社内利用⇒特例に変更は請求情報の入力有無でメッセージを表示
            If logicEXTB0103.GetBillpayFlg(dataEXTB0102) = True Then
                'メッセージを出力 
                MsgBox(E2044, MsgBoxStyle.Exclamation, "エラー")
                Me.rdoKashi2.Checked = True
                Exit Sub
            Else
                'メッセージを出力 
                MsgBox(C0014, MsgBoxStyle.Information, "正式予約登録／詳細画面")
                dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI
                Exit Sub
            End If
        ElseIf dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
            ' 一般貸出⇒特例に変更は特になし
            dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>請求情報リセット処理
    ''' <para>作成情報：2015.12.18 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim strBillypayflg As String = ""

        For i = 0 To sheet.Rows.Count - 1
            If sheet.Cells(i, 1).Value <> "" Then
                strBillypayflg = "1"
            End If
        Next

        If strBillypayflg = "1" Then
            MsgBox(String.Format(CommonEXT.E2046, "請求情報リセット処理"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '更新設定(曜日追加)
        SetRiyobiSpreadData(dataEXTB0102.PropListRiyobi)
        '請求
        dataEXTB0102.propDblTax = logicEXTB0103.GetTaxfirstRiyobi(dataEXTB0102)
        '請求データが存在しない場合、初期データを作成する
        InitBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))

        'プロジェクト情報
        If logicEXTB0103.GetExasRiyoryo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '付帯
        If logicEXTB0103.UpFutai(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        Else
            '設定
            SetSpreadFutai(dataEXTB0102.PropHtFutai)
        End If
        ' 2015.12.21 UPD START↓ h.hagiwara
        'CalcBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        If CalcBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara
        SetSpreadBillPay(dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL"))
        'プロジェクト情報
        If logicEXTB0103.CalcExasRiyoryo(dataEXTB0102) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
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
            With dataEXTB0102
                .PropStrSekininNm = Me.cmbSekininNm.Text   '画面.責任者名
                .PropStrRiyoshaCd = Me.lblRiyoshaCd.Text   '画面.利用者コード
            End With

            'SQL呼出
            If dataEXTB0102.PropStrSekininNm.Equals("") = False Then
                If dataEXTB0102.PropStrRiyoshaCd.Equals("") = False Then
                    If logicEXTB0103.GetSekininshaMailTel(dataEXTB0102) = False Then
                        MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                End If
            End If

            'テキストボックスにセット
            Me.txtSekininMail.Text = dataEXTB0102.PropStrSekininMail        '責任者.メールアドレス
            Me.txtRiyoMobileTel1.Text = dataEXTB0102.PropStrRiyoTel21       '責任者.携帯電話番号

        End If

    End Sub

End Class
