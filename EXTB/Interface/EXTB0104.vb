Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win
Imports FarPoint.Win.Spread.CellType
Imports EXTM

Public Class EXTB0104

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogicEXTB As New CommonLogicEXTB      '共通クラス
    Private mailRegexUtilities As New MailRegexUtilities
    '変数宣言
    Public dataCommon As New CommonDataEXT      '共通データクラス
    Public dataEXTB0104 As New DataEXTB0104     'データクラス
    Public logicEXTB0104 As New LogicEXTB0104   'ロジッククラス
    Public logicEXTZ0202 As New LogicEXTZ0202   'ロジッククラス
    Public logicEXTB0102 As New LogicEXTB0102   'ロジッククラス(仮予約と共通の処理はこちら)    ' 2016.08.08 ADD e.watanabe
    Private strCmbLoadFlg As String           'コンボボックス初期制御フラグ                    ' 2016.08.08 ADD e.watanabe

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTB0104_Load(sender As Object, e As EventArgs) Handles Me.Load

        'FRG初期化
        strCmbLoadFlg = "0"                   '2016.08.08 ADD e.watanabe

        Try
            ''画面項目初期化 2015.11.19 DEL h.hagiwara
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            'End If

            dataEXTB0104.PropStrShisetuKbn = SHISETU_KBN_THEATER
            'コンボボックスの内容設定
            commonLogicEXT.TodohukenLst(Me.cmbRiyoTodo)
            '予約NO無い場合==================================================
            If String.IsNullOrEmpty(dataEXTB0104.PropStrYoyakuNo) = True Then
                'ステータス
                dataEXTB0104.PropStrYoyakuSts = CommonDeclareEXT.CANCEL_STS_MATI
                'ヘッダー
                Me.lblYoyakuNo.Text = String.Empty
                Me.dtpUke.txtDate.Text = String.Empty

                dataEXTB0104.PropStrYoyakuSts = CommonDeclareEXT.YOYAKU_STS_KARI_MI '未確認
                Me.lblStatus.Text = commonLogicEXTB.GetCancelSts(dataEXTB0104.PropStrYoyakuSts)

                Me.lblAddUserCd.Text = String.Empty
                Me.lblAddUserNm.Text = String.Empty
                Me.lblAddUserDate.Text = String.Empty
                Me.lblUpUserCd.Text = String.Empty
                Me.lblUpUserNm.Text = String.Empty
                Me.lblUpUserDate.Text = String.Empty
                '催事
                Me.txtSaiji.Text = String.Empty
                Me.txtShutuen.Text = String.Empty
                '貸し出し種別
                Me.chkKashi.Checked = True
                Me.rdoKashi1.Checked = True
                '利用形状
                Me.chkRiyoType.Checked = True
                Me.rdoRiyoType1.Checked = True
                Me.txtTeiin.Text = ""
                'ワンドリンク
                Me.chkOneDrink.Checked = True
                Me.rdoOneDrink1.Checked = True
                '催事分類
                Me.chkSaijiBunrui.Checked = True
                Me.rdoSaijiBunrui1.Checked = True
                '利用日時
                'Me.btnYoyakuAdd
                'Me.btnYoyakuDel
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
                '特記事項
                Me.txtBiko.Text = String.Empty
                '予約NoないのでDISABLE
                Me.btnKari.Enabled = False
                Me.btnCancel.Enabled = False
                Me.btnResurrec.Enabled = False

                '活性/非活性
                Me.pnlSaijiBunrui.SuspendLayout()
                Me.pnlOneDrink.SuspendLayout()
                Me.pnlRiyoType.SuspendLayout()
                Me.pnlKashi.SuspendLayout()

                '初期利用日の作成
                Dim dataRiyobi As New CommonDataCancel
                Dim lstRiyobi As New ArrayList
                Dim riyobi As Date
                Dim aryStr As String()
                If dataEXTB0104.PropAryStrCancelDate Is Nothing Then
                    Dim curDt As Date = DateTime.Now
                    With dataRiyobi
                        .PropStrShisetuKbn = SHISETU_KBN_THEATER
                        .PropStrStudioKbn = STUDIO_MITEI
                        .PropStrKibobiKbn = "2"
                        .PropStrMiteiFlg = "1"
                        .PropStrRegistFlg = "0" 'DB登録済データ：未登録
                        .PropStrRiyoKeitai = RIYOKEITAI_LOCKOUT
                        .PropStrCancelYm = curDt.ToString("yyyy/MM")
                    End With
                    dataRiyobi.PropIntWakuNo = Nothing
                    If String.IsNullOrEmpty(Me.txtKiboYear.Text) Then
                        Me.txtKiboYear.Text = curDt.ToString("yyyy")
                        Me.txtKiboMonth.Text = curDt.ToString("MM")
                        Me.txtKiboBiko.Text = String.Empty
                    End If
                    lstRiyobi.Add(dataRiyobi)
                    dataRiyobi = New CommonDataCancel
                Else
                    For Each obj As Object In dataEXTB0104.PropAryStrCancelDate
                        If IsArray(obj) Then
                            aryStr = obj
                        Else
                            aryStr = {obj, Nothing}
                        End If
                        riyobi = aryStr(0)
                        With dataRiyobi
                            .PropStrCancelDt = riyobi.ToString(CommonDeclareEXT.FMT_DATE) '日付
                            .PropStrCancelDtDisp = riyobi.ToString(CommonDeclareEXT.FMT_DATE_DISP) '日付
                            .PropStrShisetuKbn = SHISETU_KBN_THEATER  '施設区分(シアター)
                            .PropStrStudioKbn = STUDIO_MITEI   'スタジオ区分(シアター)
                            .PropStrKibobiKbn = "1"
                            .PropStrMiteiFlg = "1"
                            .PropStrRegistFlg = "0" 'DB登録済データ：未登録
                        End With
                        If aryStr.Length < 2 Then
                            dataRiyobi.PropIntWakuNo = Nothing
                        ElseIf String.IsNullOrEmpty(aryStr(1)) = False Then
                            dataRiyobi.PropIntWakuNo = Integer.Parse(aryStr(1))
                        Else
                            dataRiyobi.PropIntWakuNo = Nothing
                        End If
                        If String.IsNullOrEmpty(Me.txtKiboYear.Text) Then
                            Me.txtKiboYear.Text = riyobi.ToString("yyyy")
                            Me.txtKiboMonth.Text = riyobi.ToString("MM")
                            Me.txtKiboBiko.Text = String.Empty
                        End If
                        lstRiyobi.Add(dataRiyobi)
                        dataRiyobi = New CommonDataCancel
                    Next
                End If
                SetSpreadDataRiyobi(lstRiyobi)
                dataEXTB0104.PropListRiyobi = lstRiyobi

                '設定処理
                SetEnabledSettings()
            Else
                '予約NOがある場合==================================================
                'メインデータ取得
                Me.lblYoyakuNo.Text = dataEXTB0104.PropStrYoyakuNo
                logicEXTB0104.GetYoyakuData(dataEXTB0104)
                If logicEXTB0104.GetRiyobiData(dataEXTB0104) = True Then
                    '利用日データ設定
                    SetSpreadDataRiyobi(dataEXTB0104.PropListRiyobi)
                End If
                'ヘッダ
                Me.dtpUke.txtDate.Text = dataEXTB0104.PropStrCanUkeDt
                'Me.btnKariukeCal
                Me.lblStatus.Text = commonLogicEXTB.GetCancelSts(dataEXTB0104.PropStrYoyakuSts)

                Me.lblAddUserCd.Text = dataEXTB0104.PropStrAddUserCd
                Me.lblAddUserNm.Text = dataEXTB0104.PropStrAddUserNm
                Me.lblAddUserDate.Text = dataEXTB0104.PropStrAddDt
                Me.lblUpUserCd.Text = dataEXTB0104.PropStrUpUserCd
                Me.lblUpUserNm.Text = dataEXTB0104.PropStrUpUserNm
                Me.lblUpUserDate.Text = dataEXTB0104.PropStrUpDt
                '催事
                Me.txtSaiji.Text = dataEXTB0104.PropStrSaijiNm
                Me.txtShutuen.Text = dataEXTB0104.PropStrShutsuenNm
                '貸し出し種別
                If dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI Then
                    Me.chkKashi.Checked = True
                    Me.rdoKashi1.Checked = True
                ElseIf dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
                    Me.chkKashi.Checked = False
                    Me.rdoKashi1.Checked = True
                ElseIf dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
                    Me.chkKashi.Checked = False
                    Me.rdoKashi2.Checked = True
                ElseIf dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI Then
                    Me.chkKashi.Checked = False
                    Me.rdoKashi3.Checked = True
                End If
                '利用形状
                If dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MITEI Then
                    Me.rdoRiyoType1.Checked = True
                    Me.chkRiyoType.Checked = True
                ElseIf dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_STAND Then
                    Me.chkRiyoType.Checked = False
                    Me.rdoRiyoType1.Checked = True
                ElseIf dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SEATING Then
                    Me.chkRiyoType.Checked = False
                    Me.rdoRiyoType2.Checked = True
                ElseIf dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MIX Then
                    Me.chkRiyoType.Checked = False
                    Me.rdoRiyoType3.Checked = True
                ElseIf dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SAIJI Then
                    Me.chkRiyoType.Checked = False
                    Me.rdoRiyoType4.Checked = True
                End If
                Me.txtTeiin.Text = dataEXTB0104.PropStrTeiin
                'ワンドリンク
                If dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_MITEI Then
                    Me.chkOneDrink.Checked = True
                    Me.rdoOneDrink1.Checked = True
                ElseIf dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_ARI Then
                    Me.chkOneDrink.Checked = False
                    Me.rdoOneDrink1.Checked = True
                ElseIf dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_NASHI Then
                    Me.chkOneDrink.Checked = False
                    Me.rdoOneDrink2.Checked = True
                End If
                '催事分類
                If dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MITEI Then
                    Me.chkSaijiBunrui.Checked = True
                    Me.rdoSaijiBunrui1.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MUSIC Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui1.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEKI Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui2.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEI Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui3.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_BUSINESS Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui4.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MOVIE Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui5.Checked = True
                ElseIf dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ETC Then
                    Me.chkSaijiBunrui.Checked = False
                    Me.rdoSaijiBunrui6.Checked = True
                End If
                '利用者情報
                Me.lblRiyoshaCd.Text = dataEXTB0104.PropStrRiyoshaCd
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTB0104.PropStrRiyoLvl)
                Me.txtRiyoshaNmKana.Text = dataEXTB0104.PropStrRiyoKana
                Me.txtRiyoshaNm.Text = dataEXTB0104.PropStrRiyoNm
                Me.txtDaihyoNm.Text = dataEXTB0104.PropStrDaihyoNm
                Me.txtSekininBushoNm.Text = dataEXTB0104.PropStrSekininBushoNm
                '責任者名コンボボックス作成
                cmbSekininNm.Items.Clear()                            'コンボボックス初期化  2016.11.04 m.hayabuchi ADD
                If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then
                    Dim list As ArrayList = logicEXTB0104.GetSekininshaList(dataEXTB0104.PropStrRiyoshaCd)
                    For i = 0 To list.Count - 1
                        cmbSekininNm.Items.Add(list(i))
                    Next
                End If
                Me.cmbSekininNm.Text = dataEXTB0104.PropStrSekininNm
                Me.txtSekininMail.Text = dataEXTB0104.PropStrSekininMail
                Me.lblExasAiteNm.Text = dataEXTB0104.PropStrAiteNm
                Me.lblExasAite.Text = dataEXTB0104.PropStrAiteCd
                Me.txtRiyoPost1.Text = dataEXTB0104.PropStrRiyoYubin1
                Me.txtRiyoPost2.Text = dataEXTB0104.PropStrRiyoYubin2
                Me.cmbRiyoTodo.Text = dataEXTB0104.PropStrRiyoTodo
                Me.txtRiyoShiku.Text = dataEXTB0104.PropStrRiyoShiku
                Me.txtRiyoBan.Text = dataEXTB0104.PropStrRiyoBan
                Me.txtRiyoBuild.Text = dataEXTB0104.PropStrRiyoBuild
                Me.txtRiyoTel1.Text = dataEXTB0104.PropStrRiyoTel11
                Me.txtRiyoTel2.Text = dataEXTB0104.PropStrRiyoTel12
                Me.txtRiyoTel3.Text = dataEXTB0104.PropStrRiyoTel13
                Me.txtRiyoNaisen.Text = dataEXTB0104.PropStrRiyoNaisen
                Me.txtRiyoMobileTel1.Text = dataEXTB0104.PropStrRiyoTel21
                'Me.txtRiyoMobileTel2.Text = dataEXTB0104.PropStrRiyoTel22 '2016.11.02 m.hayabuchi DEL 課題No.59
                'Me.txtRiyoMobileTel3.Text = dataEXTB0104.PropStrRiyoTel23 '2016.11.02 m.hayabuchi DEL 課題No.59
                Me.txtRiyoFax1.Text = dataEXTB0104.PropStrRiyoFax11
                Me.txtRiyoFax2.Text = dataEXTB0104.PropStrRiyoFax12
                Me.txtRiyoFax3.Text = dataEXTB0104.PropStrRiyoFax13

                '特記事項
                Me.txtBiko.Text = dataEXTB0104.PropStrBiko
                Me.btnKari.Enabled = True
                If dataEXTB0104.PropStrYoyakuSts = CANCEL_STS_MATI Then
                    Me.btnCancel.Enabled = True
                    Me.btnResurrec.Enabled = False
                ElseIf dataEXTB0104.PropStrYoyakuSts = CANCEL_STS_TORIKESI Then
                    Me.btnCancel.Enabled = False
                    Me.btnResurrec.Enabled = True
                End If

                Me.pnlSaijiBunrui.SuspendLayout()
                Me.pnlOneDrink.SuspendLayout()
                Me.pnlRiyoType.SuspendLayout()
                Me.pnlKashi.SuspendLayout()

                '設定処理
                SetEnabledSettings()

            End If

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
        End Try
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
    ''' 
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
        If Me.chkRiyoType.Checked = True Then
            Me.pnlRiyoType.Enabled = False
        Else
            Me.pnlRiyoType.Enabled = True
        End If
        If Me.chkOneDrink.Checked = True Then
            Me.pnlOneDrink.Enabled = False
        Else
            Me.pnlOneDrink.Enabled = True
        End If
        If Me.chkSaijiBunrui.Checked = True Then
            Me.pnlSaijiBunrui.Enabled = False
        Else
            Me.pnlSaijiBunrui.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 【画面表示用】利用日時リストの値を画面表示するSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadDataRiyobi(ByVal dataList As ArrayList)
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet

        sheet.ColumnCount = 6
        Dim index As New Integer
        Dim lineCnt As New Integer

        index = 0
        lineCnt = 1

        Dim maskcell As New FarPoint.Win.Spread.CellType.MaskCellType()
        maskcell.Mask = "99:99"
        maskcell.NullDisplay = "     "
        Dim lineCntCell As FarPoint.Win.Spread.Cell

        For Each dataRiyobi As CommonDataCancel In dataList

            If dataRiyobi.PropStrKibobiKbn = "1" Then
                sheet.RowCount = lineCnt
                sheet.Cells(index, 0).Value = False
                lineCntCell = sheet.Cells(index, 1)
                lineCntCell.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
                lineCntCell.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                sheet.Cells(index, 1).Value = lineCnt.ToString
                sheet.Cells(index, 1).Locked = True
                sheet.Cells(index, 2).Value = dataRiyobi.PropStrCancelDtDisp
                sheet.Cells(index, 2).Locked = True
                sheet.Cells(index, 4).CellType = maskcell
                sheet.Cells(index, 5).CellType = maskcell
                If dataRiyobi.PropStrMiteiFlg = "0" Then
                    sheet.Cells(index, 3).Value = False
                    sheet.Cells(index, 4).Value = dataRiyobi.PropStrStartTime
                    sheet.Cells(index, 5).Value = dataRiyobi.PropStrEndTime
                Else
                    sheet.Cells(index, 3).Value = True
                    sheet.Cells(index, 4).Value = Nothing
                    sheet.Cells(index, 5).Value = Nothing
                End If

                index = index + 1
                lineCnt = lineCnt + 1
            Else
                sheet.RowCount = lineCnt
                sheet.Cells(index, 0).Value = False
                sheet.Cells(index, 1).Locked = True
                sheet.Cells(index, 1).Value = lineCnt.ToString
                sheet.Cells(index, 2).Locked = True
                sheet.Cells(index, 2).Value = Nothing
                sheet.Cells(index, 3).Value = True
                sheet.Cells(index, 4).Value = Nothing
                sheet.Cells(index, 5).Value = Nothing
                sheet.Cells(index, 4).CellType = maskcell
                sheet.Cells(index, 5).CellType = maskcell
                Me.rdoKibobi2.Checked = True
                Me.txtKiboYear.Text = dataRiyobi.PropStrCancelYm.ToString.Substring(0, 4)
                Me.txtKiboMonth.Text = dataRiyobi.PropStrCancelYm.ToString.Substring(5, 2)
                Me.txtKiboBiko.Text = dataRiyobi.PropStrMemo
            End If
        Next

    End Sub

    ''' <summary>
    ''' 【画面→DB用】Dataに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRiyobiSpreadData(ByRef dataList As ArrayList)
        'SPREAD
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
        Dim index As New Integer
        Dim blnFirst As Boolean = True
        index = 0
        Dim newRiyobiList As New ArrayList
        For Each dataRiyobi As CommonDataCancel In dataList
            If Me.rdoKibobi1.Checked Then
                If sheet.RowCount > 0 Then
                    If String.IsNullOrEmpty(sheet.Cells(index, 2).Value) = False Then
                        If sheet.Cells(index, 3).Value = True Then
                            dataRiyobi.PropStrMiteiFlg = "1"
                        Else
                            dataRiyobi.PropStrMiteiFlg = "0"
                        End If
                        dataRiyobi.PropStrStartTime = sheet.Cells(index, 4).Value
                        dataRiyobi.PropStrEndTime = sheet.Cells(index, 5).Value
                        dataRiyobi.PropStrCancelYm = dataRiyobi.PropStrCancelDt.Substring(0, 7)
                        dataRiyobi.PropIntWakuNo = logicEXTB0104.GetWakuNo(dataRiyobi)
                    End If
                    dataRiyobi.PropStrKibobiKbn = "1"
                    newRiyobiList.Add(dataRiyobi)
                End If
            ElseIf blnFirst = True Then
                dataRiyobi.PropStrKibobiKbn = "2"
                dataRiyobi.PropStrMiteiFlg = "1"
                dataRiyobi.PropIntWakuNo = Nothing
                dataRiyobi.PropStrCancelDt = ""
                dataRiyobi.PropStrCancelDtDisp = ""
                If Me.txtKiboMonth.Text.Length < 2 Then
                    Me.txtKiboMonth.Text = "0" + Me.txtKiboMonth.Text
                End If
                dataRiyobi.PropStrCancelYm = Me.txtKiboYear.Text + "/" + Me.txtKiboMonth.Text
                dataRiyobi.PropStrStartTime = ""
                dataRiyobi.PropStrEndTime = ""
                dataRiyobi.PropStrMemo = Me.txtKiboBiko.Text
                dataRiyobi.PropStrRiyoKeitai = ""
                dataRiyobi.PropStrMiteiFlg = "0"
                dataRiyobi.PropStrStartTime = ""
                dataRiyobi.PropStrEndTime = ""
                blnFirst = False
                newRiyobiList.Add(dataRiyobi)
            End If
            index = index + 1
        Next
        dataList = newRiyobiList
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
    ''' 利用形状
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkRiyoType_CheckedChanged(sender As Object, e As EventArgs) Handles chkRiyoType.CheckedChanged
        If Me.chkRiyoType.Checked = True Then
            Me.pnlRiyoType.Enabled = False
            Me.txtTeiin.Text = ""
        Else
            Me.pnlRiyoType.Enabled = True
            setTeiin()
        End If
    End Sub

    ''' <summary>
    ''' 利用形状定員初期設定処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdoRiyoType1_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRiyoType1.CheckedChanged, rdoRiyoType2.CheckedChanged,
                                                                                        rdoRiyoType3.CheckedChanged, rdoRiyoType4.CheckedChanged
        setTeiin()
    End Sub

    ''' <summary>
    ''' 利用形状:定員
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setTeiin()
        If Me.rdoRiyoType1.Checked = True Then
            Me.txtTeiin.Text = CommonDeclareEXT.PropIntTeiinA.ToString
        ElseIf Me.rdoRiyoType2.Checked = True Then
            Me.txtTeiin.Text = CommonDeclareEXT.PropIntTeiinB.ToString
        End If
    End Sub

    ''' <summary>
    ''' ワンドリンク
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkOneDrink_CheckedChanged(sender As Object, e As EventArgs) Handles chkOneDrink.CheckedChanged
        If Me.chkOneDrink.Checked = True Then
            Me.pnlOneDrink.Enabled = False
        Else
            Me.pnlOneDrink.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 催事分類
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkSaijiBunrui_CheckedChanged(sender As Object, e As EventArgs) Handles chkSaijiBunrui.CheckedChanged
        If Me.chkSaijiBunrui.Checked = True Then
            Me.pnlSaijiBunrui.Enabled = False
        Else
            Me.pnlSaijiBunrui.Enabled = True
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

        Dim frm As New EXTZ0202
        SetRiyobiSpreadData(dataEXTB0104.PropListRiyobi)
        '「利用日追加」画面を表示
        frm.PropStrTorokuKbn = TOUROKU_KBN_CANCEL
        frm.PropLstRiyobi = dataEXTB0104.PropListRiyobi
        frm.PropStrYoyakuNo = dataEXTB0104.PropStrYoyakuNo
        frm.PropStrShisetuKbn = SHISETU_KBN_THEATER
        frm.PropStrStudioKbn = STUDIO_MITEI

        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim lstNewRiyobi As New ArrayList
        For Each dataRiyobi As CommonDataCancel In dataEXTB0104.PropListRiyobi
            If dataRiyobi.PropStrKibobiKbn = "1" And String.IsNullOrEmpty(dataRiyobi.PropStrCancelDt) = False Then
                lstNewRiyobi.Add(dataRiyobi)
            End If
        Next
        Dim comp As New CommonDataCancelCompareter()
        dataEXTB0104.PropListRiyobi = lstNewRiyobi
        dataEXTB0104.PropListRiyobi.Sort(comp)
        SetSpreadDataRiyobi(dataEXTB0104.PropListRiyobi)
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
        Dim lstRiyobi As ArrayList = dataEXTB0104.PropListRiyobi
        Dim dataRiyobi As CommonDataCancel

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
        i = 0
        '削除開始
        Dim j As Integer = 0
        Do While i < rowCnt
            If sheet.Cells(i, 0).Value = True Then
                'DBから制御データ削除
                dataRiyobi = lstRiyobi(i - j)
                If logicEXTZ0202.DeleteCancelCtlData(dataRiyobi) = False Then
                    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                'リストから削除
                lstRiyobi.RemoveAt(i - j)
                j = j + 1
            End If
            i = i + 1
        Loop
        SetSpreadDataRiyobi(dataEXTB0104.PropListRiyobi)
    End Sub

    ''' <summary>
    ''' 仮登録ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnKari_Click(sender As Object, e As EventArgs) Handles btnKari.Click

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
        If inputCheckCondition(True) = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0005), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '画面入力情報をDataに格納
        convertData(dataEXTB0104)

        '利用日データ
        SetRiyobiSpreadData(dataEXTB0104.PropListRiyobi)

        '更新処理
        '予約NOあり
        '予約制御削除
        If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '更新
        If logicEXTB0104.RegYoyakuInfo(dataEXTB0104, True) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表削除
        If logicEXTB0104.DeleteYoyakuList(dataEXTB0104.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTB0104.InsertYoyakuList(dataEXTB0104.PropStrYoyakuNo, dataEXTB0104.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        '正式予約画面
        Dim frm As New EXTB0102
        'パラメータセット
        Dim dataEXTB0102 As New DataEXTB0102
        'データ移し替え
        convertKariData(dataEXTB0104, dataEXTB0102)
        frm.dataEXTB0102 = dataEXTB0102
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
        If inputCheckCondition(False) = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0004), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '画面入力情報をDataに格納
        convertData(dataEXTB0104)
        '利用日データ
        SetRiyobiSpreadData(dataEXTB0104.PropListRiyobi)

        '更新処理
        If String.IsNullOrEmpty(dataEXTB0104.PropStrYoyakuNo) = True Then
            '予約NOなし
            ''予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '登録
            If logicEXTB0104.RegYoyakuInfo(dataEXTB0104, False) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTB0104.InsertYoyakuList(dataEXTB0104.PropStrYoyakuNo, dataEXTB0104.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        Else
            '予約NOあり
            ''予約制御削除
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '更新
            If logicEXTB0104.RegYoyakuInfo(dataEXTB0104, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表削除
            If logicEXTB0104.DeleteYoyakuList(dataEXTB0104.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTB0104.InsertYoyakuList(dataEXTB0104.PropStrYoyakuNo, dataEXTB0104.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTB0104_Load(sender, e)
    End Sub

    ''' <summary>
    ''' 画面入力情報をDataに格納
    ''' </summary>
    ''' <remarks>画面入力情報をDataに格納
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertData(ByRef dataEXTB0104 As DataEXTB0104)
        'ヘッダ
        '仮予約受付USERCDが空の場合設定する
        If String.IsNullOrEmpty(dataEXTB0104.PropStrCanUkeUsercd) = True Then
            dataEXTB0104.PropStrCanUkeUsercd = CommonDeclareEXT.PropComStrUserId
        End If
        dataEXTB0104.PropStrCanUkeDt = Me.dtpUke.txtDate.Text

        '催事
        dataEXTB0104.PropStrSaijiNm = Me.txtSaiji.Text
        dataEXTB0104.PropStrShutsuenNm = Me.txtShutuen.Text

        '貸し出し種別
        If Me.chkKashi.Checked = True Then
            dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI
        Else
            If Me.rdoKashi1.Checked = True Then
                dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
            ElseIf Me.rdoKashi2.Checked = True Then
                dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
            ElseIf Me.rdoKashi3.Checked = True Then
                dataEXTB0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_TOKUREI
            End If
        End If
        '利用形状
        If Me.chkRiyoType.Checked = True Then
            dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MITEI
        Else
            If Me.rdoRiyoType1.Checked = True Then
                dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_STAND
            ElseIf Me.rdoRiyoType2.Checked = True Then
                dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SEATING
            ElseIf Me.rdoRiyoType3.Checked = True Then
                dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_MIX
            ElseIf Me.rdoRiyoType4.Checked = True Then
                dataEXTB0104.PropStrRiyoType = CommonDeclareEXT.RIYO_TYPE_SAIJI
            End If
        End If
        dataEXTB0104.PropStrTeiin = Me.txtTeiin.Text
        'ワンドリンク
        If Me.chkOneDrink.Checked = True Then
            dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_MITEI
        Else
            If Me.rdoOneDrink1.Checked = True Then
                dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_ARI
            ElseIf Me.rdoOneDrink2.Checked = True Then
                dataEXTB0104.PropStrDrinkFlg = CommonDeclareEXT.ONE_DRINK_NASHI
            End If
        End If
        '催事分類
        If Me.chkSaijiBunrui.Checked = True Then
            dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MITEI
        Else
            If Me.rdoSaijiBunrui1.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MUSIC
            ElseIf Me.rdoSaijiBunrui2.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEKI
            ElseIf Me.rdoSaijiBunrui3.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ENGEI
            ElseIf Me.rdoSaijiBunrui4.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_BUSINESS
            ElseIf Me.rdoSaijiBunrui5.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_MOVIE
            ElseIf Me.rdoSaijiBunrui6.Checked = True Then
                dataEXTB0104.PropStrSaijiBunrui = CommonDeclareEXT.SAIJI_ETC
            End If
        End If
        '利用者情報
        dataEXTB0104.PropStrRiyoshaCd = Me.lblRiyoshaCd.Text
        dataEXTB0104.PropStrRiyoKana = Me.txtRiyoshaNmKana.Text
        dataEXTB0104.PropStrRiyoNm = Me.txtRiyoshaNm.Text
        dataEXTB0104.PropStrDaihyoNm = Me.txtDaihyoNm.Text
        dataEXTB0104.PropStrSekininBushoNm = Me.txtSekininBushoNm.Text
        dataEXTB0104.PropStrSekininNm = Me.cmbSekininNm.Text
        dataEXTB0104.PropStrSekininMail = Me.txtSekininMail.Text
        dataEXTB0104.PropStrAiteNm = Me.lblExasAiteNm.Text
        dataEXTB0104.PropStrAiteCd = Me.lblExasAite.Text
        dataEXTB0104.PropStrRiyoYubin1 = Me.txtRiyoPost1.Text
        dataEXTB0104.PropStrRiyoYubin2 = Me.txtRiyoPost2.Text
        dataEXTB0104.PropStrRiyoTodo = Me.cmbRiyoTodo.Text
        dataEXTB0104.PropStrRiyoShiku = Me.txtRiyoShiku.Text
        dataEXTB0104.PropStrRiyoBan = Me.txtRiyoBan.Text
        dataEXTB0104.PropStrRiyoBuild = Me.txtRiyoBuild.Text
        dataEXTB0104.PropStrRiyoTel11 = Me.txtRiyoTel1.Text
        dataEXTB0104.PropStrRiyoTel12 = Me.txtRiyoTel2.Text
        dataEXTB0104.PropStrRiyoTel13 = Me.txtRiyoTel3.Text
        dataEXTB0104.PropStrRiyoNaisen = Me.txtRiyoNaisen.Text
        dataEXTB0104.PropStrRiyoTel21 = Me.txtRiyoMobileTel1.Text
        'dataEXTB0104.PropStrRiyoTel22 = Me.txtRiyoMobileTel2.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTB0104.PropStrRiyoTel23 = Me.txtRiyoMobileTel3.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTB0104.PropStrRiyoFax11 = Me.txtRiyoFax1.Text
        dataEXTB0104.PropStrRiyoFax12 = Me.txtRiyoFax2.Text
        dataEXTB0104.PropStrRiyoFax13 = Me.txtRiyoFax3.Text

        '特記事項
        dataEXTB0104.PropStrBiko = Me.txtBiko.Text

    End Sub

    ''' <summary>
    ''' キャンセル待ち登録情報を仮予約情報に変換
    ''' </summary>
    ''' <remarks>
    ''' <para>作成情報：2015/09/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertKariData(ByRef dataEXTB0104 As DataEXTB0104, ByRef dataEXTB0102 As DataEXTB0102)

        dataEXTB0102.PropStrSaijiNm = dataEXTB0104.PropStrSaijiNm
        dataEXTB0102.PropStrShutsuenNm = dataEXTB0104.PropStrShutsuenNm
        dataEXTB0102.PropStrKashiKind = dataEXTB0104.PropStrKashiKind
        dataEXTB0102.PropStrRiyoType = dataEXTB0104.PropStrRiyoType
        dataEXTB0102.PropStrTeiin = dataEXTB0104.PropStrTeiin
        dataEXTB0102.PropStrDrinkFlg = dataEXTB0104.PropStrDrinkFlg
        dataEXTB0102.PropStrSaijiBunrui = dataEXTB0104.PropStrSaijiBunrui
        dataEXTB0102.PropStrRiyoshaCd = dataEXTB0104.PropStrRiyoshaCd
        dataEXTB0102.PropStrRiyoLvl = dataEXTB0104.PropStrRiyoLvl
        dataEXTB0102.PropStrRiyoKana = dataEXTB0104.PropStrRiyoKana
        dataEXTB0102.PropStrRiyoNm = dataEXTB0104.PropStrRiyoNm
        dataEXTB0102.PropStrDaihyoNm = dataEXTB0104.PropStrDaihyoNm
        dataEXTB0102.PropStrSekininBushoNm = dataEXTB0104.PropStrSekininBushoNm
        dataEXTB0102.PropStrSekininNm = dataEXTB0104.PropStrSekininNm
        dataEXTB0102.PropStrSekininMail = dataEXTB0104.PropStrSekininMail
        dataEXTB0102.PropStrAiteNm = dataEXTB0104.PropStrAiteNm
        dataEXTB0102.PropStrAiteCd = dataEXTB0104.PropStrAiteCd
        dataEXTB0102.PropStrRiyoYubin1 = dataEXTB0104.PropStrRiyoYubin1
        dataEXTB0102.PropStrRiyoYubin2 = dataEXTB0104.PropStrRiyoYubin2
        dataEXTB0102.PropStrRiyoTodo = dataEXTB0104.PropStrRiyoTodo
        dataEXTB0102.PropStrRiyoShiku = dataEXTB0104.PropStrRiyoShiku
        dataEXTB0102.PropStrRiyoBan = dataEXTB0104.PropStrRiyoBan
        dataEXTB0102.PropStrRiyoBuild = dataEXTB0104.PropStrRiyoBuild
        dataEXTB0102.PropStrRiyoTel11 = dataEXTB0104.PropStrRiyoTel11
        dataEXTB0102.PropStrRiyoTel12 = dataEXTB0104.PropStrRiyoTel12
        dataEXTB0102.PropStrRiyoTel13 = dataEXTB0104.PropStrRiyoTel13
        dataEXTB0102.PropStrRiyoNaisen = dataEXTB0104.PropStrRiyoNaisen
        dataEXTB0102.PropStrRiyoTel21 = dataEXTB0104.PropStrRiyoTel21
        'dataEXTB0102.PropStrRiyoTel22 = dataEXTB0104.PropStrRiyoTel22 '2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTB0102.PropStrRiyoTel23 = dataEXTB0104.PropStrRiyoTel23 '2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTB0102.PropStrRiyoFax11 = dataEXTB0104.PropStrRiyoFax11
        dataEXTB0102.PropStrRiyoFax12 = dataEXTB0104.PropStrRiyoFax12
        dataEXTB0102.PropStrRiyoFax13 = dataEXTB0104.PropStrRiyoFax13
        dataEXTB0102.PropStrCancelNo = dataEXTB0104.PropStrYoyakuNo
        dataEXTB0102.PropStrBiko = dataEXTB0104.PropStrBiko

        dataEXTB0102.PropListRiyobi = New ArrayList
        Dim dataRiyobi As CommonDataRiyobi
        For Each dataRiyobiCancel As CommonDataCancel In dataEXTB0104.PropListRiyobi
            If dataRiyobiCancel.PropStrKibobiKbn = "1" Then
                dataRiyobi = New CommonDataRiyobi
                dataRiyobi.PropIntSeq = Nothing
                dataRiyobi.PropStrShisetuKbn = dataRiyobiCancel.PropStrShisetuKbn
                dataRiyobi.PropStrStudioKbn = dataRiyobiCancel.PropStrStudioKbn
                dataRiyobi.PropStrYoyakuDt = dataRiyobiCancel.PropStrCancelDt
                dataRiyobi.PropStrYoyakuDtDisp = dataRiyobiCancel.PropStrCancelDtDisp
                dataRiyobi.PropStrStartTime = dataRiyobiCancel.PropStrStartTime
                dataRiyobi.PropStrEndTime = dataRiyobiCancel.PropStrEndTime
                dataRiyobi.PropStrYoyakuNo = Nothing
                dataRiyobi.PropStrRiyoKeitai = dataRiyobiCancel.PropStrRiyoKeitai
                dataRiyobi.PropStrMiteiFlg = dataRiyobiCancel.PropStrMiteiFlg
                dataRiyobi.PropIntTanka = Nothing
                dataRiyobi.PropDblBairitu = Nothing
                dataRiyobi.PropIntSu = Nothing
                dataRiyobi.PropIntRiyoKin = Nothing
                dataRiyobi.PropStrRegistFlg = "0"
                dataEXTB0102.PropListRiyobi.Add(dataRiyobi)
            End If
        Next
        'ステータス
        dataEXTB0102.PropStrYoyakuSts = YOYAKU_STS_KARI_MI

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
            '定員
            If commonValidate.IsHalfNmb(Me.txtTeiin.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "定員"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            'スプレッド
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
            Dim index As New Integer
            Dim lineCnt As New Integer

            If Me.rdoKibobi1.Checked Then
                index = 0
                lineCnt = 1
                If 0 = dataEXTB0104.PropListRiyobi.Count Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望日時"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If 0 = sheet.RowCount Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望日時"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                For Each dataRiyobi As CommonDataCancel In dataEXTB0104.PropListRiyobi
                    If sheet.Cells(index, 3).Value = False Then
                        If String.IsNullOrEmpty(sheet.Cells(index, 4).Value) Then
                            MsgBox(String.Format(CommonDeclareEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                            Return False
                        End If
                        If String.IsNullOrEmpty(sheet.Cells(index, 5).Value) Then
                            MsgBox(String.Format(CommonDeclareEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                            Return False
                        End If
                    End If
                    If commonValidate.IsHalfNmb(sheet.Cells(index, 4).Value) = False Then
                        MsgBox(String.Format(CommonDeclareEXT.E0003, "開始時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                    If commonValidate.IsHalfNmb(sheet.Cells(index, 5).Value) = False Then
                        MsgBox(String.Format(CommonDeclareEXT.E0003, "終了時間"), MsgBoxStyle.Exclamation, "エラー")
                        Return False
                    End If
                Next
            ElseIf Me.rdoKibobi2.Checked Then
                If String.IsNullOrEmpty(Me.txtKiboYear.Text) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望年・利用希望月"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If String.IsNullOrEmpty(Me.txtKiboMonth.Text) Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望年・利用希望月"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
            End If

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
            For Each dataRiyobi As CommonDataCancel In dataEXTB0104.PropListRiyobi
                '予約数超過
                If logicEXTZ0202.CheckCancelWaitRegister(TOUROKU_KBN_KARI, dataRiyobi) = False Then
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
    ''' キャンセル待ち取消ボタン処理
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
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '更新
            dataEXTB0104.PropStrYoyakuSts = CANCEL_STS_TORIKESI
            If logicEXTB0104.RegYoyakuInfo(dataEXTB0104, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.PropStrDeleteKbn = TORIKESHI_KBN_DELETE Then
            '物理削除
            '予約制御削除
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTB0104.DeleteYoyakuList(dataEXTB0104.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTB0104.DeleteYoyaku(dataEXTB0104.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.DialogResult = System.Windows.Forms.DialogResult.OK Then
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' キャンセル待ち復活ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnResurrec_Click(sender As Object, e As EventArgs) Handles btnResurrec.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.01.25 ADD START↓ h.hagiwara 復活時のデータチェック追加
        '入力チェック
        If inputCheckMain() = False Then
            Exit Sub
        End If
        If inputCheckCondition(False) = False Then
            Exit Sub
        End If

        '画面入力情報をDataに格納
        convertData(dataEXTB0104)
        '利用日データ
        SetRiyobiSpreadData(dataEXTB0104.PropListRiyobi)

        ' 2016.01.25 ADD END↑ h.hagiwara 復活時のデータチェック追加

        '更新
        dataEXTB0104.PropStrYoyakuSts = CANCEL_STS_MATI
        If logicEXTB0104.RegYoyakuInfo(dataEXTB0104, True) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        ' 2016.01.25 ADD START↓ h.hagiwara キャンセル枠再設定
        '予約表削除
        If logicEXTB0104.DeleteYoyakuList(dataEXTB0104.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTB0104.InsertYoyakuList(dataEXTB0104.PropStrYoyakuNo, dataEXTB0104.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        ' 2016.01.25 ADD END↑ h.hagiwara キャンセル枠再設定
        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTB0104_Load(sender, e)
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
                dataEXTB0104.PropStrRiyoLvl = .PropParamRiyoLvl
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTB0104.PropStrRiyoLvl)
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

                Dim list As ArrayList = logicEXTB0104.GetSekininshaList(Me.lblRiyoshaCd.Text)
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
        If logicEXTZ0202.DeleteYoyakuCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
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
    Private Sub EXTB0104_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '予約受付制御削除
        If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_THEATER) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
        End If
    End Sub

    ''' <summary>
    ''' 希望日有無の変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdoKibobi_CheckedChanged(sender As Object, e As EventArgs) Handles rdoKibobi1.CheckedChanged, rdoKibobi2.CheckedChanged
        If Me.rdoKibobi1.Checked Then
            Me.txtKiboYear.Enabled = False
            Me.txtKiboMonth.Enabled = False
            Me.txtKiboBiko.Enabled = False
            Me.fpRiyobi.Enabled = True
            Me.btnYoyakuAdd.Enabled = True
            Me.btnYoyakuDel.Enabled = True
        End If
        If Me.rdoKibobi2.Checked Then
            Me.txtKiboYear.Enabled = True
            Me.txtKiboMonth.Enabled = True
            Me.txtKiboBiko.Enabled = True
            Me.fpRiyobi.Enabled = False
            Me.btnYoyakuAdd.Enabled = False
            Me.btnYoyakuDel.Enabled = False
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
            With dataEXTB0104
                .PropStrSekininNm = Me.cmbSekininNm.Text     '画面.責任者名
                .PropStrRiyoshaCd = Me.lblRiyoshaCd.Text     '画面.利用者コード
            End With

            'SQL呼出
            If dataEXTB0104.PropStrSekininNm.Equals("") = False Then
                If dataEXTB0104.PropStrRiyoshaCd.Equals("") = False Then
                    If logicEXTB0104.GetSekininshaMailTel(dataEXTB0104) = False Then
                        MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                End If
            End If

            'テキストボックスにセット
            Me.txtSekininMail.Text = dataEXTB0104.PropStrSekininMail        '責任者.メールアドレス
            Me.txtRiyoMobileTel1.Text = dataEXTB0104.PropStrRiyoTel21       '責任者.携帯電話番号

        End If

    End Sub

End Class
