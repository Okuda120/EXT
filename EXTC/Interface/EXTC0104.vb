Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win
Imports FarPoint.Win.Spread.CellType
Imports EXTM

Public Class EXTC0104

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogicEXTC As New CommonLogicEXTC      '共通クラス
    Private mailRegexUtilities As New MailRegexUtilities
    '変数宣言
    Public dataCommon As New CommonDataEXT      '共通データクラス
    Public dataEXTC0104 As New DataEXTC0104     'データクラス
    Public logicEXTC0104 As New LogicEXTC0104   'ロジッククラス
    Public logicEXTZ0202 As New LogicEXTZ0202   'ロジッククラス

    Private strCmbLoadFlg As String           'コンボボックス初期制御フラグ                    ' 2016.11.4 m.hayabuchi add

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTC0104_Load(sender As Object, e As EventArgs) Handles Me.Load

        'FRG初期化
        strCmbLoadFlg = "0"                   '2016.11.4 m.hayabuchi add

        Try
            ''画面項目初期化 2015.11.19 DEL h.hagiwara
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            'End If

            dataEXTC0104.PropStrShisetuKbn = SHISETU_KBN_STUDIO
            'コンボボックスの内容設定
            commonLogicEXT.TodohukenLst(Me.cmbRiyoTodo)
            '予約NO無い場合==================================================
            If String.IsNullOrEmpty(dataEXTC0104.PropStrYoyakuNo) = True Then
                'ステータス
                dataEXTC0104.PropStrYoyakuSts = CommonDeclareEXT.CANCEL_STS_MATI
                'ヘッダー
                Me.lblYoyakuNo.Text = String.Empty
                Me.dtpUke.txtDate.Text = String.Empty

                dataEXTC0104.PropStrYoyakuSts = CommonDeclareEXT.YOYAKU_STS_KARI_MI '未確認
                Me.lblStatus.Text = commonLogicEXTC.GetCancelSts(dataEXTC0104.PropStrYoyakuSts)

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
                Me.chkStudio.Checked = True
                Me.rdoStudio1.Checked = True
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
                'Me.txtRiyoMobileTel2.Text = String.Empty '2016.11.4 m.hayabuchi DEL 課題No.59
                'Me.txtRiyoMobileTel3.Text = String.Empty '2016.11.4 m.hayabuchi DEL 課題No.59
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
                Me.pnlKashi.SuspendLayout()
                Me.pnlStudio.SuspendLayout()
                Me.pnlOpe.SuspendLayout()

                '初期利用日の作成
                Dim dataRiyobi As New CommonDataCancel
                Dim lstRiyobi As New ArrayList
                Dim riyobi As Date
                Dim aryStr As String()
                If dataEXTC0104.PropAryStrCancelDate.Count = 0 Then
                    Dim curDt As Date = DateTime.Now
                    With dataRiyobi
                        .PropStrShisetuKbn = SHISETU_KBN_STUDIO
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
                    For Each obj As Object In dataEXTC0104.PropAryStrCancelDate
                        If IsArray(obj) Then
                            aryStr = obj
                        Else
                            aryStr = {obj, Nothing}
                        End If
                        riyobi = aryStr(0)
                        With dataRiyobi
                            .PropStrCancelDt = riyobi.ToString(CommonDeclareEXT.FMT_DATE) '日付
                            .PropStrCancelDtDisp = riyobi.ToString(CommonDeclareEXT.FMT_DATE_DISP) '日付
                            .PropStrShisetuKbn = SHISETU_KBN_STUDIO  '施設区分(シアター)
                            .PropStrStudioKbn = STUDIO_MITEI   'スタジオ区分(シアター)
                            .PropStrKibobiKbn = "1"
                            .PropStrMiteiFlg = "1"
                            .PropStrRegistFlg = "0" 'DB登録済データ：未登録
                            .PropStrRiyoKeitai = RIYOKEITAI_LOCKOUT
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
                dataEXTC0104.PropListRiyobi = lstRiyobi

                '設定処理
                SetEnabledSettings()
            Else
                '予約NOがある場合==================================================
                'メインデータ取得
                Me.lblYoyakuNo.Text = dataEXTC0104.PropStrYoyakuNo
                logicEXTC0104.GetYoyakuData(dataEXTC0104)
                If logicEXTC0104.GetRiyobiData(dataEXTC0104) = True Then
                    '利用日データ設定
                    SetSpreadDataRiyobi(dataEXTC0104.PropListRiyobi)
                End If
                'ヘッダ
                Me.dtpUke.txtDate.Text = dataEXTC0104.PropStrCanUkeDt
                'Me.btnKariukeCal
                Me.lblStatus.Text = commonLogicEXTC.GetCancelSts(dataEXTC0104.PropStrYoyakuSts)

                Me.lblAddUserCd.Text = dataEXTC0104.PropStrAddUserCd
                Me.lblAddUserNm.Text = dataEXTC0104.PropStrAddUserNm
                Me.lblAddUserDate.Text = dataEXTC0104.PropStrAddDt
                Me.lblUpUserCd.Text = dataEXTC0104.PropStrUpUserCd
                Me.lblUpUserNm.Text = dataEXTC0104.PropStrUpUserNm
                Me.lblUpUserDate.Text = dataEXTC0104.PropStrUpDt
                '催事
                Me.txtShutuen.Text = dataEXTC0104.PropStrShutsuenNm
                '貸し出し種別
                If dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI Then
                    Me.chkKashi.Checked = True
                    Me.rdoKashi1.Checked = True
                ElseIf dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN Then
                    Me.chkKashi.Checked = False
                    Me.rdoKashi1.Checked = True
                ElseIf dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
                    Me.chkKashi.Checked = False
                    Me.rdoKashi2.Checked = True
                End If
                '利用スタジオ
                If dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_MITEI Then
                    Me.rdoStudio1.Checked = True
                    Me.chkStudio.Checked = True
                ElseIf dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_201 Then
                    Me.chkStudio.Checked = False
                    Me.rdoStudio1.Checked = True
                ElseIf dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_202 Then
                    Me.chkStudio.Checked = False
                    Me.rdoStudio2.Checked = True
                ElseIf dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_HOUSE_LOCK Then
                    Me.chkStudio.Checked = False
                    Me.rdoStudio3.Checked = True
                End If
                '音響オペレーター
                If dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_MITEI Then
                    Me.chkOpe.Checked = True
                    Me.rdoOpe1.Checked = True
                ElseIf dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_ARI Then
                    Me.chkOpe.Checked = False
                    Me.rdoOpe1.Checked = True
                ElseIf dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_NASHI Then
                    Me.chkOpe.Checked = False
                    Me.rdoOpe2.Checked = True
                End If
                '利用者情報
                Me.lblRiyoshaCd.Text = dataEXTC0104.PropStrRiyoshaCd
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTC0104.PropStrRiyoLvl)
                Me.txtRiyoshaNmKana.Text = dataEXTC0104.PropStrRiyoKana
                Me.txtRiyoshaNm.Text = dataEXTC0104.PropStrRiyoNm
                Me.txtDaihyoNm.Text = dataEXTC0104.PropStrDaihyoNm
                Me.txtSekininBushoNm.Text = dataEXTC0104.PropStrSekininBushoNm
                '責任者名コンボボックス作成
                cmbSekininNm.Items.Clear()                            'コンボボックス初期化  2016.11.04 m.hayabuchi ADD
                If String.IsNullOrEmpty(Me.lblRiyoshaCd.Text) = False Then
                    Dim list As ArrayList = logicEXTC0104.GetSekininshaList(dataEXTC0104.PropStrRiyoshaCd)
                    For i = 0 To list.Count - 1
                        cmbSekininNm.Items.Add(list(i))
                    Next
                End If
                Me.cmbSekininNm.Text = dataEXTC0104.PropStrSekininNm
                Me.txtSekininMail.Text = dataEXTC0104.PropStrSekininMail
                Me.lblExasAiteNm.Text = dataEXTC0104.PropStrAiteNm
                Me.lblExasAite.Text = dataEXTC0104.PropStrAiteCd
                Me.txtRiyoPost1.Text = dataEXTC0104.PropStrRiyoYubin1
                Me.txtRiyoPost2.Text = dataEXTC0104.PropStrRiyoYubin2
                Me.cmbRiyoTodo.Text = dataEXTC0104.PropStrRiyoTodo
                Me.txtRiyoShiku.Text = dataEXTC0104.PropStrRiyoShiku
                Me.txtRiyoBan.Text = dataEXTC0104.PropStrRiyoBan
                Me.txtRiyoBuild.Text = dataEXTC0104.PropStrRiyoBuild
                Me.txtRiyoTel1.Text = dataEXTC0104.PropStrRiyoTel11
                Me.txtRiyoTel2.Text = dataEXTC0104.PropStrRiyoTel12
                Me.txtRiyoTel3.Text = dataEXTC0104.PropStrRiyoTel13
                Me.txtRiyoNaisen.Text = dataEXTC0104.PropStrRiyoNaisen
                Me.txtRiyoMobileTel1.Text = dataEXTC0104.PropStrRiyoTel21
                'Me.txtRiyoMobileTel2.Text = dataEXTC0104.PropStrRiyoTel22 '2016.11.2 m.hayabuchi DEL 課題No.59
                'Me.txtRiyoMobileTel3.Text = dataEXTC0104.PropStrRiyoTel23 '2016.11.2 m.hayabuchi DEL 課題No.59
                Me.txtRiyoFax1.Text = dataEXTC0104.PropStrRiyoFax11
                Me.txtRiyoFax2.Text = dataEXTC0104.PropStrRiyoFax12
                Me.txtRiyoFax3.Text = dataEXTC0104.PropStrRiyoFax13

                '特記事項
                Me.txtBiko.Text = dataEXTC0104.PropStrBiko
                Me.btnKari.Enabled = True
                If dataEXTC0104.PropStrYoyakuSts = CANCEL_STS_MATI Then
                    Me.btnCancel.Enabled = True
                    Me.btnResurrec.Enabled = False
                ElseIf dataEXTC0104.PropStrYoyakuSts = CANCEL_STS_TORIKESI Then
                    Me.btnCancel.Enabled = False
                    Me.btnResurrec.Enabled = True
                End If

                Me.pnlKashi.SuspendLayout()
                Me.pnlStudio.SuspendLayout()
                Me.pnlOpe.SuspendLayout()

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
        If Me.chkStudio.Checked = True Then
            Me.pnlStudio.Enabled = False
        Else
            Me.pnlStudio.Enabled = True
        End If
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
            Else
                'カラムの設定
                sheet.RowCount = lineCnt
                sheet.Cells(index, 0).Value = False
                sheet.Cells(index, 1).Locked = True
                sheet.Cells(index, 1).Value = lineCnt.ToString
                sheet.Cells(index, 2).Value = Nothing
                sheet.Cells(index, 2).Locked = True
                sheet.Cells(index, 3).Value = True
                sheet.Cells(index, 4).CellType = cmb
                sheet.Cells(index, 4).Value = Nothing
                sheet.Cells(index, 5).CellType = maskcell
                sheet.Cells(index, 6).CellType = maskcell
                sheet.Cells(index, 5).Value = Nothing
                sheet.Cells(index, 6).Value = Nothing
                '選択状態
                Me.rdoKibobi2.Checked = True
                Me.txtKiboYear.Text = dataRiyobi.PropStrCancelYm.ToString.Substring(0, 4)
                Me.txtKiboMonth.Text = dataRiyobi.PropStrCancelYm.ToString.Substring(5, 2)
                Me.txtKiboBiko.Text = dataRiyobi.PropStrMemo
                Me.fpRiyobi.Enabled = False
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
                    ' 2016.02.19 DEL START↓ h.hagiwara 位置変更
                    dataRiyobi.PropStrKibobiKbn = "1"
                    dataRiyobi.PropStrStudioKbn = dataEXTC0104.PropStrStudioKbn
                    ' 2016.02.19 DEL END↑ h.hagiwara 位置変更
                    If String.IsNullOrEmpty(sheet.Cells(index, 2).Value) = False Then
                        If sheet.Cells(index, 3).Value = True Then
                            dataRiyobi.PropStrMiteiFlg = "1"
                        Else
                            dataRiyobi.PropStrMiteiFlg = "0"
                        End If
                        dataRiyobi.PropStrRiyoKeitai = sheet.Cells(index, 4).Value
                        dataRiyobi.PropStrStartTime = sheet.Cells(index, 5).Value
                        dataRiyobi.PropStrEndTime = sheet.Cells(index, 6).Value
                        dataRiyobi.PropStrCancelYm = dataRiyobi.PropStrCancelDt.Substring(0, 7)
                        dataRiyobi.PropIntWakuNo = logicEXTC0104.GetWakuNo(dataRiyobi)
                    End If
                    ' 2016.02.19 DEL START↓ h.hagiwara 位置変更
                    'dataRiyobi.PropStrKibobiKbn = "1"
                    'dataRiyobi.PropStrStudioKbn = dataEXTC0104.PropStrStudioKbn
                    ' 2016.02.19 DEL END↑ h.hagiwara 位置変更
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
                dataRiyobi.PropStrStudioKbn = dataEXTC0104.PropStrStudioKbn
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
    ''' スタジオ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkStudio_CheckedChanged(sender As Object, e As EventArgs) Handles chkStudio.CheckedChanged
        If Me.chkStudio.Checked = True Then
            Me.pnlStudio.Enabled = False
        Else
            Me.pnlStudio.Enabled = True
        End If
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
        convertData(dataEXTC0104)
        Dim frm As New EXTZ0202
        SetRiyobiSpreadData(dataEXTC0104.PropListRiyobi)
        '「利用日追加」画面を表示
        frm.PropStrTorokuKbn = TOUROKU_KBN_CANCEL
        frm.PropLstRiyobi = dataEXTC0104.PropListRiyobi
        frm.PropStrYoyakuNo = dataEXTC0104.PropStrYoyakuNo
        frm.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.PropStrStudioKbn = dataEXTC0104.PropStrStudioKbn

        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim lstNewRiyobi As New ArrayList
        For Each dataRiyobi As CommonDataCancel In dataEXTC0104.PropListRiyobi
            If dataRiyobi.PropStrKibobiKbn = "1" And String.IsNullOrEmpty(dataRiyobi.PropStrCancelDt) = False Then
                lstNewRiyobi.Add(dataRiyobi)
            End If
        Next
        Dim comp As New CommonDataCancelCompareter()
        dataEXTC0104.PropListRiyobi = lstNewRiyobi
        dataEXTC0104.PropListRiyobi.Sort(comp)
        SetSpreadDataRiyobi(dataEXTC0104.PropListRiyobi)
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
        Dim lstRiyobi As ArrayList = dataEXTC0104.PropListRiyobi
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
        SetSpreadDataRiyobi(dataEXTC0104.PropListRiyobi)
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
        '画面入力情報をDataに格納
        convertData(dataEXTC0104)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0104.PropListRiyobi)
        If inputCheckCondition(True) = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0005), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '更新処理
        '予約NOあり
        '予約制御削除
        If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '更新
        If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, True) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表削除
        If logicEXTC0104.DeleteYoyakuList(dataEXTC0104.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTC0104.InsertYoyakuList(dataEXTC0104.PropStrYoyakuNo, dataEXTC0104.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        '正式予約画面
        Dim frm As New EXTC0102
        'パラメータセット
        Dim dataEXTC0102 As New DataEXTC0102
        'データ移し替え
        convertKariData(dataEXTC0104, dataEXTC0102)
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
        convertData(dataEXTC0104)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0104.PropListRiyobi)
        If inputCheckCondition(False) = False Then
            Exit Sub
        End If
        '確認処理
        If MsgBox(String.Format(CommonEXT.C0004), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            Return
        End If

        '更新処理
        If String.IsNullOrEmpty(dataEXTC0104.PropStrYoyakuNo) = True Then
            '予約NOなし
            ''予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '登録
            If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, False) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTC0104.InsertYoyakuList(dataEXTC0104.PropStrYoyakuNo, dataEXTC0104.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        Else
            '予約NOあり
            ''予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            'If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
            '    MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            '    Return
            'End If
            '更新
            If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表削除
            If logicEXTC0104.DeleteYoyakuList(dataEXTC0104.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約表登録
            If logicEXTC0104.InsertYoyakuList(dataEXTC0104.PropStrYoyakuNo, dataEXTC0104.PropListRiyobi) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除 2015.11.19 UPD h.hagiwara 削除位置変更
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTC0104_Load(sender, e)
    End Sub

    ''' <summary>
    ''' 画面入力情報をDataに格納
    ''' </summary>
    ''' <remarks>画面入力情報をDataに格納
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertData(ByRef dataEXTC0104 As DataEXTC0104)
        'ヘッダ
        '仮予約受付USERCDが空の場合設定する
        If String.IsNullOrEmpty(dataEXTC0104.PropStrCanUkeUsercd) = True Then
            dataEXTC0104.PropStrCanUkeUsercd = CommonDeclareEXT.PropComStrUserId
        End If
        dataEXTC0104.PropStrCanUkeDt = Me.dtpUke.txtDate.Text

        '催事
        dataEXTC0104.PropStrShutsuenNm = Me.txtShutuen.Text

        '貸し出し種別
        If Me.chkKashi.Checked = True Then
            dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_MITEI
        Else
            If Me.rdoKashi1.Checked = True Then
                dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_IPPAN
            ElseIf Me.rdoKashi2.Checked = True Then
                dataEXTC0104.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE
            End If
        End If
        'STUDIO
        If Me.chkStudio.Checked = True Then
            dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_MITEI
        Else
            If Me.rdoStudio1.Checked = True Then
                dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_201
            ElseIf Me.rdoStudio2.Checked = True Then
                dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_202
            ElseIf Me.rdoStudio3.Checked = True Then
                dataEXTC0104.PropStrStudioKbn = CommonDeclareEXT.STUDIO_HOUSE_LOCK
            End If
        End If
        '音響
        If Me.chkOpe.Checked = True Then
            dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_MITEI
        Else
            If Me.rdoOpe1.Checked = True Then
                dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_ARI
            ElseIf Me.rdoOpe2.Checked = True Then
                dataEXTC0104.PropStrOnkyoOpe = CommonDeclareEXT.ONKYO_OPE_NASHI
            End If
        End If
        '利用者情報
        dataEXTC0104.PropStrRiyoshaCd = Me.lblRiyoshaCd.Text
        dataEXTC0104.PropStrRiyoKana = Me.txtRiyoshaNmKana.Text
        dataEXTC0104.PropStrRiyoNm = Me.txtRiyoshaNm.Text
        dataEXTC0104.PropStrDaihyoNm = Me.txtDaihyoNm.Text
        dataEXTC0104.PropStrSekininBushoNm = Me.txtSekininBushoNm.Text
        dataEXTC0104.PropStrSekininNm = Me.cmbSekininNm.Text
        dataEXTC0104.PropStrSekininMail = Me.txtSekininMail.Text
        dataEXTC0104.PropStrAiteNm = Me.lblExasAiteNm.Text
        dataEXTC0104.PropStrAiteCd = Me.lblExasAite.Text
        dataEXTC0104.PropStrRiyoYubin1 = Me.txtRiyoPost1.Text
        dataEXTC0104.PropStrRiyoYubin2 = Me.txtRiyoPost2.Text
        dataEXTC0104.PropStrRiyoTodo = Me.cmbRiyoTodo.Text
        dataEXTC0104.PropStrRiyoShiku = Me.txtRiyoShiku.Text
        dataEXTC0104.PropStrRiyoBan = Me.txtRiyoBan.Text
        dataEXTC0104.PropStrRiyoBuild = Me.txtRiyoBuild.Text
        dataEXTC0104.PropStrRiyoTel11 = Me.txtRiyoTel1.Text
        dataEXTC0104.PropStrRiyoTel12 = Me.txtRiyoTel2.Text
        dataEXTC0104.PropStrRiyoTel13 = Me.txtRiyoTel3.Text
        dataEXTC0104.PropStrRiyoNaisen = Me.txtRiyoNaisen.Text
        dataEXTC0104.PropStrRiyoTel21 = Me.txtRiyoMobileTel1.Text
        'dataEXTC0104.PropStrRiyoTel22 = Me.txtRiyoMobileTel2.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTC0104.PropStrRiyoTel23 = Me.txtRiyoMobileTel3.Text '2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTC0104.PropStrRiyoFax11 = Me.txtRiyoFax1.Text
        dataEXTC0104.PropStrRiyoFax12 = Me.txtRiyoFax2.Text
        dataEXTC0104.PropStrRiyoFax13 = Me.txtRiyoFax3.Text
        dataEXTC0104.PropStrOnkyoNm = Me.txtOnkyoNm.Text
        dataEXTC0104.PropStrOnkyoTantoNm = Me.txtOnkyoTantoNm.Text
        dataEXTC0104.PropStrOnkyoTel11 = Me.txtOnkyoTel11.Text
        dataEXTC0104.PropStrOnkyoTel12 = Me.txtOnkyoTel12.Text
        dataEXTC0104.PropStrOnkyoTel13 = Me.txtOnkyoTel13.Text
        dataEXTC0104.PropStrOnkyoNaisen = Me.txtOnkyoNaisen.Text
        dataEXTC0104.PropStrOnkyoFax11 = Me.txtOnkyoFax11.Text
        dataEXTC0104.PropStrOnkyoFax12 = Me.txtOnkyoFax12.Text
        dataEXTC0104.PropStrOnkyoFax13 = Me.txtOnkyoFax13.Text
        dataEXTC0104.PropStrOnkyoMail = Me.txtOnkyoMail.Text

        '特記事項
        dataEXTC0104.PropStrBiko = Me.txtBiko.Text

    End Sub

    ''' <summary>
    ''' キャンセル待ち登録情報を仮予約情報に変換
    ''' </summary>
    ''' <remarks>
    ''' <para>作成情報：2015/09/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub convertKariData(ByRef dataEXTC0104 As DataEXTC0104, ByRef dataEXTC0102 As DataEXTC0102)

        dataEXTC0102.PropStrShutsuenNm = dataEXTC0104.PropStrShutsuenNm
        dataEXTC0102.PropStrKashiKind = dataEXTC0104.PropStrKashiKind
        dataEXTC0102.PropStrStudioKbn = dataEXTC0104.PropStrStudioKbn
        dataEXTC0102.PropStrOnkyoOpe = dataEXTC0104.PropStrOnkyoOpe
        dataEXTC0102.PropStrRiyoshaCd = dataEXTC0104.PropStrRiyoshaCd
        dataEXTC0102.PropStrRiyoLvl = dataEXTC0104.PropStrRiyoLvl
        dataEXTC0102.PropStrRiyoKana = dataEXTC0104.PropStrRiyoKana
        dataEXTC0102.PropStrRiyoNm = dataEXTC0104.PropStrRiyoNm
        dataEXTC0102.PropStrDaihyoNm = dataEXTC0104.PropStrDaihyoNm
        dataEXTC0102.PropStrSekininBushoNm = dataEXTC0104.PropStrSekininBushoNm
        dataEXTC0102.PropStrSekininNm = dataEXTC0104.PropStrSekininNm
        dataEXTC0102.PropStrSekininMail = dataEXTC0104.PropStrSekininMail
        dataEXTC0102.PropStrAiteNm = dataEXTC0104.PropStrAiteNm
        dataEXTC0102.PropStrAiteCd = dataEXTC0104.PropStrAiteCd
        dataEXTC0102.PropStrRiyoYubin1 = dataEXTC0104.PropStrRiyoYubin1
        dataEXTC0102.PropStrRiyoYubin2 = dataEXTC0104.PropStrRiyoYubin2
        dataEXTC0102.PropStrRiyoTodo = dataEXTC0104.PropStrRiyoTodo
        dataEXTC0102.PropStrRiyoShiku = dataEXTC0104.PropStrRiyoShiku
        dataEXTC0102.PropStrRiyoBan = dataEXTC0104.PropStrRiyoBan
        dataEXTC0102.PropStrRiyoBuild = dataEXTC0104.PropStrRiyoBuild
        dataEXTC0102.PropStrRiyoTel11 = dataEXTC0104.PropStrRiyoTel11
        dataEXTC0102.PropStrRiyoTel12 = dataEXTC0104.PropStrRiyoTel12
        dataEXTC0102.PropStrRiyoTel13 = dataEXTC0104.PropStrRiyoTel13
        dataEXTC0102.PropStrRiyoNaisen = dataEXTC0104.PropStrRiyoNaisen
        dataEXTC0102.PropStrRiyoTel21 = dataEXTC0104.PropStrRiyoTel21
        'dataEXTC0102.PropStrRiyoTel22 = dataEXTC0104.PropStrRiyoTel22 '2016.11.2 m.hayabuchi DEL 課題No.59
        'dataEXTC0102.PropStrRiyoTel23 = dataEXTC0104.PropStrRiyoTel23 '2016.11.2 m.hayabuchi DEL 課題No.59
        dataEXTC0102.PropStrRiyoFax11 = dataEXTC0104.PropStrRiyoFax11
        dataEXTC0102.PropStrRiyoFax12 = dataEXTC0104.PropStrRiyoFax12
        dataEXTC0102.PropStrRiyoFax13 = dataEXTC0104.PropStrRiyoFax13
        dataEXTC0102.PropStrCancelNo = dataEXTC0104.PropStrYoyakuNo
        dataEXTC0102.PropStrBiko = dataEXTC0104.PropStrBiko
        dataEXTC0102.PropStrOnkyoNm = dataEXTC0104.PropStrOnkyoNm
        dataEXTC0102.PropStrOnkyoTantoNm = dataEXTC0104.PropStrOnkyoTantoNm
        dataEXTC0102.PropStrOnkyoTel11 = dataEXTC0104.PropStrOnkyoTel11
        dataEXTC0102.PropStrOnkyoTel12 = dataEXTC0104.PropStrOnkyoTel12
        dataEXTC0102.PropStrOnkyoTel13 = dataEXTC0104.PropStrOnkyoTel13
        dataEXTC0102.PropStrOnkyoNaisen = dataEXTC0104.PropStrOnkyoNaisen
        dataEXTC0102.PropStrOnkyoFax11 = dataEXTC0104.PropStrOnkyoFax11
        dataEXTC0102.PropStrOnkyoFax12 = dataEXTC0104.PropStrOnkyoFax12
        dataEXTC0102.PropStrOnkyoFax13 = dataEXTC0104.PropStrOnkyoFax13
        dataEXTC0102.PropStrOnkyoMail = dataEXTC0104.PropStrOnkyoMail

        dataEXTC0102.PropListRiyobi = New ArrayList
        Dim dataRiyobi As CommonDataRiyobi
        For Each dataRiyobiCancel As CommonDataCancel In dataEXTC0104.PropListRiyobi
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
                dataEXTC0102.PropListRiyobi.Add(dataRiyobi)
            End If
        Next
        'ステータス
        dataEXTC0102.PropStrYoyakuSts = YOYAKU_STS_KARI

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
            'アーティスト名
            If String.IsNullOrEmpty(Me.txtShutuen.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "アーティスト名"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            'スプレッド
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fpRiyobi.ActiveSheet
            Dim index As New Integer
            Dim lineCnt As New Integer

            If Me.rdoKibobi1.Checked Then
                '希望日がある場合はスタジオ必須
                If Me.chkStudio.Checked Then
                    MsgBox(String.Format(CommonDeclareEXT.E0002, "利用スタジオ"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                'SPREAD
                index = 0
                lineCnt = 1
                If 0 = dataEXTC0104.PropListRiyobi.Count Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望日時"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If 0 = sheet.RowCount Then
                    MsgBox(String.Format(CommonDeclareEXT.E0001, "利用希望日時"), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                For Each dataRiyobi As CommonDataCancel In dataEXTC0104.PropListRiyobi
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
            If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel1.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号1"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            ' 2016.11.2 m.hayabuchi DEL Start 課題No.59
            'If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel2.Text) = False Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号2"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            'If commonValidate.IsHalfNmb(Me.txtRiyoMobileTel3.Text) = False Then
            '    MsgBox(String.Format(CommonDeclareEXT.E0003, "携帯番号3"), MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            ' 2016.11.2 m.hayabuchi DEL End 課題No.59
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
            For Each dataRiyobi As CommonDataCancel In dataEXTC0104.PropListRiyobi
                '予約数超過
                If logicEXTZ0202.CheckCancelWaitRegister(TOUROKU_KBN_KARI, dataRiyobi) = False Then
                    MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                End If
                If listRiyobi.Contains(dataRiyobi.PropStrCancelDt) Then
                    MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                    Return False
                Else
                    listRiyobi.Add(dataRiyobi.PropStrCancelDt)
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
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '更新
            dataEXTC0104.PropStrYoyakuSts = CANCEL_STS_TORIKESI
            If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, True) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
        End If
        If frm.PropStrDeleteKbn = TORIKESHI_KBN_DELETE Then
            '物理削除
            '予約制御削除
            If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTC0104.DeleteYoyakuList(dataEXTC0104.PropStrYoyakuNo) = False Then
                MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                Return
            End If
            '予約制御削除
            If logicEXTC0104.DeleteYoyaku(dataEXTC0104.PropStrYoyakuNo) = False Then
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
        convertData(dataEXTC0104)
        '利用日データ
        SetRiyobiSpreadData(dataEXTC0104.PropListRiyobi)
        ' 2016.01.25 ADD END↑ h.hagiwara 復活時のデータチェック追加

        '更新
        dataEXTC0104.PropStrYoyakuSts = CANCEL_STS_MATI
        If logicEXTC0104.RegYoyakuInfo(dataEXTC0104, True) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        ' 2016.01.25 ADD START↓ h.hagiwara キャンセル枠再設定
        '予約表削除
        If logicEXTC0104.DeleteYoyakuList(dataEXTC0104.PropStrYoyakuNo) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '予約表登録
        If logicEXTC0104.InsertYoyakuList(dataEXTC0104.PropStrYoyakuNo, dataEXTC0104.PropListRiyobi) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        ' 2016.01.25 ADD END↑ h.hagiwara キャンセル枠再設定

        '完了メッセージ
        MsgBox(String.Format(CommonDeclareEXT.I0002, "更新"), MsgBoxStyle.Information, "完了")
        '表示処理
        EXTC0104_Load(sender, e)
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
                dataEXTC0104.PropStrRiyoLvl = .PropParamRiyoLvl
                Me.lblRiyoshaLvl.Text = commonLogicEXT.getRiyoshalvlNm(dataEXTC0104.PropStrRiyoLvl)
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
                Dim list As ArrayList = logicEXTC0104.GetSekininshaList(Me.lblRiyoshaCd.Text)
                cmbSekininNm.Items.Clear()
                ' 2016.08.09 UPD START e.watanabe　' ELSE追加（0件の場合）変更
                'For i = 0 To list.Count - 1
                '    cmbSekininNm.Items.Add(list(i))
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
    Private Sub EXTC0104_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '予約受付制御削除
        If logicEXTZ0202.DeleteCancelCtlShisetuData(CommonDeclareEXT.SHISETU_KBN_STUDIO) = False Then
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
            With dataEXTC0104
                .PropStrSekininNm = Me.cmbSekininNm.Text   '画面.責任者名
                .PropStrRiyoshaCd = Me.lblRiyoshaCd.Text   '画面.利用者コード
            End With

            'SQL呼出
            If dataEXTC0104.PropStrSekininNm.Equals("") = False Then
                If dataEXTC0104.PropStrRiyoshaCd.Equals("") = False Then
                    If logicEXTC0104.GetSekininshaMailTel(dataEXTC0104) = False Then
                        MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                End If
            End If

            'テキストボックスにセット
            Me.txtSekininMail.Text = dataEXTC0104.PropStrSekininMail        '責任者.メールアドレス
            Me.txtRiyoMobileTel1.Text = dataEXTC0104.PropStrRiyoTel21       '責任者.携帯電話番号

        End If

    End Sub

End Class
