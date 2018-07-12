Imports Common
Imports CommonEXT
Imports FarPoint.Win.Spread

''' <summary>
''' EXTZ0204
''' </summary>
''' <remarks>EXASプロジェクト設定
''' <para>作成情報：2015/09/01 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0204

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Public dataEXTZ0204 As New DataEXTZ0204     'データクラス
    Public logicEXTZ0204 As New LogicEXTZ0204   'ロジッククラス
    Public CommonValidation As New CommonValidation     'CommonValidation                                      ' 2015.12.11 ADD h.hagiwara

    Public seikyubi As String

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0204_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dataEXTZ0204.PropBlnChangeFlg = False
            'SPREAD 請求額
            Dim RowSeikyu As DataRow = dataEXTZ0204.PropDrBillReq
            Dim shtSeikyu As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
            shtSeikyu.ColumnCount = 7
            shtSeikyu.Columns(0).Locked = True
            shtSeikyu.Columns(1).Locked = True
            shtSeikyu.Columns(2).Locked = True
            shtSeikyu.Columns(3).Locked = True
            shtSeikyu.Columns(4).Locked = True
            shtSeikyu.Columns(5).Locked = True
            shtSeikyu.Columns(6).Locked = True
            shtSeikyu.RowCount = 1
            shtSeikyu.Cells(0, 0).Value = RowSeikyu("seikyu_naiyo_nm")
            shtSeikyu.Cells(0, 1).Value = RowSeikyu("kakutei_kin")
            shtSeikyu.Cells(0, 2).Value = RowSeikyu("chosei_kin")
            shtSeikyu.Cells(0, 3).Value = RowSeikyu("shokei")
            shtSeikyu.Cells(0, 4).Value = dataEXTZ0204.PropDblTax
            shtSeikyu.Cells(0, 5).Value = RowSeikyu("tax_kin")
            shtSeikyu.Cells(0, 6).Value = RowSeikyu("seikyu_kin")
            Me.txtTitle1.Text = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title1")
            Me.txtTitle2.Text = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title2")

            'SPREAD 利用料
            Dim seikyuNaiyo As String = RowSeikyu("seikyu_naiyo")
            Dim TblRiyoryo As DataTable = dataEXTZ0204.PropDtExasRiyoryo
            Dim shtRiyoryo As FarPoint.Win.Spread.SheetView = Me.fbRiyoryo.ActiveSheet
            shtRiyoryo.ColumnCount = 18
            shtRiyoryo.Columns(12).Visible = False
            shtRiyoryo.Columns(13).Visible = False
            shtRiyoryo.Columns(14).Visible = False
            shtRiyoryo.Columns(15).Visible = False
            shtRiyoryo.Columns(16).Visible = False
            shtRiyoryo.Columns(17).Visible = False
            shtRiyoryo.Columns(0).Locked = True
            shtRiyoryo.Columns(1).Locked = True
            shtRiyoryo.Columns(2).Locked = True
            shtRiyoryo.Columns(3).Locked = True
            shtRiyoryo.Columns(4).Locked = True
            shtRiyoryo.Columns(9).Locked = True
            shtRiyoryo.Columns(10).Locked = True
            'Dim riyoKei As Integer = 0                                    ' 2015.12.21 UPD h.hagiwara
            'Dim riyoTaxKei As Integer = 0                                 ' 2015.12.21 UPD h.hagiwara
            Dim riyoKei As Long = 0                                        ' 2015.12.21 UPD h.hagiwara
            Dim riyoTaxKei As Long = 0                                     ' 2015.12.21 UPD h.hagiwara
            Dim i As Long = 0
            Dim row As DataRow
            Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()
            'numCellKeijyo.MaximumValue = 99999999                                ' 2015.12.21 UPD h.hagiwara
            numCellKeijyo.MaximumValue = 99999999999                              ' 2015.12.21 UPD h.hagiwara
            numCellKeijyo.ShowSeparator = True
            numCellKeijyo.DecimalPlaces = 0
            Dim numCellTax As New FarPoint.Win.Spread.CellType.NumberCellType()
            'numCellTax.MaximumValue = 9999999                                   ' 2015.12.21 UPD h.hagiwara
            numCellTax.MaximumValue = 999999999                                  ' 2015.12.21 UPD h.hagiwara
            numCellTax.ShowSeparator = True
            numCellTax.DecimalPlaces = 0
            Dim txtCellTitle As New FarPoint.Win.Spread.CellType.TextCellType()
            txtCellTitle.MaxLength = 25
            If seikyuNaiyo = SEIKYU_NAIYOU_RIYO Or seikyuNaiyo = SEIKYU_NAIYOU_RIYOFUTAI Then
                shtRiyoryo.RowCount = TblRiyoryo.Rows.Count
                Do While i < TblRiyoryo.Rows.Count
                    row = TblRiyoryo.Rows(i)
                    shtRiyoryo.Cells(i, 0).Value = commonLogicEXT.convYmDateStr(commonLogicEXT.DbNullToNothing(row, "riyo_ym"))
                    shtRiyoryo.Cells(i, 1).Value = commonLogicEXT.DbNullToNothing(row, "kamoku_nm")
                    shtRiyoryo.Cells(i, 2).Value = commonLogicEXT.DbNullToNothing(row, "saimoku_nm")
                    shtRiyoryo.Cells(i, 3).Value = commonLogicEXT.DbNullToNothing(row, "uchi_nm")
                    shtRiyoryo.Cells(i, 4).Value = commonLogicEXT.DbNullToNothing(row, "shosai_nm")
                    'shtRiyoryo.Cells(i, 5).CellType = numCellKeijyo
                    shtRiyoryo.Cells(i, 5).Value = Long.Parse(commonLogicEXT.DbNullToNothing(row, "keijo_kin"))
                    'shtRiyoryo.Cells(i, 6).CellType = numCellTax
                    If String.IsNullOrEmpty(shtRiyoryo.Cells(i, 5).Value) = False Then
                        'shtRiyoryo.Cells(i, 6).Value = Integer.Parse(Math.Round(Integer.Parse(shtRiyoryo.Cells(i, 5).Value) * dataEXTZ0204.PropDblTax))                 ' 2015.12.21 UPD h.hagiwara
                        'shtRiyoryo.Cells(i, 6).Value = Long.Parse(Math.Round(Long.Parse(shtRiyoryo.Cells(i, 5).Value) * dataEXTZ0204.PropDblTax))                       ' 2015.12.21 UPD h.hagiwara
                        shtRiyoryo.Cells(i, 6).Value = Long.Parse(commonLogicEXT.DbNullToNothing(row, "tax_kin"))                                                        ' 2015.12.21 UPD h.hagiwara
                    Else
                        shtRiyoryo.Cells(i, 6).Value = Nothing
                    End If
                    riyoKei = riyoKei + shtRiyoryo.Cells(i, 5).Value
                    riyoTaxKei = riyoTaxKei + shtRiyoryo.Cells(i, 6).Value
                    ' 2015.12.07 UPD START↓ h.hagiwara
                    'shtRiyoryo.Cells(i, 7).CellType = txtCellTitle
                    'shtRiyoryo.Cells(i, 7).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo1")
                    'shtRiyoryo.Cells(i, 8).CellType = txtCellTitle
                    'shtRiyoryo.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo2")
                    If commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_input_flg") Is Nothing Then
                        shtRiyoryo.Cells(i, 7).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 7).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title1")
                        shtRiyoryo.Cells(i, 8).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title2")
                    ElseIf commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_input_flg") = "0" Then
                        shtRiyoryo.Cells(i, 7).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 7).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title1")
                        shtRiyoryo.Cells(i, 8).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title2")
                    Else
                        shtRiyoryo.Cells(i, 7).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 7).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo1")
                        shtRiyoryo.Cells(i, 8).CellType = txtCellTitle
                        shtRiyoryo.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo2")
                    End If
                    ' 2015.12.07 UPD END↑ h.hagiwara
                    shtRiyoryo.Cells(i, 9).Value = commonLogicEXT.DbNullToNothing(row, "event_nm")
                    shtRiyoryo.Cells(i, 10).Value = commonLogicEXT.DbNullToNothing(row, "content_uchi_nm")
                    shtRiyoryo.Cells(i, 11).Value = commonLogicEXT.DbNullToNothing(row, "kamoku_cd")
                    shtRiyoryo.Cells(i, 12).Value = commonLogicEXT.DbNullToNothing(row, "saimoku_cd")
                    shtRiyoryo.Cells(i, 13).Value = commonLogicEXT.DbNullToNothing(row, "uchi_cd")
                    shtRiyoryo.Cells(i, 14).Value = commonLogicEXT.DbNullToNothing(row, "shosai_cd")
                    shtRiyoryo.Cells(i, 15).Value = commonLogicEXT.DbNullToNothing(row, "content_cd")
                    shtRiyoryo.Cells(i, 16).Value = commonLogicEXT.DbNullToNothing(row, "content_uchi_cd")
                    shtRiyoryo.Cells(i, 17).Value = commonLogicEXT.DbNullToNothing(row, "riyo_ym")
                    i = i + 1
                Loop
            End If

            'SPREAD 付帯
            Dim TblFutai As DataTable = dataEXTZ0204.PropDtExasFutai
            Dim shtFutai As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
            shtFutai.ColumnCount = 21
            shtFutai.Columns(14).Visible = False
            shtFutai.Columns(15).Visible = False
            shtFutai.Columns(16).Visible = False
            shtFutai.Columns(17).Visible = False
            shtFutai.Columns(18).Visible = False
            shtFutai.Columns(19).Visible = False
            shtFutai.Columns(20).Visible = False
            shtFutai.Columns(0).Locked = True
            shtFutai.Columns(1).Locked = True
            shtFutai.Columns(2).Locked = True
            shtFutai.Columns(3).Locked = True
            shtFutai.Columns(4).Locked = True
            shtFutai.Columns(5).Locked = True
            shtFutai.Columns(10).Locked = True
            shtFutai.Columns(11).Locked = True
            ' 2015.12.21 ADD START↓ h.hagiwara
            Dim numCellKeijyoFutai As New FarPoint.Win.Spread.CellType.NumberCellType()
            numCellKeijyo.MaximumValue = 99999999
            numCellKeijyo.ShowSeparator = True
            numCellKeijyo.DecimalPlaces = 0
            Dim numCellTaxFutai As New FarPoint.Win.Spread.CellType.NumberCellType()
            numCellTax.MaximumValue = 9999999
            numCellTax.ShowSeparator = True
            numCellTax.DecimalPlaces = 0
            ' 2015.12.21 ADD END↑ h.hagiwara
            'Dim futaiKei As Integer = 0                                    ' 2015.12.21 UPD h.hagiwara
            'Dim futaiTaxKei As Integer = 0                                 ' 2015.12.21 UPD h.hagiwara
            Dim futaiKei As Long = 0                                        ' 2015.12.21 UPD h.hagiwara
            Dim futaiTaxKei As Long = 0                                     ' 2015.12.21 UPD h.hagiwara
            i = 0
            If seikyuNaiyo = SEIKYU_NAIYOU_FUTAI Or seikyuNaiyo = SEIKYU_NAIYOU_RIYOFUTAI Then
                shtFutai.RowCount = TblFutai.Rows.Count
                Do While i < TblFutai.Rows.Count
                    row = TblFutai.Rows(i)
                    Dim ary As Array = row.ItemArray
                    shtFutai.Cells(i, 0).Value = commonLogicEXT.convYmDateStr(commonLogicEXT.DbNullToNothing(row, "riyo_ym"))
                    shtFutai.Cells(i, 1).Value = commonLogicEXT.DbNullToNothing(row, "shukei_grp")
                    shtFutai.Cells(i, 2).Value = commonLogicEXT.DbNullToNothing(row, "kamoku_nm")
                    shtFutai.Cells(i, 3).Value = commonLogicEXT.DbNullToNothing(row, "saimoku_nm")
                    shtFutai.Cells(i, 4).Value = commonLogicEXT.DbNullToNothing(row, "uchi_nm")
                    shtFutai.Cells(i, 5).Value = commonLogicEXT.DbNullToNothing(row, "shosai_nm")
                    ''shtFutai.Cells(i, 6).CellType = numCellKeijyo                                         ' 2015.12.21 UPD h.hagiwara
                    'shtFutai.Cells(i, 6).CellType = numCellKeijyoFutai                                     ' 2015.12.21 UPD h.hagiwara
                    'shtFutai.Cells(i, 6).Value = Integer.Parse(commonLogicEXT.DbNullToNothing(row, "keijo_kin"))        ' 2015.12.21 UPD h.hagiwara
                    shtFutai.Cells(i, 6).Value = Long.Parse(commonLogicEXT.DbNullToNothing(row, "keijo_kin"))            ' 2015.12.21 UPD h.hagiwara
                    ''shtFutai.Cells(i, 7).CellType = numCellTax                                            ' 2015.12.21 UPD h.hagiwara
                    'shtFutai.Cells(i, 7).CellType = numCellTaxFutai                                        ' 2015.12.21 UPD h.hagiwara
                    'shtFutai.Cells(i, 7).Value = Integer.Parse(commonLogicEXT.DbNullToNothing(row, "tax_kin"))         ' 2015.12.21 UPD h.hagiwara
                    shtFutai.Cells(i, 7).Value = Long.Parse(commonLogicEXT.DbNullToNothing(row, "tax_kin"))             ' 2015.12.21 UPD h.hagiwara
                    futaiKei = futaiKei + shtFutai.Cells(i, 6).Value
                    futaiTaxKei = futaiTaxKei + shtFutai.Cells(i, 7).Value
                    ' 2015.12.07 UPD START↓ h.hagiwara
                    'shtFutai.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo1")
                    'shtFutai.Cells(i, 8).CellType = txtCellTitle
                    'shtFutai.Cells(i, 9).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo2")
                    'shtFutai.Cells(i, 9).CellType = txtCellTitle
                    If commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_input_flg") Is Nothing Then
                        shtFutai.Cells(i, 8).CellType = txtCellTitle
                        shtFutai.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title1")
                        shtFutai.Cells(i, 9).CellType = txtCellTitle
                        shtFutai.Cells(i, 9).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title2")
                    ElseIf commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_input_flg") = "0" Then
                        shtFutai.Cells(i, 8).CellType = txtCellTitle
                        shtFutai.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title1")
                        shtFutai.Cells(i, 9).CellType = txtCellTitle
                        shtFutai.Cells(i, 9).Value = commonLogicEXT.DbNullToNothing(RowSeikyu, "seikyu_title2")
                    Else
                        shtFutai.Cells(i, 8).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo1")
                        shtFutai.Cells(i, 8).CellType = txtCellTitle
                        shtFutai.Cells(i, 9).Value = commonLogicEXT.DbNullToNothing(row, "tekiyo2")
                        shtFutai.Cells(i, 9).CellType = txtCellTitle
                    End If
                    ' 2015.12.07 UPD END↑ h.hagiwara
                    shtFutai.Cells(i, 10).Value = commonLogicEXT.DbNullToNothing(row, "event_nm")
                    shtFutai.Cells(i, 11).Value = commonLogicEXT.DbNullToNothing(row, "content_uchi_nm")
                    shtFutai.Cells(i, 14).Value = commonLogicEXT.DbNullToNothing(row, "kamoku_cd")
                    shtFutai.Cells(i, 15).Value = commonLogicEXT.DbNullToNothing(row, "saimoku_cd")
                    shtFutai.Cells(i, 16).Value = commonLogicEXT.DbNullToNothing(row, "uchi_cd")
                    shtFutai.Cells(i, 17).Value = commonLogicEXT.DbNullToNothing(row, "shosai_cd")
                    shtFutai.Cells(i, 18).Value = commonLogicEXT.DbNullToNothing(row, "content_cd")
                    shtFutai.Cells(i, 19).Value = commonLogicEXT.DbNullToNothing(row, "content_uchi_cd")
                    shtFutai.Cells(i, 20).Value = commonLogicEXT.DbNullToNothing(row, "riyo_ym")
                    i = i + 1
                Loop
            End If

            Me.lblSagakuKin.Text = shtSeikyu.Cells(0, 3).Value - riyoKei - futaiKei
            Me.lblTaxSagaku.Text = shtSeikyu.Cells(0, 5).Value - riyoTaxKei - futaiTaxKei

            'fbBillPay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                                  ' 2015.12.21 UPD h.hagiwara
            fbBillPay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                                ' 2015.12.21 UPD h.hagiwara
            fbBillPay.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
            'fbRiyoryo.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                                  ' 2015.12.21 UPD h.hagiwara
            fbRiyoryo.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                                ' 2015.12.21 UPD h.hagiwara
            fbRiyoryo.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
            'fbFutai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                                    ' 2015.12.21 UPD h.hagiwara
            fbFutai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                                  ' 2015.12.21 UPD h.hagiwara
            fbFutai.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
            'seq発行済の場合、入力完了不可
            'If IsDBNull(RowSeikyu("seq")) = False And dataEXTZ0204.propStrSeikyuInputFlg = "1" Then
            If IsDBNull(RowSeikyu("seikyu_irai_no")) = False And dataEXTZ0204.propStrSeikyuInputFlg = "1" Then
                Me.btnComplate.Enabled = False
            End If

            ' 背景色設定
            Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
            If CommonEXT.PropConfigrationFlg = "1" Then
                ' 検証機の場合には背景イメージも表示しない
                Me.BackgroundImage = Nothing
            End If

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' スプレッド利用料のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRiyoryo_Click(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbRiyoryo.ButtonClicked
        '表示ボタンが押された場合
        If e.Column = 11 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim frm As New EXTZ0205
            Dim dt As Date = seikyubi
            frm.dataEXTZ0205.PropStrRiyobi = dt.ToString("yyyyMMdd")
            frm.ShowDialog()
            If frm.dataEXTZ0205.PropBlnChangeFlg Then
                Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbRiyoryo.ActiveSheet
                sheet.Cells(e.Row, 9).Value = frm.dataEXTZ0205.PropStrResPrjNm
                sheet.Cells(e.Row, 10).Value = frm.dataEXTZ0205.PropStrResUchiNm
                sheet.Cells(e.Row, 15).Value = frm.dataEXTZ0205.PropStrResPrjCd
                sheet.Cells(e.Row, 16).Value = frm.dataEXTZ0205.PropStrResUchiCd
            End If
        End If
    End Sub

    ''' <summary>
    ''' スプレッド利用料のボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFutai_Click(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbFutai.ButtonClicked
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        '下にコピーボタンが押された場合
        If e.Column = 12 Then
            '最下行はスキップ
            If sheet.Rows.Count > e.Row + 1 Then
                Dim cpIndex As Integer = e.Row + 1
                ' 2015．12.07 DEL START↓ h.hagiwara コピーはコンテンツのみ
                'sheet.Cells(cpIndex, 6).Value = sheet.Cells(e.Row, 6).Value
                'sheet.Cells(cpIndex, 7).Value = sheet.Cells(e.Row, 7).Value
                'sheet.Cells(cpIndex, 8).Value = sheet.Cells(e.Row, 8).Value
                'sheet.Cells(cpIndex, 9).Value = sheet.Cells(e.Row, 9).Value
                ' 2015．12.07 DEL END↑ h.hagiwara コピーはコンテンツのみ
                sheet.Cells(cpIndex, 10).Value = sheet.Cells(e.Row, 10).Value
                sheet.Cells(cpIndex, 11).Value = sheet.Cells(e.Row, 11).Value
                ' 2015．12.07 DEL START↓ h.hagiwara コピーはコンテンツのみ
                'sheet.Cells(cpIndex, 14).Value = sheet.Cells(e.Row, 14).Value
                'sheet.Cells(cpIndex, 15).Value = sheet.Cells(e.Row, 15).Value
                'sheet.Cells(cpIndex, 16).Value = sheet.Cells(e.Row, 16).Value
                'sheet.Cells(cpIndex, 17).Value = sheet.Cells(e.Row, 17).Value
                ' 2015．12.07 DEL END↑ h.hagiwara コピーはコンテンツのみ
                sheet.Cells(cpIndex, 18).Value = sheet.Cells(e.Row, 18).Value
                sheet.Cells(cpIndex, 19).Value = sheet.Cells(e.Row, 19).Value
                calcSagaku()
            End If
        End If
        '表示ボタンが押された場合
        If e.Column = 13 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim frm As New EXTZ0205
            Dim dt As Date = seikyubi
            frm.dataEXTZ0205.PropStrRiyobi = dt.ToString("yyyyMMdd")
            frm.ShowDialog()
            'If String.IsNullOrEmpty(frm.dataEXTZ0205.PropStrResPrjNm) = False Then      ' 2015.11.25 UPD h.hagiwara 利用料時と同じ判定にする
            If frm.dataEXTZ0205.PropBlnChangeFlg Then                                    ' 2015.11.25 UPD h.hagiwara 利用料時と同じ判定にする
                sheet.Cells(e.Row, 10).Value = frm.dataEXTZ0205.PropStrResPrjNm
                sheet.Cells(e.Row, 11).Value = frm.dataEXTZ0205.PropStrResUchiNm
                sheet.Cells(e.Row, 18).Value = frm.dataEXTZ0205.PropStrResPrjCd
                sheet.Cells(e.Row, 19).Value = frm.dataEXTZ0205.PropStrResUchiCd
            End If
        End If
    End Sub

    ''' <summary>
    ''' 利用料セルOnchange処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtRiyoryo_Change(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fbRiyoryo.Change
        If e.Column = 5 Or e.Column = 6 Then
            calcSagaku()
        End If
    End Sub

    ''' <summary>
    ''' 付帯設備セルOnchange処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtFutairyo_Change(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fbFutai.Change
        If e.Column = 6 Or e.Column = 7 Then
            calcSagaku()
        End If
    End Sub

    ''' <summary>
    ''' 差額計算処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub calcSagaku()
        Dim shtSeikyu As FarPoint.Win.Spread.SheetView = Me.fbBillPay.ActiveSheet
        Dim shtRiyoryo As FarPoint.Win.Spread.SheetView = Me.fbRiyoryo.ActiveSheet
        Dim shtFutai As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet

        'Dim riyoKei As Integer = 0                                 ' 2015.12.21 UPD h.hagiwara
        'Dim riyoTaxKei As Integer = 0                              ' 2015.12.21 UPD h.hagiwara
        Dim riyoKei As Long = 0                                     ' 2015.12.21 UPD h.hagiwara
        Dim riyoTaxKei As Long = 0
        Dim i As Integer = 0
        Do While i < shtRiyoryo.RowCount
            riyoKei = riyoKei + shtRiyoryo.Cells(i, 5).Value
            riyoTaxKei = riyoTaxKei + shtRiyoryo.Cells(i, 6).Value
            i = i + 1
        Loop
        i = 0
        Do While i < shtFutai.RowCount
            riyoKei = riyoKei + shtFutai.Cells(i, 6).Value
            riyoTaxKei = riyoTaxKei + shtFutai.Cells(i, 7).Value
            i = i + 1
        Loop
        Me.lblSagakuKin.Text = shtSeikyu.Cells(0, 3).Value - riyoKei
        Me.lblTaxSagaku.Text = shtSeikyu.Cells(0, 5).Value - riyoTaxKei
    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If inputCheckMain() = False Then
            Exit Sub
        End If
        Dim shtRiyoryo As FarPoint.Win.Spread.SheetView = Me.fbRiyoryo.ActiveSheet
        Dim shtFutai As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        Dim TblRiyoryo As DataTable = dataEXTZ0204.PropDtExasRiyoryo
        Dim TblFutai As DataTable = dataEXTZ0204.PropDtExasFutai
        Dim row As DataRow
        Dim i As Integer = 0
        Dim bigTaxkin As Long = 0                                                ' 2016.01.26 ADD h.hagiwara
        '利用料の処理
        Do While i < shtRiyoryo.RowCount
            row = TblRiyoryo.Rows(i)
            row("keijo_kin") = commonLogicEXT.convInsNumStr(shtRiyoryo.Cells(i, 5).Value)
            row("tax_kin") = commonLogicEXT.convInsNumStr(shtRiyoryo.Cells(i, 6).Value)
            row("tekiyo1") = shtRiyoryo.Cells(i, 7).Value
            row("tekiyo2") = shtRiyoryo.Cells(i, 8).Value
            row("event_nm") = shtRiyoryo.Cells(i, 9).Value
            row("content_uchi_nm") = shtRiyoryo.Cells(i, 10).Value
            row("content_cd") = shtRiyoryo.Cells(i, 15).Value
            row("content_uchi_cd") = shtRiyoryo.Cells(i, 16).Value
            ' 2016.01.26 ADD START↓ h.hagiwara
            If String.IsNullOrEmpty(shtRiyoryo.Cells(i, 6).Value) Then
            Else
                bigTaxkin += CLng(shtRiyoryo.Cells(i, 6).Value)
            End If
            ' 2016.01.26 ADD END↑ h.hagiwara
            i = i + 1
        Loop
        '付帯設備の処理
        i = 0
        Do While i < shtFutai.RowCount
            row = TblFutai.Rows(i)
            row("keijo_kin") = commonLogicEXT.convInsNumStr(shtFutai.Cells(i, 6).Value)
            row("tax_kin") = commonLogicEXT.convInsNumStr(shtFutai.Cells(i, 7).Value)
            row("tekiyo1") = shtFutai.Cells(i, 8).Value
            row("tekiyo2") = shtFutai.Cells(i, 9).Value
            row("event_nm") = shtFutai.Cells(i, 10).Value
            row("content_uchi_nm") = shtFutai.Cells(i, 11).Value
            row("content_cd") = shtFutai.Cells(i, 18).Value
            row("content_uchi_cd") = shtFutai.Cells(i, 19).Value
            ' 2016.01.26 ADD START↓ h.hagiwara
            If String.IsNullOrEmpty(shtFutai.Cells(i, 7).Value) Then
            Else
                bigTaxkin += CLng(shtFutai.Cells(i, 7).Value)
            End If
            ' 2016.01.26 ADD END↑ h.hagiwara
            i = i + 1
        Loop
        dataEXTZ0204.PropDrBillReq("seikyu_title1") = Me.txtTitle1.Text
        dataEXTZ0204.PropDrBillReq("seikyu_title2") = Me.txtTitle2.Text
        dataEXTZ0204.PropDrBillReq("tax_kin") = bigTaxkin                        ' 2016.01.26 ADD h.hagiwara
        dataEXTZ0204.PropBlnChangeFlg = True
        Me.Close()
    End Sub
    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function inputCheckMain() As Boolean
        If String.IsNullOrEmpty(Me.txtTitle1.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0001, "タイトル１"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        If String.IsNullOrEmpty(Me.txtTitle2.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0001, "タイトル２"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        ' 2015.12.11 ADD START↓ h.hagiwara
        '全角チェック
        If CommonValidation.IsFullChar(Me.txtTitle2.Text) = False Then
            puErrMsg = String.Format(CommonDeclareEXT.E0006, "タイトル２")
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return False
        End If
        ' 2015.12.11 ADD END↑ h.hagiwara
        'SPREAD
        Dim shtRiyoryo As FarPoint.Win.Spread.SheetView = Me.fbRiyoryo.ActiveSheet
        Dim shtFutai As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        Dim i As Integer = 0
        '利用料の処理
        Do While i < shtRiyoryo.RowCount
            If String.IsNullOrEmpty(shtRiyoryo.Cells(i, 15).Value) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "プロジェクト"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(shtRiyoryo.Cells(i, 16).Value) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "プロジェクト内訳"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            i = i + 1
        Loop
        '付帯設備の処理
        i = 0
        Do While i < shtFutai.RowCount
            If String.IsNullOrEmpty(shtFutai.Cells(i, 18).Value) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "プロジェクト"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(shtFutai.Cells(i, 19).Value) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "プロジェクト内訳"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            i = i + 1
        Loop
        If Me.lblSagakuKin.Text <> 0 Then
            MsgBox(String.Format(CommonDeclareEXT.E2033, "請求金額との差額"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        ' 2016.01.26 UPD START↓ h.hagiwara 消費税差額を可能とする 
        'If Me.lblTaxSagaku.Text <> 0 Then
        '    MsgBox(String.Format(CommonDeclareEXT.E2033, "消費税差額"), MsgBoxStyle.Exclamation, "エラー")
        '    Return False
        'End If
        If Me.lblTaxSagaku.Text <> 0 Then
            If MsgBox(String.Format(CommonDeclareEXT.C0016, "消費税"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = vbYes Then
            Else
                Return False
            End If
        End If
        ' 2016.01.26 UPD START↓ h.hagiwara 消費税差額を可能とする 

        Return True
    End Function

    ''' <summary>
    ''' 戻るボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
    End Sub

End Class
