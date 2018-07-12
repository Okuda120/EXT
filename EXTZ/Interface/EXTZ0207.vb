Imports CommonEXT
Imports Common

Public Class EXTZ0207

    Public dataEXTZ0207 As New DataEXTZ0207     'データクラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogic As New CommonLogic
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    'Public logicEXTZ0207 As New LogicEXTZ0207

    Private Sub EXTB0103_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblSeikyuNo.Text = dataEXTZ0207.propStrSeikyuIraiNo
        Me.lblSeikyuDt.Text = dataEXTZ0207.propStrSeikyuDt
        Me.lblNyukinYoteiDt.Text = dataEXTZ0207.propStrNyukinYoteiDt
        'Me.lblKakuteiKin.Text = Integer.Parse(dataEXTZ0207.propStrKakuteiKin).ToString("#,##0")                   ' 2015.12.21 UPD h.hagiwara
        'Me.lblShokei.Text = Integer.Parse(dataEXTZ0207.propStrShokei).ToString("#,##0")                           ' 2015.12.21 UPD h.hagiwara
        Me.lblKakuteiKin.Text = Long.Parse(dataEXTZ0207.propStrKakuteiKin).ToString("#,##0")                       ' 2015.12.21 UPD h.hagiwara
        Me.lblShokei.Text = Long.Parse(dataEXTZ0207.propStrShokei).ToString("#,##0")                               ' 2015.12.21 UPD h.hagiwara
        'Me.lblTaxKin.Text = Integer.Parse(dataEXTZ0207.propStrTaxKin).ToString("#,##0")                           ' 2015.12.21 UPD h.hagiwara
        Me.lblTaxKin.Text = Long.Parse(dataEXTZ0207.propStrTaxKin).ToString("#,##0")                               ' 2015.12.21 UPD h.hagiwara
        'Me.lblSeikyuKin.Text = Integer.Parse(dataEXTZ0207.propStrSeikyuKin).ToString("#,##0")                     ' 2015.12.21 UPD h.hagiwara
        Me.lblSeikyuKin.Text = Long.Parse(dataEXTZ0207.propStrSeikyuKin).ToString("#,##0")                         ' 2015.12.21 UPD h.hagiwara
        Me.lblSeikyuNaiyo.Text = dataEXTZ0207.propStrSeikyuNaiyoNm
        Me.lblAiteCd.Text = dataEXTZ0207.propStrAiteCd
        Me.lblAiteNm.Text = dataEXTZ0207.propStrAiteNm
        Me.dtpNyukinDt.txtDate.Text = dataEXTZ0207.propStrNyukinDt
        Me.txtNyukin.Text = dataEXTZ0207.propStrNyukinKin
        If dataEXTZ0207.propStrNyukinKbn = "1" Then
            Me.rdoNyukin1.Checked = True
        ElseIf dataEXTZ0207.propStrNyukinKbn = "2" Then
            Me.rdoNyukin2.Checked = True
        End If
        SetNyukinSpread(dataEXTZ0207.PropRowNyukin)
        'seq発行済の場合、入力完了不可
        Dim row As DataRow = dataEXTZ0207.PropDsBillReq.Tables("BILLPAY_TBL").Rows(dataEXTZ0207.propIntEditLine)
        If IsDBNull(row("seq")) = False And dataEXTZ0207.propStrNyukinInputFlg = "1" Then
            Me.btnComplate.Enabled = False
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub


    Public Sub SetNyukinSpread(ByRef row As DataRow)
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbNyukin.ActiveSheet
        sheet.RowCount = 1
        sheet.ColumnCount = 11
        sheet.Columns(10).Visible = False
        Try
            sheet.Cells(0, 0).Value = commonLogicEXT.DbNullToNothing(row, "aite_cd")
            sheet.Cells(0, 1).Value = commonLogicEXT.DbNullToNothing(row, "aite_nm")
            sheet.Cells(0, 2).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_yotei_dt")
            sheet.Cells(0, 3).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_dt")
            sheet.Cells(0, 4).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_kin")
            sheet.Cells(0, 5).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_dt")
            sheet.Cells(0, 6).Value = commonLogicEXT.DbNullToNothing(row, "input_dt")
            sheet.Cells(0, 7).Value = commonLogicEXT.DbNullToNothing(row, "sekikyu_no")
            sheet.Cells(0, 8).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_no")
            sheet.Cells(0, 10).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_link_no")
            sheet.Cells(0, 0).Locked = True
            sheet.Cells(0, 1).Locked = True
            sheet.Cells(0, 2).Locked = True
            sheet.Cells(0, 3).Locked = True
            sheet.Cells(0, 4).Locked = True
            sheet.Cells(0, 5).Locked = True
            sheet.Cells(0, 6).Locked = True
            sheet.Cells(0, 7).Locked = True
            sheet.Cells(0, 8).Locked = True
            Me.lblNyukinTtl.Text = commonLogicEXT.convKingaku(sheet.Cells(0, 4).Value)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnExasList_Click(sender As Object, e As EventArgs) Handles btnExasList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0208
        frm.dataEXTZ0208.propStrNyukinInputFlg = dataEXTZ0207.propStrNyukinInputFlg
        frm.dataEXTZ0208.PropDrBillReq = dataEXTZ0207.PropDsBillReq.Tables("BILLPAY_TBL").Rows(dataEXTZ0207.propIntEditLine)
        frm.ShowDialog()
        If frm.dataEXTZ0208.PropBlnChangeFlg Then
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbNyukin.ActiveSheet
            sheet.Cells(0, 0).Value = frm.dataEXTZ0208.PropStrResAiteCd
            sheet.Cells(0, 1).Value = frm.dataEXTZ0208.PropStrResAiteNm
            sheet.Cells(0, 2).Value = frm.dataEXTZ0208.PropStrResNyukinYoteiDt
            sheet.Cells(0, 3).Value = frm.dataEXTZ0208.PropStrResNyukinDt
            sheet.Cells(0, 4).Value = frm.dataEXTZ0208.PropIntResNyukinKin
            sheet.Cells(0, 5).Value = frm.dataEXTZ0208.PropStrResSeikyuDt
            sheet.Cells(0, 6).Value = frm.dataEXTZ0208.PropStrResInputDt
            sheet.Cells(0, 7).Value = frm.dataEXTZ0208.PropStrResSeikyuNo
            sheet.Cells(0, 8).Value = frm.dataEXTZ0208.PropStrResSeikyuIraiNo
            sheet.Cells(0, 10).Value = frm.dataEXTZ0208.PropIntResNyukinLink
            Me.lblNyukinTtl.Text = commonLogicEXT.convKingaku(sheet.Cells(0, 4).Value)
        End If
    End Sub

    Private Sub btnSetNyukinTtl_Click(sender As Object, e As EventArgs) Handles btnSetNyukinTtl.Click
        'Me.txtNyukin.Text = commonLogicEXT.convInteger(Me.lblNyukinTtl.Text)                         ' 2015.12.21 UPD h.hagiwara
        Me.txtNyukin.Text = commonLogicEXT.convLong(Me.lblNyukinTtl.Text)                             ' 2015.12.21 UPD h.hagiwara
    End Sub


    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        dataEXTZ0207.PropBlnChangeFlg = False
        Me.Close()
    End Sub

    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If inputCheckMain() = False Then
            Return
        End If
        dataEXTZ0207.propStrNyukinDt = Me.dtpNyukinDt.txtDate.Text
        dataEXTZ0207.propStrNyukinKin = Me.txtNyukin.Text
        If Me.rdoNyukin1.Checked = True Then
            dataEXTZ0207.propStrNyukinKbn = "1"
        ElseIf Me.rdoNyukin2.Checked = True Then
            dataEXTZ0207.propStrNyukinKbn = "2"
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbNyukin.ActiveSheet
        Dim row As DataRow = dataEXTZ0207.PropRowNyukin
        row("aite_cd") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 0).Value)
        row("aite_nm") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 1).Value)
        row("nyukin_yotei_dt") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 2).Value)
        row("nyukin_dt") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 3).Value)
        row("seikyu_kin") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 4).Value)
        row("seikyu_dt") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 5).Value)
        row("input_dt") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 6).Value)
        row("sekikyu_no") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 7).Value)
        row("seikyu_irai_no") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 8).Value)
        row("nyukin_link_no") = commonLogicEXT.DbNothingToNull(sheet.Cells(0, 10).Value)
        dataEXTZ0207.propStrNyukinLinkNo = commonLogicEXT.DbNullToNothing(row, "nyukin_link_no")
        dataEXTZ0207.propStrNyukinInputFlg = "1"
        dataEXTZ0207.PropBlnChangeFlg = True
        Me.Close()
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
            '必須チェック
            If String.IsNullOrEmpty(Me.dtpNyukinDt.txtDate.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "入金日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(Me.txtNyukin.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "入金額"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            '入金額
            ' 2015.12.21 UPD START↓ h.hagiwara
            'If commonLogicEXT.convInteger(Me.txtNyukin.Text) <> commonLogicEXT.convInteger(Me.lblNyukinTtl.Text) Then
            '    If MsgBox(String.Format(CommonEXT.C0012, "入金額", "EXAS入金データリンク入金合計額（A）"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            '        Return False
            '    End If
            'End If
            If commonLogicEXT.convLong(Me.txtNyukin.Text) <> commonLogicEXT.convLong(Me.lblNyukinTtl.Text) Then
                If MsgBox(String.Format(CommonEXT.C0012, "入金額", "EXAS入金データリンク入金合計額（A）"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
                    Return False
                End If
            End If
            ' 2015.12.21 UPD END↑ h.hagiwara
            If String.IsNullOrEmpty(Me.lblNyukinTtl.Text) Then
                If MsgBox(String.Format(CommonEXT.C0013), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
                    Return False
                End If
            End If
        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, ex.ToString, Nothing, Nothing)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' 入金一覧ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub fbNyukin_ButtonClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbNyukin.ButtonClicked
        If e.Column = 9 Then
            Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbNyukin.ActiveSheet
            sheet.Cells(0, 0).Value = Nothing
            sheet.Cells(0, 1).Value = Nothing
            sheet.Cells(0, 2).Value = Nothing
            sheet.Cells(0, 3).Value = Nothing
            sheet.Cells(0, 4).Value = Nothing
            sheet.Cells(0, 5).Value = Nothing
            sheet.Cells(0, 6).Value = Nothing
            sheet.Cells(0, 7).Value = Nothing
            sheet.Cells(0, 8).Value = Nothing
            sheet.Cells(0, 10).Value = Nothing
        End If
    End Sub

End Class
