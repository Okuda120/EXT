Imports CommonEXT
Imports Common

Public Class EXTZ0203

    Public dataEXTZ0203 As New DataEXTZ0203     'データクラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Private commonLogic As New CommonLogic
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    'Public logicEXTZ0212 As New LogicEXTZ0212     'データクラス
    Private LogicEXTZ0203 As New LogicEXTZ0203      ' ロジック

    Private Const PLUS As String = "＋"
    Private Const MINUS As String = "－"
    Private Const PLMI As String = "±"

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0203_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblSeikyuIraiNo.Text = dataEXTZ0203.propStrSeikyuIraiNo
        Me.dtpSeikyuDt.txtDate.Text = dataEXTZ0203.propStrSeikyuDt
        Me.dtpNyukinYoteiDt.txtDate.Text = dataEXTZ0203.propStrNyukinYoteiDt
        '  2015.12.21 UPD START↓ h.hagiwara
        'Me.lblKakuteiKin.Text = Integer.Parse(dataEXTZ0203.propStrKakuteiKin).ToString("#,##0")
        'Me.lblShokei.Text = Integer.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")      
        'Me.lblTaxKin.Text = Integer.Parse(dataEXTZ0203.propStrTaxKin).ToString("#,##0")               
        'Me.lblSeikyuKin.Text = Integer.Parse(dataEXTZ0203.propStrSeikyuKin).ToString("#,##0")         
        Me.lblKakuteiKin.Text = Long.Parse(dataEXTZ0203.propStrKakuteiKin).ToString("#,##0")
        Me.lblShokei.Text = Long.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        Me.lblTaxKin.Text = Long.Parse(dataEXTZ0203.propStrTaxKin).ToString("#,##0")
        Me.lblSeikyuKin.Text = Long.Parse(dataEXTZ0203.propStrSeikyuKin).ToString("#,##0")
        '  2015.12.21 UPD END↑ h.hagiwara
        Me.cmbFugo.Items.Add(PLMI)
        Me.cmbFugo.Items.Add(PLUS)
        Me.cmbFugo.Items.Add(MINUS)
        ' 2015.12.21 UPD START↓ h.hagiwara  
        'If String.IsNullOrEmpty(dataEXTZ0203.propStrChoseiKin) Or "0" = dataEXTZ0203.propStrChoseiKin Then
        '    Me.cmbFugo.Text = PLMI
        '    Me.txtChoseiKin.Text = dataEXTZ0203.propStrChoseiKin
        'ElseIf 0 < Integer.Parse(dataEXTZ0203.propStrChoseiKin) Then
        '    Me.cmbFugo.Text = PLUS
        '    Me.txtChoseiKin.Text = dataEXTZ0203.propStrChoseiKin
        'ElseIf 0 > Integer.Parse(dataEXTZ0203.propStrChoseiKin) Then
        '    Me.cmbFugo.Text = MINUS
        '    Dim i = Integer.Parse(dataEXTZ0203.propStrChoseiKin)
        '    i = i * -1
        '    Me.txtChoseiKin.Text = i.ToString
        'End If
        If String.IsNullOrEmpty(dataEXTZ0203.propStrChoseiKin) Or "0" = dataEXTZ0203.propStrChoseiKin Then
            Me.cmbFugo.Text = PLMI
            Me.txtChoseiKin.Text = dataEXTZ0203.propStrChoseiKin
        ElseIf 0 < Long.Parse(dataEXTZ0203.propStrChoseiKin) Then
            Me.cmbFugo.Text = PLUS
            Me.txtChoseiKin.Text = dataEXTZ0203.propStrChoseiKin
        ElseIf 0 > Long.Parse(dataEXTZ0203.propStrChoseiKin) Then
            Me.cmbFugo.Text = MINUS
            Dim i = Long.Parse(dataEXTZ0203.propStrChoseiKin)
            i = i * -1
            Me.txtChoseiKin.Text = i.ToString
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara  
        Me.lblSeikyuNaiyo.Text = dataEXTZ0203.propStrSeikyuNaiyoNm
        Me.lblAiteCd.Text = dataEXTZ0203.propStrAiteCd
        Me.lblAiteNm.Text = dataEXTZ0203.propStrAiteNm
        If dataEXTZ0203.propStrNyukinKbn = "1" Then
            Me.rdoNyukin1.Checked = True
        ElseIf dataEXTZ0203.propStrNyukinKbn = "2" Then
            Me.rdoNyukin2.Checked = True
        End If
        If dataEXTZ0203.propStrSeikyuInputFlg = "0" Then
            Me.lblSts.Text = "未入力"
        ElseIf dataEXTZ0203.propStrSeikyuInputFlg = "1" Then
            Me.lblSts.Text = "入力済"
        End If
        'seq発行済の場合、入力完了不可
        Dim row As DataRow = dataEXTZ0203.PropDsBillReq.Tables("BILLPAY_TBL").Rows(dataEXTZ0203.propIntEditLine)
        'If IsDBNull(row("seq")) = False And dataEXTZ0203.propStrSeikyuInputFlg = "1" Then
        If IsDBNull(row("seikyu_irai_no")) = False And dataEXTZ0203.propStrSeikyuInputFlg = "1" Then
            Me.btnComplate.Enabled = False
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 調整金額変更時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtChoseiKin_TextChanged(sender As Object, e As EventArgs) Handles txtChoseiKin.TextChanged, cmbFugo.SelectedIndexChanged
        calc()
    End Sub

    ''' <summary>
    ''' 計算処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub calc()
        If String.IsNullOrEmpty(Me.txtChoseiKin.Text) Or String.IsNullOrEmpty(Me.cmbFugo.Text) Or PLMI = Me.cmbFugo.Text Then
            Exit Sub
        End If
        dataEXTZ0203.propStrChoseiKin = Me.txtChoseiKin.Text
        ' 2015.12.21 UPD START↓ h.hagiwara
        'If Me.cmbFugo.Text = PLUS Then
        '    dataEXTZ0203.propStrShokei = (Integer.Parse(dataEXTZ0203.propStrKakuteiKin) + Integer.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
        '    Me.lblShokei.Text = Integer.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        'ElseIf Me.cmbFugo.Text = MINUS Then
        '    dataEXTZ0203.propStrShokei = (Integer.Parse(dataEXTZ0203.propStrKakuteiKin) - Integer.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
        '    Me.lblShokei.Text = Integer.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        'End If
        'dataEXTZ0203.propStrTaxKin = Integer.Parse(Math.Round(Integer.Parse(dataEXTZ0203.propStrShokei) * dataEXTZ0203.propDblTax))
        'dataEXTZ0203.propStrSeikyuKin = Integer.Parse(dataEXTZ0203.propStrShokei) + Integer.Parse(dataEXTZ0203.propStrTaxKin)
        'Me.lblTaxKin.Text = Integer.Parse(dataEXTZ0203.propStrTaxKin).ToString("#,##0")
        'Me.lblSeikyuKin.Text = Integer.Parse(dataEXTZ0203.propStrSeikyuKin).ToString("#,##0")
        If Me.cmbFugo.Text = PLUS Then
            dataEXTZ0203.propStrShokei = (Long.Parse(dataEXTZ0203.propStrKakuteiKin) + Long.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
            Me.lblShokei.Text = Long.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        ElseIf Me.cmbFugo.Text = MINUS Then
            dataEXTZ0203.propStrShokei = (Long.Parse(dataEXTZ0203.propStrKakuteiKin) - Long.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
            Me.lblShokei.Text = Long.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        End If
        ' 2016.01.25 UPD START↓ h.hagiwara
        'dataEXTZ0203.propStrTaxKin = Long.Parse(Math.Round(Long.Parse(dataEXTZ0203.propStrShokei) * dataEXTZ0203.propDblTax))
        If dataEXTZ0203.propStrSeikyuInputFlg = "0" Then
            dataEXTZ0203.propStrTaxKin = Long.Parse(Math.Round(Long.Parse(dataEXTZ0203.propStrShokei) * dataEXTZ0203.propDblTax))
        End If
        ' 2016.01.25 UPD END↑ h.hagiwara
        dataEXTZ0203.propStrSeikyuKin = Long.Parse(dataEXTZ0203.propStrShokei) + Long.Parse(dataEXTZ0203.propStrTaxKin)
        Me.lblTaxKin.Text = Long.Parse(dataEXTZ0203.propStrTaxKin).ToString("#,##0")
        Me.lblSeikyuKin.Text = Long.Parse(dataEXTZ0203.propStrSeikyuKin).ToString("#,##0")
        ' 2015.12.21 UPD START↓ h.hagiwara
    End Sub

    ''' <summary>
    ''' 計算処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub calc2()
        dataEXTZ0203.propStrChoseiKin = Me.txtChoseiKin.Text
        If Me.cmbFugo.Text = PLUS Then
            dataEXTZ0203.propStrShokei = (Long.Parse(dataEXTZ0203.propStrKakuteiKin) + Long.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
            Me.lblShokei.Text = Long.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        ElseIf Me.cmbFugo.Text = MINUS Then
            dataEXTZ0203.propStrShokei = (Long.Parse(dataEXTZ0203.propStrKakuteiKin) - Long.Parse(dataEXTZ0203.propStrChoseiKin)).ToString
            Me.lblShokei.Text = Long.Parse(dataEXTZ0203.propStrShokei).ToString("#,##0")
        End If
        dataEXTZ0203.propStrSeikyuKin = Long.Parse(dataEXTZ0203.propStrShokei) + Long.Parse(dataEXTZ0203.propStrTaxKin)
        Me.lblTaxKin.Text = Long.Parse(dataEXTZ0203.propStrTaxKin).ToString("#,##0")
        Me.lblSeikyuKin.Text = Long.Parse(dataEXTZ0203.propStrSeikyuKin).ToString("#,##0")
    End Sub

    ''' <summary>
    ''' EXAS一覧へボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExasPrj_Click(sender As Object, e As EventArgs) Handles btnExasPrj.Click
        Dim frm As New EXTZ0204
        Dim tbl As DataTable = dataEXTZ0203.PropDsBillReq.Tables("BILLPAY_TBL")
        Dim tmpTbl As DataTable = tbl.Clone
        Dim row As DataRow = tmpTbl.NewRow

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        row("seikyu_irai_no") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuIraiNo)
        row("seikyu_dt") = Me.dtpSeikyuDt.txtDate.Text
        row("nyukin_yotei_dt") = Me.dtpNyukinYoteiDt.txtDate.Text
        If Me.rdoNyukin1.Checked = True Then
            row("nyukin_kbn") = "1"
        Else
            row("nyukin_kbn") = "2"
        End If
        row("kakutei_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrKakuteiKin)
        If Me.cmbFugo.Text = PLUS Then
            row("chosei_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrChoseiKin)
        ElseIf Me.cmbFugo.Text = MINUS Then
            'row("chosei_kin") = commonLogicEXT.DbNothingToNull(Integer.Parse(dataEXTZ0203.propStrChoseiKin) * -1)                      ' 2015.12.21 UPD h.hagiwara
            row("chosei_kin") = commonLogicEXT.DbNothingToNull(Long.Parse(dataEXTZ0203.propStrChoseiKin) * -1)                          ' 2015.12.21 UPD h.hagiwara
        End If
        row("shokei") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrShokei)
        row("tax_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrTaxKin)
        row("seikyu_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuKin)
        row("seikyu_naiyo") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuNaiyo)
        row("seikyu_naiyo_nm") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuNaiyoNm)
        row("aite_cd") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrAiteCd)
        row("seikyu_title1") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuTitle1)
        row("seikyu_title2") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuTitle2)
        row("nyukin_dt") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinDt)
        row("nyukin_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinKin)
        row("seikyu_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuInputFlg)
        row("seikyu_irai_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuIraiFlg)
        row("nyukin_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinInputFlg)
        row("nyukin_link_no") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinLinkNo)
        row("get_seikyu_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrGetSeikyuInputFlg)                          ' 2016.02.03 ADD h.hagiwara
        frm.dataEXTZ0204.PropDrBillReq = row
        frm.dataEXTZ0204.PropDblTax = dataEXTZ0203.propDblTax
        frm.dataEXTZ0204.PropDtExasRiyoryo = dataEXTZ0203.PropDtExasRiyoryo
        frm.dataEXTZ0204.PropDtExasFutai = dataEXTZ0203.PropDtExasFutai
        frm.dataEXTZ0204.propStrSeikyuInputFlg = dataEXTZ0203.propStrSeikyuInputFlg
        frm.seikyubi = Me.dtpSeikyuDt.txtDate.Text
        frm.ShowDialog()
        If frm.dataEXTZ0204.PropBlnChangeFlg = True Then
            'プロジェクト設定済
            dataEXTZ0203.PropDtExasRiyoryo = frm.dataEXTZ0204.PropDtExasRiyoryo
            dataEXTZ0203.PropDtExasFutai = frm.dataEXTZ0204.PropDtExasFutai
            dataEXTZ0203.propStrSeikyuTitle1 = frm.dataEXTZ0204.PropDrBillReq("seikyu_title1")
            dataEXTZ0203.propStrSeikyuTitle2 = frm.dataEXTZ0204.PropDrBillReq("seikyu_title2")
            dataEXTZ0203.propStrSeikyuInputFlg = "1"
            dataEXTZ0203.propStrTaxKin = frm.dataEXTZ0204.PropDrBillReq("tax_kin")                                ' 2016.01.26 ADD h.hagiwara
            calc2()                                                                                               ' 2016.01.26 ADD h.hagiwara  
            Me.lblTaxKin.Text = Long.Parse(frm.dataEXTZ0204.PropDrBillReq("tax_kin")).ToString("#,##0")           ' 2016.01.26 ADD h.hagiwara
            Me.lblSts.Text = "入力済"
        End If
    End Sub

    ''' <summary>
    ''' 完了ボタン
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
        Dim row As DataRow = dataEXTZ0203.PropDsBillReq.Tables("BILLPAY_TBL").Rows(dataEXTZ0203.propIntEditLine)
        row("seikyu_irai_no") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuIraiNo)
        row("seikyu_dt") = Me.dtpSeikyuDt.txtDate.Text
        row("nyukin_yotei_dt") = Me.dtpNyukinYoteiDt.txtDate.Text
        If Me.rdoNyukin1.Checked = True Then
            row("nyukin_kbn") = "1"
        Else
            row("nyukin_kbn") = "2"
        End If
        row("kakutei_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrKakuteiKin)
        If Me.cmbFugo.Text = PLUS Then
            row("chosei_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrChoseiKin)
        ElseIf Me.cmbFugo.Text = MINUS Then
            'row("chosei_kin") = commonLogicEXT.DbNothingToNull(Integer.Parse(dataEXTZ0203.propStrChoseiKin) * -1)            ' 2015.12.21 UPD h.hagiwara
            row("chosei_kin") = commonLogicEXT.DbNothingToNull(Long.Parse(dataEXTZ0203.propStrChoseiKin) * -1)                ' 2015.12.21 UPD h.hagiwara
        End If
        row("shokei") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrShokei)
        row("tax_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrTaxKin)
        row("seikyu_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuKin)
        row("seikyu_naiyo") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuNaiyo)
        row("seikyu_naiyo_nm") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuNaiyoNm)
        row("aite_cd") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrAiteCd)
        row("seikyu_title1") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuTitle1)
        row("seikyu_title2") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuTitle2)
        row("nyukin_dt") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinDt)
        row("nyukin_kin") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinKin)
        row("seikyu_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuInputFlg)
        row("seikyu_irai_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrSeikyuIraiFlg)
        row("nyukin_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinInputFlg)
        row("nyukin_link_no") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrNyukinLinkNo)
        row("get_seikyu_input_flg") = commonLogicEXT.DbNothingToNull(dataEXTZ0203.propStrGetSeikyuInputFlg)                          ' 2016.02.03 ADD h.hagiwara
        dataEXTZ0203.propStrSeikyuDt = row("seikyu_dt")
        dataEXTZ0203.propStrNyukinDt = row("nyukin_yotei_dt")
        dataEXTZ0203.propStrNyukinKbn = row("nyukin_kbn")
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
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Chk_START", Nothing, Nothing)
        Try
            If String.IsNullOrEmpty(Me.dtpSeikyuDt.txtDate.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "請求日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If String.IsNullOrEmpty(Me.dtpNyukinYoteiDt.txtDate.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "入金予定日"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            ' 入金予定日の土日祝祭日チェック
            If LogicEXTZ0203.WeekDayCheck(Me.dtpNyukinYoteiDt.txtDate.Text) = False Then
                MsgBox(CommonDeclareEXT.E2039, MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If commonValidate.IsHalfNmb(Me.txtChoseiKin.Text) = False Then
                MsgBox(String.Format(CommonDeclareEXT.E0003, "調整"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            If Me.rdoNyukin1.Checked = True And String.IsNullOrEmpty(dataEXTZ0203.propStrAiteCd) Then
                MsgBox(CommonDeclareEXT.E2022, MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            ' 2015.12.01 UPD START↓ h.hagiwara 入金種別：現金(ALSOK)の場合、請求用情報未設定でもエラーとしない
            'If dataEXTZ0203.propStrSeikyuInputFlg = "0" Then
            '    MsgBox(CommonDeclareEXT.E2024, MsgBoxStyle.Exclamation, "エラー")
            '    Return False
            'End If
            If Me.rdoNyukin1.Checked = True And dataEXTZ0203.propStrSeikyuInputFlg = "0" Then
                MsgBox(CommonDeclareEXT.E2024, MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
            ' 2015.12.01 UPD END↑ h.hagiwara
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
    ''' 戻るボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        '画面を閉じる
        Me.Close()
    End Sub

End Class
