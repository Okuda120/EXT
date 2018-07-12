Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win.Spread
Imports FarPoint.Win.Spread.Model

''' <summary>
''' EXTZ0209
''' </summary>
''' <remarks>付帯設備登録／詳細
''' <para>作成情報：2015/09/01 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0209

    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT      '共通ロジッククラス
    Public dataEXTZ0209 As New DataEXTZ0209     'データクラス
    Public logicEXTZ0209 As New LogicEXTZ0209     'データクラス
    Private ppBlnChangeFlg As Boolean

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0209_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblYoyakuNo.Text = dataEXTZ0209.PropStrYoyakuNo
        Me.lblSaiji.Text = dataEXTZ0209.PropStrSaijiNm
        Me.lblRiyobi.Text = dataEXTZ0209.PropStrRiyobiDisp

        Dim index As Integer = 0
        logicEXTZ0209.GetBunrui(dataEXTZ0209)
        Me.cmbBunrui.DataSource = dataEXTZ0209.PropFutaiBunruiTable
        Me.cmbBunrui.DisplayMember = "bunrui_nm"
        Me.cmbBunrui.ValueMember = "bunrui_cd"

        dataEXTZ0209.PropVwFutaiSelecSheet = Me.fbFutaiSelected                     ' 2015.12.01 ADD h.hagiwara

        logicEXTZ0209.GetFutai(dataEXTZ0209)
        selectBunrui(dataEXTZ0209.PropFutaiBunruiTable.Rows(0)("bunrui_cd"))
        setDetailToSpread()
        updateTotal()

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 明細情報をスプレッドに表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setDetailToSpread()
        If dataEXTZ0209.PropFutaiDetailTable Is Nothing Then
            Exit Sub
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        Dim table As DataTable = dataEXTZ0209.PropFutaiDetailTable
        ' 2015.12.21 ADD START↓ h.hagiwara
        Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellKeijyo.MaximumValue = 99999999
        numCellKeijyo.ShowSeparator = True
        numCellKeijyo.DecimalPlaces = 0
        ' 2015.12.21 ADD END↑ h.hagiwara
        sheet.RowCount = table.Rows.Count
        'sheet.ColumnCount = 25        ' 2015.11.13 ADD h.hagiwara
        sheet.ColumnCount = 26         ' 2015.11.13 ADD h.hagiwara
        'カラム設定(コード列の非表示)
        sheet.Columns(9).Visible = False
        sheet.Columns(10).Visible = False
        sheet.Columns(11).Visible = False
        sheet.Columns(12).Visible = False
        sheet.Columns(13).Visible = False
        sheet.Columns(14).Visible = False
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
        sheet.Columns(25).Visible = False         ' 2015/11/13 ADD h.hagwiara
        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Locked = True
            sheet.Cells(index, 4).Locked = True
            sheet.Cells(index, 5).Locked = True
            sheet.Cells(index, 7).Locked = True
            sheet.Cells(index, 0).Value = row("futai_bunrui_nm")
            sheet.Cells(index, 1).Value = row("futai_nm")
            sheet.Cells(index, 2).Value = row("futai_tanka")
            sheet.Cells(index, 3).Value = row("futai_su")
            sheet.Cells(index, 4).Value = row("futai_tani")
            sheet.Cells(index, 5).Value = row("futai_shokei")
            sheet.Cells(index, 6).CellType = numCellKeijyo                      ' 2015.12.21 ADD h.hagiwara
            sheet.Cells(index, 6).Value = row("futai_chosei")
            sheet.Cells(index, 7).Value = row("futai_kin")
            sheet.Cells(index, 8).Value = row("futai_biko")
            sheet.Cells(index, 9).Value = row("futai_bunrui_cd")
            sheet.Cells(index, 10).Value = row("futai_cd")
            sheet.Cells(index, 11).Value = row("notax_flg")
            sheet.Cells(index, 12).Value = row("shukei_grp")
            sheet.Cells(index, 13).Value = row("kamoku_cd")
            sheet.Cells(index, 14).Value = row("saimoku_cd")
            sheet.Cells(index, 15).Value = row("uchi_cd")
            sheet.Cells(index, 16).Value = row("shosai_cd")
            sheet.Cells(index, 17).Value = row("karikamoku_cd")
            sheet.Cells(index, 18).Value = row("kari_saimoku_cd")
            sheet.Cells(index, 19).Value = row("kari_uchi_cd")
            sheet.Cells(index, 20).Value = row("kari_shosai_cd")
            sheet.Cells(index, 21).Value = row("kamoku_nm")
            sheet.Cells(index, 22).Value = row("saimoku_nm")
            sheet.Cells(index, 23).Value = row("uchi_nm")
            sheet.Cells(index, 24).Value = row("shosai_nm")
            sheet.Cells(index, 25).Value = 0                        ' 2015/11/13 ADD h.hagwiara
            index = index + 1
        Loop
        'スクロールバーを必要な場合のみ表示させます
        'Me.fbFutaiSelected.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                        ' 2015.12.21 UPD h.hagiwara
        Me.fbFutaiSelected.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                      ' 2015.12.21 UPD h.hagiwara
        Me.fbFutaiSelected.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 分類マスタのデータセットに対してセレクトを行う
    ''' </summary>
    ''' <param name="strBunruiCd"></param>
    ''' <remarks></remarks>
    Private Sub selectBunrui(ByVal strBunruiCd As String)
        If dataEXTZ0209.PropFutaiMstTable Is Nothing Then
            Exit Sub
        End If
        Dim drFutai() As DataRow
        Dim strBunruiKey As String = "bunrui_cd = '" + strBunruiCd + "'"
        drFutai = dataEXTZ0209.PropFutaiMstTable.Select(strBunruiKey, "sort")
        Dim drBunrui() As DataRow
        drBunrui = dataEXTZ0209.PropFutaiBunruiTable.Select(strBunruiKey, "sort")

        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        sheet.RowCount = drFutai.Count
        sheet.ColumnCount = 20
        'カラム設定(コード列の非表示)
        sheet.Columns(1).Visible = False
        sheet.Columns(2).Visible = False
        sheet.Columns(3).Visible = False
        sheet.Columns(4).Visible = False
        sheet.Columns(5).Visible = False
        sheet.Columns(6).Visible = False
        sheet.Columns(7).Visible = False
        sheet.Columns(8).Visible = False
        sheet.Columns(9).Visible = False
        sheet.Columns(10).Visible = False
        sheet.Columns(11).Visible = False
        sheet.Columns(12).Visible = False
        sheet.Columns(13).Visible = False
        sheet.Columns(14).Visible = False
        sheet.Columns(15).Visible = False
        sheet.Columns(16).Visible = False
        sheet.Columns(17).Visible = False
        sheet.Columns(18).Visible = False
        sheet.Columns(19).Visible = False
        Dim index As Integer = 0
        Dim bunruiRow As DataRow = drBunrui(0)
        For Each row As DataRow In drFutai
            sheet.Cells(index, 0).Value = row("futai_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("futai_cd")
            sheet.Cells(index, 2).Value = row("tanka")
            sheet.Cells(index, 3).Value = row("tani")
            sheet.Cells(index, 4).Value = row("bunrui_nm")
            sheet.Cells(index, 5).Value = row("bunrui_cd")
            sheet.Cells(index, 6).Value = row("notax_flg")
            sheet.Cells(index, 7).Value = row("shukei_grp")
            sheet.Cells(index, 8).Value = row("kamoku_cd")
            sheet.Cells(index, 9).Value = row("saimoku_cd")
            sheet.Cells(index, 10).Value = row("uchi_cd")
            sheet.Cells(index, 11).Value = row("shosai_cd")
            sheet.Cells(index, 12).Value = row("karikamoku_cd")
            sheet.Cells(index, 13).Value = row("kari_saimoku_cd")
            sheet.Cells(index, 14).Value = row("kari_uchi_cd")
            sheet.Cells(index, 15).Value = row("kari_shosai_cd")
            sheet.Cells(index, 16).Value = row("kamoku_nm")
            sheet.Cells(index, 17).Value = row("saimoku_nm")
            sheet.Cells(index, 18).Value = row("uchi_nm")
            sheet.Cells(index, 19).Value = row("shosai_nm")
            index = index + 1
        Next
        'スクロールバーを必要な場合のみ表示させます
        Me.fbFutai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbFutai.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 分類選択時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmbBunrui_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBunrui.SelectedIndexChanged
        selectBunrui(cmbBunrui.SelectedValue.ToString())
    End Sub

    ''' <summary>
    ''' 行コピー
    ''' </summary>
    ''' <param name="activeRow"></param>
    ''' <remarks></remarks>
    Private Sub copyrow(activeRow As Integer)
        Dim lastRow As Integer = fbFutaiSelected.ActiveSheet.RowCount
        ' 2015.12.21 ADD START↓ h.hagiwara
        Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellKeijyo.MaximumValue = 99999999
        numCellKeijyo.ShowSeparator = True
        numCellKeijyo.DecimalPlaces = 0
        ' 2015.12.21 ADD END↑ h.hagiwara

        '最終行に1行追加
        fbFutaiSelected.Sheets(0).AddUnboundRows(lastRow, 1)

        '設定
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 0).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 1).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 2).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 4).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 5).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 7).Locked = True
        '選択された設備情報をコピー
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 0).Value = fbFutai.ActiveSheet.Cells(activeRow, 4).Value 'ジャンル
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 1).Value = fbFutai.ActiveSheet.Cells(activeRow, 0).Value '設備名
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 2).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '単価
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 3).Value = 1                                             '数量
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 4).Value = fbFutai.ActiveSheet.Cells(activeRow, 3).Value '単位
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 5).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '小計→ tani * su
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 6).CellType = numCellKeijyo                              ' 2015.12.21 ADD h.hagiwara
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 6).Value = 0                                             '調整額
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 7).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '合計→ tani * su - chosei
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 8).Value = Nothing
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 9).Value = fbFutai.ActiveSheet.Cells(activeRow, 5).Value '分類CD
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 10).Value = fbFutai.ActiveSheet.Cells(activeRow, 1).Value '付帯CD
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 11).Value = fbFutai.ActiveSheet.Cells(activeRow, 6).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 12).Value = fbFutai.ActiveSheet.Cells(activeRow, 7).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 13).Value = fbFutai.ActiveSheet.Cells(activeRow, 8).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 14).Value = fbFutai.ActiveSheet.Cells(activeRow, 9).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 15).Value = fbFutai.ActiveSheet.Cells(activeRow, 10).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 16).Value = fbFutai.ActiveSheet.Cells(activeRow, 11).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 17).Value = fbFutai.ActiveSheet.Cells(activeRow, 12).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 18).Value = fbFutai.ActiveSheet.Cells(activeRow, 13).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 19).Value = fbFutai.ActiveSheet.Cells(activeRow, 14).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 20).Value = fbFutai.ActiveSheet.Cells(activeRow, 15).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 21).Value = fbFutai.ActiveSheet.Cells(activeRow, 16).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 22).Value = fbFutai.ActiveSheet.Cells(activeRow, 17).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 23).Value = fbFutai.ActiveSheet.Cells(activeRow, 18).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 24).Value = fbFutai.ActiveSheet.Cells(activeRow, 19).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, 25).Value = 0                                           ' 2015/11/13 ADD h.hagiwara
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 業削除
    ''' </summary>
    ''' <param name="activeRow"></param>
    ''' <remarks></remarks>
    Private Sub removerow(activeRow As Integer)
        Dim row As Integer = fbFutaiSelected.Sheets(0).ActiveRowIndex
        If row <> -1 Then
            fbFutaiSelected.Sheets(0).RemoveRows(row, 1)
        End If
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 料金自動計算
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub ChangeFutaiKin(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fbFutaiSelected.Change
        If 3 <> e.Column And 6 <> e.Column Then
            Exit Sub
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        'Dim tanka As Integer                      ' 2015.12.21 UPD  h.hagiwara
        Dim tanka As Long                          ' 2015.12.21 UPD  h.hagiwara
        Dim su As Integer
        ' 2015.12.21 UPD START↓ h.hagiwara
        'Dim shokei As Integer
        'Dim chosei As Integer
        'Dim total As Integer
        Dim shokei As Long
        Dim chosei As Integer
        Dim total As Long
        ' 2015.12.21 UPD END↑ h.hagiwara
        Dim strSu As String = sheet.Cells(e.Row, 3).Value
        Dim strChosei As String = sheet.Cells(e.Row, 6).Value
        If commonValidate.IsHalfNmb(strSu) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0003, "数量"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If commonValidate.IsHalfMNmb(strChosei) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0004, "調整金額"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If String.IsNullOrEmpty(strSu) = True Then
            strSu = 1
            sheet.Cells(e.Row, 3).Value = 1
        End If
        If String.IsNullOrEmpty(strChosei) = True Then
            sheet.Cells(e.Row, 6).Value = 0
            strChosei = 0
        End If
        tanka = sheet.Cells(e.Row, 2).Value
        su = Integer.Parse(strSu)
        shokei = tanka * su
        chosei = Integer.Parse(strChosei)
        total = shokei + chosei

        ' 2015.12.21 UPD START↓ h.hagiwara
        If total.ToString.Length > 11 Then
            MsgBox(String.Format(CommonDeclareEXT.E2049, "付帯設備"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara

        sheet.Cells(e.Row, 5).Value = shokei
        sheet.Cells(e.Row, 7).Value = total

        '合計再計算
        updateTotal()
    End Sub

    ''' <summary>
    ''' ＞ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        copyrow(fbFutai.ActiveSheet.ActiveRowIndex)
    End Sub

    ''' <summary>
    ''' ＞＞ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAddAll_Click(sender As Object, e As EventArgs) Handles btnAddAll.Click
        Dim lastRow As Integer = fbFutai.ActiveSheet.GetLastNonEmptyRow(FarPoint.Win.Spread.NonEmptyItemFlag.Data)
        For i = 0 To lastRow
            copyrow(i)
        Next
    End Sub

    ''' <summary>
    ''' ＜ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        removerow(fbFutai.ActiveSheet.ActiveRowIndex)
    End Sub

    ''' <summary>
    ''' ＜＜ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelAll_Click(sender As Object, e As EventArgs) Handles btnDelAll.Click
        Dim lastRow As Integer = fbFutaiSelected.ActiveSheet.GetLastNonEmptyRow(FarPoint.Win.Spread.NonEmptyItemFlag.Data)
        For i = 0 To lastRow
            removerow(0)
        Next
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 合計の再計算
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub updateTotal()
        fbFutaiSelected.ActiveSheet.ColumnFooter.SetAggregationType(0, 7, FarPoint.Win.Spread.Model.AggregationType.Sum)
        lblTotal.Text = Format(fbFutaiSelected.Sheets(0).ColumnFooter.Cells(0, 7).Value, "#,#")
    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click
        'input check
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        Dim index As New Integer
        Dim newRow As DataRow
        Dim table As DataTable = dataEXTZ0209.PropFutaiDetailTable.clone
        Dim futaiTotalNm As String = ""
        Dim isFirst As Boolean = True

        ' 2015.12.01 ADD START↓ h.hagiwara チェック追加(重複チェックも含む)
        If logicEXTZ0209.InputCheck(dataEXTZ0209) = False Then
            'メッセージを出力 
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.01 ADD END↑ h.hagiwara チェック追加(重複チェックも含む)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 消費税率取得  2015/11/13 ADD h.hagiwara
        'Dim intTaxkin As Integer                                 ' 2015.12.21 UPD h.hagiwara
        'Dim intTotalTaxkin As Integer                            ' 2015.12.21 UPD h.hagiwara
        Dim intTaxkin As Long                                     ' 2015.12.21 UPD h.hagiwara
        Dim intTotalTaxkin As Long                                ' 2015.12.21 UPD h.hagiwara
        Dim dblTaxritu As Double
        dblTaxritu = logicEXTZ0209.GetTax(dataEXTZ0209.PropStrRiyobi)

        'クリア
        table.Clear()
        index = 0
        Do While index < sheet.Rows.Count
            newRow = table.NewRow()
            newRow("yoyaku_dt") = dataEXTZ0209.PropStrRiyobi
            newRow("futai_bunrui_cd") = sheet.Cells(index, 9).Value
            newRow("futai_cd") = sheet.Cells(index, 10).Value
            newRow("futai_tanka") = sheet.Cells(index, 2).Value
            newRow("futai_su") = sheet.Cells(index, 3).Value
            newRow("futai_shokei") = sheet.Cells(index, 5).Value
            newRow("futai_chosei") = sheet.Cells(index, 6).Value
            newRow("futai_kin") = sheet.Cells(index, 7).Value
            newRow("futai_biko") = commonLogicEXT.DbNothingToNull(sheet.Cells(index, 8).Value)
            newRow("futai_nm") = sheet.Cells(index, 1).Value
            newRow("futai_tani") = sheet.Cells(index, 4).Value
            newRow("futai_bunrui_nm") = sheet.Cells(index, 0).Value
            newRow("notax_flg") = sheet.Cells(index, 11).Value
            newRow("shukei_grp") = sheet.Cells(index, 12).Value
            newRow("kamoku_cd") = sheet.Cells(index, 13).Value
            newRow("saimoku_cd") = sheet.Cells(index, 14).Value
            newRow("uchi_cd") = sheet.Cells(index, 15).Value
            newRow("shosai_cd") = sheet.Cells(index, 16).Value
            newRow("karikamoku_cd") = sheet.Cells(index, 17).Value
            newRow("kari_saimoku_cd") = sheet.Cells(index, 18).Value
            newRow("kari_uchi_cd") = sheet.Cells(index, 19).Value
            newRow("kari_shosai_cd") = sheet.Cells(index, 20).Value
            newRow("kamoku_nm") = sheet.Cells(index, 21).Value
            newRow("saimoku_nm") = sheet.Cells(index, 22).Value
            newRow("uchi_nm") = sheet.Cells(index, 23).Value
            newRow("shosai_nm") = sheet.Cells(index, 24).Value
            ' 2015.11.13 ADD START↓ h.hagiwara
            If sheet.Cells(index, 11).Value = "0" Then
                'intTaxkin = Integer.Parse(sheet.Cells(index, 7).Value) * dblTaxritu              ' 2015.12.21 UPD h.hagiwara
                intTaxkin = Long.Parse(sheet.Cells(index, 7).Value) * dblTaxritu              ' 2015.12.21 UPD h.hagiwara
            Else
                intTaxkin = 0
            End If
            newRow("tax_kin") = intTaxkin
            ' 2015.11.13 ADD END↑ h.hagiwara
            table.Rows.Add(newRow)
            If isFirst = False Then
                futaiTotalNm = futaiTotalNm + "／" + sheet.Cells(index, 1).Value
            Else
                futaiTotalNm = sheet.Cells(index, 1).Value
                isFirst = False
            End If
            intTotalTaxkin += intTaxkin                 ' 2015.11.13 ADD h.hagiwara
            index = index + 1
        Loop
        If String.IsNullOrEmpty(Me.lblTotal.Text) = True Then
            dataEXTZ0209.PropIntTotal = 0
            dataEXTZ0209.PropIntTotalTax = 0                                                       ' 2015.11.13 ADD h.hagiwara
        Else
            'dataEXTZ0209.PropIntTotal = Integer.Parse(Me.lblTotal.Text.Replace(",", ""))           ' 2015.12.21 UPD h.hagiwara
            dataEXTZ0209.PropIntTotal = Long.Parse(Me.lblTotal.Text.Replace(",", ""))               ' 2015.12.21 UPD h.hagiwara
            dataEXTZ0209.PropIntTotalTax = intTotalTaxkin                                           ' 2015.11.13 ADD h.hagiwara
        End If
        dataEXTZ0209.PropStrFutaiTotalNm = futaiTotalNm
        dataEXTZ0209.PropFutaiDetailTable = table
        ppBlnChangeFlg = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻る処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        ppBlnChangeFlg = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
        End Set
    End Property

End Class
