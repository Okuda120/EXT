Imports Common
Imports CommonEXT

''' <summary>
''' EXTZ0208
''' </summary>
''' <remarks>EXAS入金情報一覧
''' <para>作成情報：2015/09/02 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0208

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Public dataEXTZ0208 As New DataEXTZ0208             'データクラス
    Public logicEXTZ0208 As New LogicEXTZ0208           'ロジッククラス

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0208_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'seq発行済の場合、入力完了不可
        Dim row As DataRow = dataEXTZ0208.PropDrBillReq
        If dataEXTZ0208.propStrNyukinInputFlg Is Nothing Then
            Me.btnKakutei.Enabled = False
        ElseIf IsDBNull(row("seq")) = False And dataEXTZ0208.propStrNyukinInputFlg = "1" Then
            Me.btnKakutei.Enabled = False
        End If

        '画面.スプレッドシート
        dataEXTZ0208.PropResult = Me.fbResult

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbResult.ActiveSheet
        dataEXTZ0208.PropStrAiteCd = Me.txtAiteCd.Text
        dataEXTZ0208.PropStrAiteNm = Me.txtAiteNm.Text
        dataEXTZ0208.PropStrNyukinYoteiFrom = Me.dtpNyukinYoteiFrom.txtDate.Text
        dataEXTZ0208.PropStrNyukinYoteiTo = Me.dtpNyukinYoteiTo.txtDate.Text
        dataEXTZ0208.PropStrNyukinFrom = Me.dtpNyukinFrom.txtDate.Text
        dataEXTZ0208.PropStrNyukinTo = Me.dtpNyukinTo.txtDate.Text
        dataEXTZ0208.PropStrSeikyuNo = Me.txtSeikyuNo.Text
        dataEXTZ0208.PropStrSeikyuIraiNo = Me.txtSeikyuIraiNo.Text

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If logicEXTZ0208.GetEXASNyukin(dataEXTZ0208) = False Then
            '検索結果0
            sheet.RowCount = 0
        Else
            sheet.RowCount = dataEXTZ0208.PropDtResult.Rows.Count
            sheet.ColumnCount = 17
            sheet.Columns(16).Visible = False
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < sheet.RowCount
                row = dataEXTZ0208.PropDtResult.Rows(index)
                sheet.Cells(index, 0).Value = False
                sheet.Cells(index, 1).Value = commonLogicEXT.DbNullToNothing(row, "aite_cd")
                sheet.Cells(index, 2).Value = commonLogicEXT.DbNullToNothing(row, "yoyaku_start_dt")
                sheet.Cells(index, 3).Value = commonLogicEXT.DbNullToNothing(row, "yoyaku_end_dt")
                sheet.Cells(index, 4).Value = commonLogicEXT.DbNullToNothing(row, "aite_nm")
                sheet.Cells(index, 5).Value = commonLogicEXT.DbNullToNothing(row, "yoyaku_no")
                sheet.Cells(index, 6).Value = commonLogicEXT.DbNullToNothing(row, "saiji_nm")
                sheet.Cells(index, 7).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_yotei_dt")
                sheet.Cells(index, 8).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_dt")
                sheet.Cells(index, 9).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_kin")
                sheet.Cells(index, 10).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_input_flg")
                sheet.Cells(index, 11).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_dt")
                sheet.Cells(index, 12).Value = commonLogicEXT.DbNullToNothing(row, "input_dt")
                sheet.Cells(index, 13).Value = commonLogicEXT.DbNullToNothing(row, "sekikyu_no")
                sheet.Cells(index, 14).Value = commonLogicEXT.DbNullToNothing(row, "seikyu_irai_no")
                sheet.Cells(index, 15).Value = commonLogicEXT.DbNullToNothing(row, "monitor_no")
                sheet.Cells(index, 16).Value = commonLogicEXT.DbNullToNothing(row, "nyukin_link_no")
                If dataEXTZ0208.propStrNyukinInputFlg Is Nothing Then
                    sheet.Cells(index, 0).Locked = True
                End If
                sheet.Cells(index, 1).Locked = True
                sheet.Cells(index, 2).Locked = True
                sheet.Cells(index, 3).Locked = True
                sheet.Cells(index, 4).Locked = True
                sheet.Cells(index, 5).Locked = True
                sheet.Cells(index, 6).Locked = True
                sheet.Cells(index, 7).Locked = True
                sheet.Cells(index, 8).Locked = True
                sheet.Cells(index, 9).Locked = True
                sheet.Cells(index, 10).Locked = True
                sheet.Cells(index, 11).Locked = True
                sheet.Cells(index, 12).Locked = True
                sheet.Cells(index, 13).Locked = True
                sheet.Cells(index, 14).Locked = True
                sheet.Cells(index, 15).Locked = True
                index = index + 1
            Loop
        End If
        'スクロールバーを必要な場合のみ表示させます
        fbResult.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        fbResult.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 確定ボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnKakutei_Click(sender As Object, e As EventArgs) Handles btnKakutei.Click
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbResult.ActiveSheet
        Dim selected As Boolean = False
        Dim index As New Integer
        index = 0
        Do While index < sheet.Rows.Count
            If sheet.Cells(index, 0).Value = True Then
                If selected = True Then
                    MsgBox(String.Format(CommonDeclareEXT.E2035, "入金情報"), MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If
                dataEXTZ0208.PropStrResAiteCd = sheet.Cells(index, 1).Value
                dataEXTZ0208.PropStrResAiteNm = sheet.Cells(index, 4).Value
                dataEXTZ0208.PropStrResNyukinYoteiDt = sheet.Cells(index, 7).Value
                dataEXTZ0208.PropStrResNyukinDt = sheet.Cells(index, 8).Value
                dataEXTZ0208.PropIntResNyukinKin = sheet.Cells(index, 9).Value
                dataEXTZ0208.PropStrResSeikyuDt = sheet.Cells(index, 11).Value
                dataEXTZ0208.PropStrResInputDt = sheet.Cells(index, 12).Value
                dataEXTZ0208.PropStrResSeikyuNo = sheet.Cells(index, 13).Value
                dataEXTZ0208.PropStrResSeikyuIraiNo = sheet.Cells(index, 14).Value
                dataEXTZ0208.PropIntResNyukinLink = sheet.Cells(index, 16).Value
                selected = True
                '入力チェック   ' 2016.08.16  m.hayabuchi add
                If logicEXTZ0208.InputCheck(dataEXTZ0208) = False Then
                    'メッセージを出力 
                    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If
            End If
            index = index + 1
        Loop
        If selected = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "入金情報"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        dataEXTZ0208.PropBlnChangeFlg = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻るボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        dataEXTZ0208.PropBlnChangeFlg = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' スプレッドシートクリック時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>スプレッドシートのセルをクリックした際の処理(単一選択時の疑似ラジオボックス処理）
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub fbResult_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles fbResult.ButtonClicked

        RemoveHandler fbResult.ButtonClicked, AddressOf fbResult_ButtonClicked

        '選択行番号取得
        dataEXTZ0208.PropIntCheckRow = e.Row

        '選択行にチェックをつけ、それ以外はチェックを外す
        If logicEXTZ0208.ClickResultCellMain(dataEXTZ0208) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        AddHandler fbResult.ButtonClicked, AddressOf fbResult_ButtonClicked

    End Sub
End Class
