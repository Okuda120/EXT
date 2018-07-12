Imports Common
Imports CommonEXT

''' <summary>
''' EXTZ0205
''' </summary>
''' <remarks>プロジェクト一覧
''' <para>作成情報：2015/09/01 k.machida
''' <p>改訂情報:2015/11/30 hayabuchi(チェック制御)</p>
''' </para></remarks>
Public Class EXTZ0205

    Private commonLogic As New CommonLogic              '共通クラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス
    Public dataEXTZ0205 As New DataEXTZ0205             'データクラス
    Public logicEXTZ0205 As New LogicEXTZ0205           'ロジッククラス

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0205_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataEXTZ0205.PropBlnChangeFlg = False

        '画面.スプレッドシート
        dataEXTZ0205.PropResult = Me.fbResult

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

        '2015/12/11 m.hayabuchi add START
        '画面表示
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbResult.ActiveSheet
        dataEXTZ0205.PropStrPrjCd = Me.txtPrjCd.Text
        dataEXTZ0205.PropStrPrjNm = Me.txtPrjNm.Text
        dataEXTZ0205.PropUchiCd = Me.txtUchiCd.Text
        dataEXTZ0205.PropUchiNm = Me.txtUchiNm.Text
        If logicEXTZ0205.InitSearch(dataEXTZ0205) = False Then
            '検索結果0
            sheet.RowCount = 0
        Else
            sheet.RowCount = dataEXTZ0205.PropDtResult.Rows.Count
            sheet.ColumnCount = 5
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < sheet.RowCount
                row = dataEXTZ0205.PropDtResult.Rows(index)
                sheet.Cells(index, 0).Value = False
                sheet.Cells(index, 1).Value = row("event_cd")
                sheet.Cells(index, 2).Value = row("event_nm")
                sheet.Cells(index, 3).Value = row("content_uchi_cd")
                sheet.Cells(index, 4).Value = row("content_uchi_nm")
                index = index + 1
            Loop
        End If
        '2015/12/11 m.hayabuchi add END
    End Sub

    ''' <summary>
    ''' 検索処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbResult.ActiveSheet
        dataEXTZ0205.PropStrPrjCd = Me.txtPrjCd.Text
        dataEXTZ0205.PropStrPrjNm = Me.txtPrjNm.Text
        dataEXTZ0205.PropUchiCd = Me.txtUchiCd.Text
        dataEXTZ0205.PropUchiNm = Me.txtUchiNm.Text
        If logicEXTZ0205.GetProject(dataEXTZ0205) = False Then
            '検索結果0
            sheet.RowCount = 0
        Else
            sheet.RowCount = dataEXTZ0205.PropDtResult.Rows.Count
            sheet.ColumnCount = 5
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < sheet.RowCount
                row = dataEXTZ0205.PropDtResult.Rows(index)
                sheet.Cells(index, 0).Value = False
                sheet.Cells(index, 1).Value = row("event_cd")
                sheet.Cells(index, 2).Value = row("event_nm")
                sheet.Cells(index, 3).Value = row("content_uchi_cd")
                sheet.Cells(index, 4).Value = row("content_uchi_nm")
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
                    MsgBox(String.Format(CommonDeclareEXT.E2035, "プロジェクト"), MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If
                dataEXTZ0205.PropStrResPrjCd = sheet.Cells(index, 1).Value
                dataEXTZ0205.PropStrResPrjNm = sheet.Cells(index, 2).Value
                dataEXTZ0205.PropStrResUchiCd = sheet.Cells(index, 3).Value
                dataEXTZ0205.PropStrResUchiNm = sheet.Cells(index, 4).Value
                selected = True
            End If
            index = index + 1
        Loop

        ' 2015.11.25 ADD h.hagiwara 未選択時のエラー判定
        If selected = True Then
        Else
            MsgBox(String.Format(CommonDeclareEXT.E0002, "プロジェクト"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        dataEXTZ0205.PropBlnChangeFlg = True
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
        dataEXTZ0205.PropIntCheckRow = e.Row

        '選択行にチェックをつけ、それ以外はチェックを外す
        If logicEXTZ0205.ClickResultCellMain(dataEXTZ0205) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        AddHandler fbResult.ButtonClicked, AddressOf fbResult_ButtonClicked

    End Sub

    ''' <summary>
    ''' 戻るボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        dataEXTZ0205.PropBlnChangeFlg = False        ' 2015.11.25 ADD h.hagiwara 重複選択後に戻った際に１件目選択情報が引き継がれる
        Me.Close()
    End Sub

End Class
