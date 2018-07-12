Imports Common
Imports CommonEXT
Imports FarPoint.Win.Spread

Public Class EXTZ0212

    Public dataEXTZ0212 As New DataEXTZ0212        'データクラス
    Public logicEXTZ0212 As New LogicEXTZ0212      'データクラス
    Private commonLogicEXT As New CommonLogicEXT   '共通ロジッククラス


    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0212_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataEXTZ0212.PropBlnChangeFlg = False

        'データセット
        With dataEXTZ0212
            .PropRyokin = Me.fbRyokin            '画面.料金一覧シート
            .PropBairitu = Me.fbBairitu          '画面.倍率一覧シート
        End With

        'DS初期化
        Dim dsTankaBairitu As New DataSet
        dsTankaBairitu.Clear()
        dataEXTZ0212.PropDsTanakaBairitu = dsTankaBairitu
        '分類取得
        If logicEXTZ0212.GetBunrui(dataEXTZ0212) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '料金取得(初期値)
        Dim dsTemp = dataEXTZ0212.PropDsTanakaBairitu
        Dim rowTemp As DataRow = dsTemp.tables("RBUNRUI_MST").Rows(0)
        If logicEXTZ0212.GetRyokin(dataEXTZ0212, rowTemp("bunrui_cd")) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        '倍率取得
        If logicEXTZ0212.GetBairitu(dataEXTZ0212) = False Then
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        'SPREAD設定処理
        SetSpreadBunrui()
        SetSpreadRyokin()
        SetSpreadBairitsu()
        SetSpreadRiyobi()

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 【画面表示用】分類一覧をSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadBunrui()
        Dim dataTable As DataTable = dataEXTZ0212.PropDsTanakaBairitu.tables("RBUNRUI_MST")
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBunrui.ActiveSheet
        sheet.RowCount = dataTable.Rows.Count
        sheet.ColumnCount = 2
        'カラム設定(コード列の非表示)
        Dim col As FarPoint.Win.Spread.Column
        col = sheet.Columns(1)
        col.Visible = False

        Dim index As New Integer
        index = 0
        For Each row As DataRow In dataTable.Rows
            sheet.Cells(index, 0).Value = row("bunrui_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("bunrui_cd")
            index = index + 1
        Next
        'スクロールバーを必要な場合のみ表示させます
        Me.fbBunrui.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbBunrui.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 【画面表示用】料金一覧をSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadRyokin()
        Dim dataTable As DataTable = dataEXTZ0212.PropDsTanakaBairitu.tables("RIYORYO_MST")
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbRyokin.ActiveSheet
        sheet.RowCount = dataTable.Rows.Count
        sheet.ColumnCount = 6
        'カラム設定(コード列の非表示)
        Dim col As FarPoint.Win.Spread.Column
        col = sheet.Columns(5)
        col.Visible = False

        Dim index As New Integer
        index = 0
        For Each row As DataRow In dataTable.Rows
            sheet.Cells(index, 0).Value = row("ryokin_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("ryokin_hour")
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = False
            sheet.Cells(index, 3).Value = row("ryokin")
            sheet.Cells(index, 3).Locked = True
            sheet.Cells(index, 4).Value = False
            sheet.Cells(index, 5).Value = row("ryokin_cd")
            index = index + 1
        Next
        'スクロールバーを必要な場合のみ表示させます
        Me.fbRyokin.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbRyokin.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 【画面表示用】倍率一覧をSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadBairitsu()
        Dim dataTable As DataTable = dataEXTZ0212.PropDsTanakaBairitu.tables("BAIRITU_MST")
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBairitu.ActiveSheet
        sheet.RowCount = dataTable.Rows.Count
        sheet.ColumnCount = 4
        'カラム設定(コード列の非表示)
        Dim col As FarPoint.Win.Spread.Column
        col = sheet.Columns(3)
        col.Visible = False

        Dim index As New Integer
        index = 0
        For Each row As DataRow In dataTable.Rows
            sheet.Cells(index, 0).Value = row("bairitu_nm")
            sheet.Cells(index, 0).Locked = True
            sheet.Cells(index, 1).Value = row("bairitu")
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = False
            sheet.Cells(index, 3).Value = row("bairitu_cd")
            index = index + 1
        Next

        'スクロールバーを必要な場合のみ表示させます
        Me.fbBairitu.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbBairitu.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 【画面表示用】利用日時一覧をSPREADに設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetSpreadRiyobi()
        Dim aryListRiyobi = dataEXTZ0212.PropLstRiyobi
        'SPREAD クリア
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbRiyobi.ActiveSheet
        sheet.RowCount = aryListRiyobi.Count
        sheet.ColumnCount = 3
        'カラム設定(コード列の非表示)
        Dim col As FarPoint.Win.Spread.Column
        col = sheet.Columns(2)
        col.Visible = False

        Dim index As New Integer
        index = 0
        For Each row As CommonDataRiyobi In aryListRiyobi
            sheet.Cells(index, 0).Value = False
            If row.PropBlnSelect Then
                sheet.Cells(index, 0).Value = True
            End If
            sheet.Cells(index, 1).Value = row.PropStrYoyakuDtDisp
            sheet.Cells(index, 1).Locked = True
            sheet.Cells(index, 2).Value = row.PropStrYoyakuDt
            index = index + 1
        Next

        'スクロールバーを必要な場合のみ表示させます
        Me.fbRiyobi.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbRiyobi.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 分類クリックイベント(マウスクリック)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub fbBunrui_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles fbBunrui.CellClick, fbBunrui.CellDoubleClick

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If e.ColumnHeader = True Then
            Exit Sub
        End If
        If e.RowHeader = True Then
            Exit Sub
        End If
        dataEXTZ0212.PropDsTanakaBairitu.Tables.Remove("RIYORYO_MST")
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBunrui.ActiveSheet
        logicEXTZ0212.GetRyokin(dataEXTZ0212, sheet.Cells(e.Row, 1).Value)
        SetSpreadRyokin()
    End Sub

    ''' <summary>
    ''' 分類選択イベント(キーで選択)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub fbBunrui_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles fbBunrui.LeaveCell

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        dataEXTZ0212.PropDsTanakaBairitu.Tables.Remove("RIYORYO_MST")
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbBunrui.ActiveSheet
        logicEXTZ0212.GetRyokin(dataEXTZ0212, sheet.Cells(e.NewRow, 1).Value)
        SetSpreadRyokin()
    End Sub

    ''' <summary>
    ''' 確定ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnKakutei_Click(sender As Object, e As EventArgs) Handles btnKakutei.Click
        Dim index As Integer = 0
        Dim tanka As String = Nothing
        Dim bairitu As String = Nothing
        Dim ryokin As Decimal = Nothing
        Dim sheetRyokin As FarPoint.Win.Spread.SheetView = Me.fbRyokin.ActiveSheet
        Dim sheetBairitu As FarPoint.Win.Spread.SheetView = Me.fbBairitu.ActiveSheet
        Dim sheetRiyobi As FarPoint.Win.Spread.SheetView = Me.fbRiyobi.ActiveSheet
        Dim listRiyobi As ArrayList = dataEXTZ0212.PropLstRiyobi
        '料金の取得
        For Each row As Row In sheetRyokin.Rows
            If sheetRyokin.Cells(index, 2).Value = True Then
                If tanka IsNot Nothing Then
                    MsgBox(String.Format(CommonDeclareEXT.E2035, "料金"), MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                tanka = sheetRyokin.Cells(index, 1).Value
            End If
            If sheetRyokin.Cells(index, 4).Value = True Then
                If tanka IsNot Nothing Then
                    MsgBox(String.Format(CommonDeclareEXT.E2035, "料金"), MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                tanka = sheetRyokin.Cells(index, 3).Value
            End If
            index = index + 1
        Next
        '倍率の取得
        index = 0
        For Each row As Row In sheetBairitu.Rows
            If sheetBairitu.Cells(index, 2).Value = True Then
                If bairitu IsNot Nothing Then
                    MsgBox(String.Format(CommonDeclareEXT.E2035, "倍率"), MsgBoxStyle.Exclamation, "エラー")
                    Return
                End If
                bairitu = sheetBairitu.Cells(index, 1).Value
            End If
            index = index + 1
        Next
        '反映
        If tanka Is Nothing Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "料金"), MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        If bairitu Is Nothing Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "倍率"), MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        index = 0
        Dim blnSelected As Boolean = False
        For Each row As CommonDataRiyobi In listRiyobi
            If sheetRiyobi.Cells(index, 0).Value = True Then
                row.PropIntTanka = Integer.Parse(tanka)
                row.PropDblBairitu = Double.Parse(bairitu)
                ryokin = row.PropIntTanka * row.PropDblBairitu
                row.PropIntRiyoKin = Integer.Parse(Math.Round(ryokin))
                blnSelected = True
            End If
            index = index + 1
        Next
        If blnSelected = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "反映する日付"), MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        '変更フラグ
        dataEXTZ0212.PropBlnChangeFlg = True
        '画面を閉じる
        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻るボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        '画面を閉じる
        Me.Close()
    End Sub

    ''' <summary>
    ''' 料金一覧シートクリック時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>料金一覧シートのセルをクリックした際の処理(単一選択時の疑似ラジオボックス処理）
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub fbRyokin_ButtonClicked(sender As Object, e As EditorNotifyEventArgs) Handles fbRyokin.ButtonClicked

        RemoveHandler fbRyokin.ButtonClicked, AddressOf fbRyokin_ButtonClicked

        '選択行番号取得
        dataEXTZ0212.PropIntCheckRow = e.Row
        dataEXTZ0212.PropIntCheckCol = e.Column

        '選択行にチェックをつけ、それ以外はチェックを外す
        If logicEXTZ0212.ClickRyokinCellMain(dataEXTZ0212) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        AddHandler fbRyokin.ButtonClicked, AddressOf fbRyokin_ButtonClicked

    End Sub

    ''' <summary>
    ''' 倍率一覧シートクリック時の処理
    ''' </summary>
    ''' <param name="sender">[IN]</param>
    ''' <param name="e">[IN]</param>
    ''' <remarks>倍率一覧シートのセルをクリックした際の処理(単一選択時の疑似ラジオボックス処理）
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub fbBairitu_ButtonClicked(sender As Object, e As EditorNotifyEventArgs) Handles fbBairitu.ButtonClicked

        RemoveHandler fbBairitu.ButtonClicked, AddressOf fbBairitu_ButtonClicked

        '選択行番号取得
        dataEXTZ0212.PropIntCheckRow = e.Row
        dataEXTZ0212.PropIntCheckCol = e.Column

        '選択行にチェックをつけ、それ以外はチェックを外す
        If logicEXTZ0212.ClickBairituCellMain(dataEXTZ0212) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Exit Sub
        End If

        AddHandler fbBairitu.ButtonClicked, AddressOf fbBairitu_ButtonClicked

    End Sub
End Class
