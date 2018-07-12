Imports Common
Imports CommonEXT

Public Class EXTY0101

    'インスタンスを生成
    Public dataEXTY0101 As New DataEXTY0101
    Private logicEXTY0101 As New LogicEXTY0101
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス


    ''' <summary>
    ''' 初期表示処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>EXAS請求依頼データ作成画面初期表示処理を行う
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub EXTY0101_Load(sender As Object, e As EventArgs) Handles Me.Load

        'データクラスの初期設定を行う
        With dataEXTY0101
            .PropRdoStudio = Me.rdoStudio
            .PropRdoTheater = Me.rdoTheater
            .PropVwBillpay = Me.vwBillpay
        End With

        'シアターにチェックを入れる
        dataEXTY0101.PropRdoTheater.Checked = True

        '表示処理を行う
        If logicEXTY0101.InitDisplayMain(dataEXTY0101) = False Then
            MsgBox(puErrMsg)
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 「表示」ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>選択されたラジオボタンに紐付く表示処理を行う
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnDisplay_Click(sender As Object, e As EventArgs) Handles btnDisplay.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '表示処理を行う
        If logicEXTY0101.InitDisplayMain(dataEXTY0101) = False Then
            MsgBox(puErrMsg)
        End If

    End Sub

    ''' <summary>
    ''' 「請求依頼データ出力」ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>EXAS請求依頼データの出力処理を行う
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'CSVファイル出力処理を行う
        If logicEXTY0101.OutputCsvMain(dataEXTY0101) = False Then
            MsgBox(puErrMsg)
        End If

    End Sub

    ''' <summary>
    ''' 「戻る」ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「戻る」ボタン押下時の処理を行う
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        'ウィンドウを閉じる
        Me.Close()

    End Sub

    ' 2016.01.07 DEL START↓ h.hagiwara
    ' ''' <summary>
    ' ''' スプレッドシートクリック時の処理
    ' ''' </summary>
    ' ''' <param name="sender"></param>
    ' ''' <param name="e"></param>
    ' ''' <remarks>チェックボックスの制御を行う
    ' ''' <para>作成情報：2015/12/01 y.ozawa
    ' ''' <p>改訂情報：</p>
    ' ''' </para></remarks>
    'Private Sub vwBillpay_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwBillpay.ButtonClicked
    '    '選択行をクリックした場合
    '    If e.Column = 0 Then

    '        RemoveHandler vwBillpay.ButtonClicked, AddressOf vwBillpay_ButtonClicked

    '        '選択行番号取得
    '        dataEXTY0101.PropIntCheckRow = e.Row

    '        '選択行にチェックをつけ、それ以外はチェックを外す
    '        If logicEXTY0101.ClickVwCellMain(dataEXTY0101) = False Then
    '            'エラーメッセージ表示
    '            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
    '            Exit Sub
    '        End If

    '        AddHandler vwBillpay.ButtonClicked, AddressOf vwBillpay_ButtonClicked

    '    End If
    'End Sub
    ' 2016.01.07 DEL END↑ h.hagiwara

End Class