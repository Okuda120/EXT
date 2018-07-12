Imports Common
Imports CommonEXT

Public Class EXTM0203

    'インスタンス生成
    Public dataEXTM0203 As New DataEXTM0203
    Private logicEXTM0203 As New LogicEXTM0203
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    '変数宣言
    Private aryReturnValue As New ArrayList
    Private strSetflg As String = ""                        ' 2016.02.02. ADD h.hagiwara

    ''' <summary>
    ''' 画面表示時の処理
    ''' </summary>
    ''' <returns>EXAS相手先情報ArrayList</returns>
    ''' <remarks>画面のポップアップ表示と戻り値の設定を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Overloads Function ShowDialog() As ArrayList

        '当画面をポップアップ表示
        MyBase.ShowDialog()

        ' 2016.02.02 ADD START↓ h.hagiwara
        If strSetflg <> "1" Then
            '戻り値にNothingを設定
            aryReturnValue = Nothing
        End If
        ' 2016.02.02 ADD END ↑ h.hagiwara

        '戻り値を返す
        Return aryReturnValue

    End Function

    ''' <summary>
    ''' EXAS相手先一覧画面ロード時処理
    ''' </summary>
    ''' <remarks>画面ロード時の処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub EXTM0203_Load(sender As Object, e As EventArgs) Handles MyBase.Shown

        'データクラスの初期設定を行う
        With dataEXTM0203
            .PropTxtExasCode = Me.txtExasCode
            .PropTxtExasName = Me.txtExasName
            .PropVwResult = Me.vwResult

            ' 2015.10.08 ADD START↓ h.hagiwara 初期表示時のクリア対応
            If .PropVwResult.ActiveSheet.RowCount > 0 Then
                .PropVwResult.Sheets(0).Rows.Remove(0, .PropVwResult.ActiveSheet.RowCount)
            End If
            ' 2015.10.08 ADD END↑ h.hagiwara 初期表示時のクリア対応

            '画面表示設定を行う
            If .PropStrMenuID = "" Then
                '選択確定ボタンを非活性
                Me.btnConfirm.Enabled = False

            End If

        End With

        strSetflg = ""                        ' 2016.02.02. ADD h.hagiwara

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 「検索」ボタン押下時処理
    ''' </summary>
    ''' <remarks>「検索」ボタン押下時の処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2015.12.03 ADD START↓ h.hagiwara       検索条件の指定を必須とする           
        If logicEXTM0203.SearchInputCheck(dataEXTM0203) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        ' 2015.12.03 ADD END↑ h.hagiwara       検索条件の指定を必須とする           

        'EXAS相手先情報検索処理を行う
        If logicEXTM0203.ExasSearchMain(dataEXTM0203) = False Then
            MsgBox(puErrMsg)
        End If

    End Sub

    ''' <summary>
    ''' 「選択確定」ボタン押下時処理
    ''' </summary>
    ''' <remarks>「選択確定」ボタン押下時の処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェックを行う
        If logicEXTM0203.InputCheck(dataEXTM0203) = False Then
            MsgBox(puErrMsg)
            Return
        End If

        Dim aryValue As New ArrayList

        '戻り値を設定
        'aryReturnValue.Add(dataEXTM0203.PropVwResult.Sheets(0).Cells(dataEXTM0203.PropIntCheckRow, 1).Value)
        'aryReturnValue.Add(dataEXTM0203.PropVwResult.Sheets(0).Cells(dataEXTM0203.PropIntCheckRow, 2).Value)
        aryValue.Add(dataEXTM0203.PropVwResult.Sheets(0).Cells(dataEXTM0203.PropIntCheckRow, 1).Value)
        aryValue.Add(dataEXTM0203.PropVwResult.Sheets(0).Cells(dataEXTM0203.PropIntCheckRow, 2).Value)
        aryValue.Add(dataEXTM0203.PropVwResult.Sheets(0).Cells(dataEXTM0203.PropIntCheckRow, 3).Value)

        If aryReturnValue Is Nothing Then
            aryReturnValue = New ArrayList
        End If

        strSetflg = "1"                        ' 2016.02.02. ADD h.hagiwara
        aryReturnValue = aryValue

        'ウィンドウを閉じる
        Me.Close()

    End Sub

    ''' <summary>
    ''' 「戻る」ボタン押下時処理
    ''' </summary>
    ''' <remarks>「戻る」ボタン押下時の処理を行う
    ''' <para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        '戻り値にNothingを設定
        aryReturnValue = Nothing
        'ウィンドウを閉じる
        Me.Close()

    End Sub

    ''' <summary>
    ''' スプレッドシートクリック時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>チェックボックスの制御を行う
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwResult_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwResult.ButtonClicked
        '選択行をクリックした場合
        If e.Column = 0 Then

            RemoveHandler vwResult.ButtonClicked, AddressOf vwResult_ButtonClicked

            '選択行番号取得
            dataEXTM0203.PropIntCheckRow = e.Row

            '選択行にチェックをつけ、それ以外はチェックを外す
            If logicEXTM0203.ClickVwCellMain(dataEXTM0203) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Exit Sub
            End If

            AddHandler vwResult.ButtonClicked, AddressOf vwResult_ButtonClicked

        End If

    End Sub

End Class