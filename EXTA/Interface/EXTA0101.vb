Imports Common
Imports CommonEXT

Public Class EXTA0101

    Private commonLogic As New CommonLogic      '共通クラス
    Private isCloseOk As Boolean = False

    '変数宣言
    Public dataCommon As New CommonDataEXT         '共通データクラス
    Public commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス
    Public dataEXTA0101 As New DataEXTA0101        'データクラス
    Public logicEXTA0101 As New LogicEXTA0101      'ロジッククラス

    ''' <summary>
    ''' 「ログイン画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>初期化処理を行う
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EKJA0101_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        '共通設定値取得
        If commonLogic.InitCommonSetting(Nothing) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
            Return
        End If

        'ログ出力設定（操作ログ）
        If logicEXTA0101.SetOpLog() = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
            Return
        End If

        ' 背景色用フラグ取得
        If logicEXTA0101.GetConfigrationFlg(Nothing) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
            Return
        End If

        '初期化
        Me.txtUserId.Text = String.Empty    'ユーザーID
        Me.txtPassword.Text = String.Empty  'パスワード

        Dim cont As Control = Me.ActiveControl
        If Not (cont Is Nothing) = False Then
            Me.txtUserId.Focus()
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' ログインボタンクリック処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>ログイン処理
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        'プロパティにセット
        With dataEXTA0101
            .PropTxtUserId = Me.txtUserId       'ユーザーID
            .PropTxtPassword = Me.txtPassword   'password
        End With

        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'ログイン処理メインメソッド
        '入力チェック処理
        If logicEXTA0101.LoginInputCheck(dataEXTA0101) = False Then
            Exit Sub
        End If

        'ログイン情報の取得
        If logicEXTA0101.GetLoginData(dataEXTA0101) = False Then
            'メッセージを出力 
            MsgBox(CommonEXT.E1001, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        'ログイン情報設定
        CommonDeclareEXT.PropComStrUserId = DirectCast(dataEXTA0101.PropTxtUserId, TextBox).Text
        CommonDeclareEXT.PropComStrUserName = dataEXTA0101.PropStrUserName
        CommonDeclareEXT.PropStrComMailAddr = dataEXTA0101.PropStrMailAddr
        CommonDeclareEXT.PropStrComBusho = dataEXTA0101.PropStrCode_BUSHO
        CommonDeclareEXT.PropStrComFlgShonin = dataEXTA0101.PropStrFlg_SHOHIN
        CommonDeclareEXT.PropStrComFlgMst = dataEXTA0101.PropStrFlg_MST
        'ログイン成功時処理
        If logicEXTA0101.DelYoyakuCtlData(dataEXTA0101) = False Then
            'メッセージを出力 
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If logicEXTA0101.SetSystemProperties() = False Then
            'メッセージを出力 
            MsgBox(CommonEXT.E0000, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'メニュー遷移処理
        Dim frm As New EXTA0102

        'ログインログ
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "LOGIN:" & CommonDeclareEXT.PropComStrUserId, Nothing, Nothing)

        Me.Hide()

        '「メニュー」画面を表示
        frm.ShowDialog()
        'Me.Show()
        isCloseOk = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' EXITボタンクリック処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>閉じる処理
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        '終了確認
        If MsgBox(String.Format(CommonEXT.I0004), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            '処理を抜ける
            Return
        End If

        '画面を閉じる
        isCloseOk = True
        Me.Close()

    End Sub

    ''' <summary>
    ''' 閉じるButton処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTA0101_FormClosing(ByVal sender As Object, _
            ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        '終了確認
        If isCloseOk = True Then
            'exit
        ElseIf MsgBox(String.Format(CommonEXT.I0004), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then
            '処理を抜ける
            e.Cancel = True
            Return
        End If
    End Sub

    ''' <summary>
    ''' 「！」Button処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCapture_Click(sender As Object, e As EventArgs) Handles btnCapture.Click

        If logicEXTA0101.OpenCaptureDir() = False Then
            'メッセージを出力 
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Return

    End Sub

End Class
