Imports Common
Imports CommonEXT
Imports EXTB
Imports EXTC
Imports EXTM
Imports EXTY
Imports EXTZ

''' <summary>
''' EXTA0102
''' </summary>
''' <remarks>メニュー画面
''' <para>作成情報：2015/08/04 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTA0102

    Public commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    Private isCloseOk As Boolean = False

    ''' <summary>
    ''' 「メニュー画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>ログイン情報を基にユーザが行える機能を表示する
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTA0102_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        'マスタ操作権限の判定
        If CommonDeclareEXT.PropStrComFlgMst.ToString <> "1" Then
            '権限がない場合は非活性
            Me.btnUserMaster.Enabled = False
            Me.btnFutaiSetsubiMaster.Enabled = False
            Me.btnRyokinMaster.Enabled = False
            Me.btnExciseTaxMaster.Enabled = False
        Else '権限がある場合は活性
            Me.btnUserMaster.Enabled = True
            Me.btnFutaiSetsubiMaster.Enabled = True
            Me.btnRyokinMaster.Enabled = True
            Me.btnExciseTaxMaster.Enabled = True
        End If

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
            Me.lblKenshou.Visible = True
        Else
            Me.lblKenshou.Visible = False
        End If

    End Sub

    ''' <summary>
    ''' 「予約カレンダー（シアター）」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnCalendarTheater_Click(sender As Object, e As EventArgs) Handles btnCalendarTheater.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.08.12 ADD START↓ m.hayabuchi
        'データセット
        If Me.ChkDaikoFlg.Checked = True Then
            CommonDeclareEXT.PropStrDaikoFlg = "1"
        Else
            CommonDeclareEXT.PropStrDaikoFlg = ""
        End If
        ' 2016.08.12 ADD END↑ m.hayabuchi

        Dim frm As New EXTB0101
        Me.Hide()

        '「予約カレンダー（シアター）」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「予約カレンダー（スタジオ）」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnCalendarStudio_Click(sender As Object, e As EventArgs) Handles btnCalendarStudio.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.08.12 ADD START↓ m.hayabuchi
        'データセット
        If Me.ChkDaikoFlg.Checked = True Then
            CommonDeclareEXT.PropStrDaikoFlg = "1"
        Else
            CommonDeclareEXT.PropStrDaikoFlg = ""
        End If
        ' 2016.08.12 ADD END↑ m.hayabuchi

        Dim frm As New EXTC0101
        Me.Hide()

        '「予約カレンダー（スタジオ）」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「予約一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnYoyakuList_Click(sender As Object, e As EventArgs) Handles btnYoyakuList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.08.12 ADD START↓ m.hayabuchi
        'データセット
        If Me.ChkDaikoFlg.Checked = True Then
            CommonDeclareEXT.PropStrDaikoFlg = "1"
        Else
            CommonDeclareEXT.PropStrDaikoFlg = ""
        End If
        ' 2016.08.12 ADD END↑ m.hayabuchi

        Dim frm As New EXTZ0101
        Me.Hide()

        'パラメータに選択モードを設定
        frm.dataEXTZ0101.PropStrSelectMode = SELECTMODE_MENU

        '「予約一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「請求一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSeikyuList_Click(sender As Object, e As EventArgs) Handles btnSeikyuList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.08.12 ADD START↓ m.hayabuchi
        'データセット
        If Me.ChkDaikoFlg.Checked = True Then
            CommonDeclareEXT.PropStrDaikoFlg = "1"
        Else
            CommonDeclareEXT.PropStrDaikoFlg = ""
        End If
        ' 2016.08.12 ADD END↑ m.hayabuchi

        Dim frm As New EXTZ0102
        Me.Hide()

        '「請求一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「EXAS入金一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnExasPayList_Click(sender As Object, e As EventArgs) Handles btnExasPayList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0208
        Me.Hide()

        '「EXAS入金一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「ALSOK入金管理」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnAlsokPayManage_Click(sender As Object, e As EventArgs) Handles btnAlsokPayManage.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTY0102
        Me.Hide()

        '「ALSOK入金管理」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「EXAS請求依頼」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnExasSeikyuRequest_Click(sender As Object, e As EventArgs) Handles btnExasSeikyuRequest.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 2016.08.12 ADD START↓ m.hayabuchi
        'データセット
        If Me.ChkDaikoFlg.Checked = True Then
            CommonDeclareEXT.PropStrDaikoFlg = "1"
        Else
            CommonDeclareEXT.PropStrDaikoFlg = ""
        End If
        ' 2016.08.12 ADD END↑ m.hayabuchi

        Dim frm As New EXTY0101
        Me.Hide()

        '「EXAS請求依頼」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「利用者一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnUserList_Click(sender As Object, e As EventArgs) Handles btnUserList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0201
        Me.Hide()

        '「利用者一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「EXAS相手先一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnAitesakiList_Click(sender As Object, e As EventArgs) Handles btnAitesakiList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0203
        Me.Hide()

        '「EXAS相手先一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「日別売上一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/06 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnHibetsuUriageList_Click(sender As Object, e As EventArgs) Handles btnHibetsuUriageList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0103
        Me.Hide()

        '「日別売上一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「利用状況一覧」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/06 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRiyoJokyoList_Click(sender As Object, e As EventArgs) Handles btnRiyoJokyoList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTZ0104
        Me.Hide()

        '「利用状況一覧」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「ユーザマスタ」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnUserMaster_Click(sender As Object, e As EventArgs) Handles btnUserMaster.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0101
        Me.Hide()

        '「ユーザマスタ」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「付帯設備マスタ押下」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnFutaiSetsubiMaster_Click(sender As Object, e As EventArgs) Handles btnFutaiSetsubiMaster.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0102
        Me.Hide()

        '「付帯設備マスタ押下」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「料金マスタ押下」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRyokinMaster_Click(sender As Object, e As EventArgs) Handles btnRyokinMaster.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0104
        Me.Hide()

        '「料金マスタ押下」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「消費税マスタ」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnExciseTaxMaster_Click(sender As Object, e As EventArgs) Handles btnExciseTaxMaster.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0103
        Me.Hide()

        '「消費税マスタ」画面を表示
        frm.ShowDialog()
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「終了」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

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
    Private Sub EXTA0102_FormClosing(ByVal sender As Object, _
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

End Class