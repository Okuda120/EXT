Imports Common
Imports CommonEXT

''' <summary>
''' EXTB0105
''' </summary>
''' <remarks>その他施設利用登録（シアター）画面
''' <para>作成情報：2015/08/26 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>

Public Class EXTB0105

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTB0105 As New LogicEXTB0105      'ロジッククラス
    Public dataEXTB0105 As New DataEXTB0105         'データクラス
    Public commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「その他施設利用登録（シアター）画面」ロード時処理 "

    ''' <summary>
    ''' 「その他施設利用登録（シアター）画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「その他施設利用登録（シアター）画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTB0105_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '変数宣言
        Dim strWkDate As String = String.Empty '作業用日付
        Dim strYobi As String = String.Empty '曜日

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        'その他情報取得
        If logicEXTB0105.GetSonotaInfo(dataEXTB0105) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        '画面表示
        With dataEXTB0105
            strYobi = WeekdayName(Weekday(.PropStrRiyoDt)).Substring(0, 1)
            strWkDate = Date.Parse(.PropStrRiyoDt).ToString("yyyy/M/d") & "(" & strYobi & ")"
            Me.lblRiyoDt.Text = strWkDate
            If .PropIntDataCnt > 0 Then
                ' 2015.12.04 UPD START↓ h.hagiwara
                'Me.txtStartTime.Text = .PropStrStartTime.Substring(0, 2) & ":" & .PropStrStartTime.Substring(2, 2)
                'Me.txtEndTime.Text = .PropStrEndTime.Substring(0, 2) & ":" & .PropStrEndTime.Substring(2, 2)
                Me.txtStartTime.PropTxtTime.Text = .PropStrStartTime.Substring(0, 2) & ":" & .PropStrStartTime.Substring(2, 2)
                Me.txtEndTime.PropTxtTime.Text = .PropStrEndTime.Substring(0, 2) & ":" & .PropStrEndTime.Substring(2, 2)
                ' 2015.12.04 UPD END↑ h.hagiwara
            Else
                Me.txtStartTime.Text = .PropStrStartTime
                Me.txtEndTime.Text = .PropStrEndTime
            End If
            Me.txtRiyoBasho.Text = .PropStrRiyoBasho
            Me.txtRiyoYoto.Text = .PropStrRiyoYoto
            Me.txtRiyosha.Text = .PropStrRiyosha
        End With

        '削除ボタン制御
        If dataEXTB0105.PropIntDataCnt > 0 Then
            Me.btnDelete.Enabled = True
        Else
            Me.btnDelete.Enabled = False
        End If

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

#Region "「戻る」ボタン押下時処理 "

    ''' <summary>
    ''' 「戻る」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「戻る」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "「削除」ボタン押下時処理 "

    ''' <summary>
    ''' 「削除」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「削除」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'その他情報更新処理
        If logicEXTB0105.DeleteSonotaInfo(dataEXTB0105) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

        '戻り値をOKに設定
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "「登録」ボタン押下時処理 "

    ''' <summary>
    ''' 「登録」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「登録」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRegist_Click(sender As Object, e As EventArgs) Handles btnRegist.Click

        '変数宣言
        Dim datResult As Date = Nothing     '時刻フォーマット変換結果

        '入力チェック
        With dataEXTB0105
            '---必須チェック
            ' 2015.12.04 UPD START↓ h.hagiwara
            'If Me.txtStartTime.Text = String.Empty Then
            '    MsgBox(String.Format(CommonEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            '    Return
            'End If
            'If Me.txtEndTime.Text = String.Empty Then
            '    MsgBox(String.Format(CommonEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            '    Return
            'End If
            If Me.txtStartTime.PropTxtTime.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "開始時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
            If Me.txtEndTime.PropTxtTime.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "終了時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
            ' 2015.12.04 UPD END↑ h.hagiwara
            If Me.txtRiyoBasho.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "場所"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
            If Me.txtRiyoYoto.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "用途"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
            If Me.txtRiyosha.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "利用者"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If

            ' 2015.12.04 DEL START↓ h.hagiwara
            '---フォーマットチェック(HH:MM)
            'If System.Text.RegularExpressions.Regex.IsMatch(Me.txtStartTime.Text, "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$") = False Then
            '    MsgBox(String.Format(CommonEXT.E2007, "開始時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            '    Return
            'End If
            'If System.Text.RegularExpressions.Regex.IsMatch(Me.txtEndTime.Text, "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$") = False Then
            '    MsgBox(String.Format(CommonEXT.E2007, "終了時間"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            '    Return
            'End If
            ' 2015.12.04 DEL END↑ h.hagiwara

            'データクラスに格納
            ' 2015.12.04 UPD START↓ h.hagiwara
            '.PropStrStartTime = Replace(Me.txtStartTime.Text, ":", "")
            '.PropStrEndTime = Replace(Me.txtEndTime.Text, ":", "")
            .PropStrStartTime = Replace(Me.txtStartTime.PropTxtTime.Text, ":", "")
            .PropStrEndTime = Replace(Me.txtEndTime.PropTxtTime.Text, ":", "")
            ' 2015.12.04 UPD END↑ h.hagiwara
            .PropStrRiyoBasho = Me.txtRiyoBasho.Text
            .PropStrRiyoYoto = Me.txtRiyoYoto.Text
            .PropStrRiyosha = Me.txtRiyosha.Text

        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'その他情報更新処理
        If logicEXTB0105.UpdateSonotaInfo(dataEXTB0105) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

        '戻り値をOKに設定
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

End Class