Imports Common
Imports CommonEXT

''' <summary>
''' EXTZ0201
''' </summary>
''' <remarks>メンテ日登録画面
''' <para>作成情報：2015/08/28 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>

Public Class EXTZ0201

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTZ0201 As New LogicEXTZ0201     'ロジッククラス
    Public dataEXTZ0201 As New DataEXTZ0201         'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス


#Region "「メンテ日登録画面」ロード時処理 "

    ''' <summary>
    ''' 「メンテ日登録画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「メンテ日登録画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTB0201_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '変数宣言
        Dim strWkDate As String = String.Empty '作業用日付
        Dim strYobi As String = String.Empty '曜日

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '休館メンテ情報取得
        If LogicEXTZ0201.GetHolmentInfo(dataEXTZ0201) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If

        '画面表示
        With dataEXTZ0201
            strYobi = WeekdayName(Weekday(.PropStrHolmentDt)).Substring(0, 1)
            strWkDate = Date.Parse(.PropStrHolmentDt).ToString("yyyy/M/d") & "(" & strYobi & ")"
            Me.lblHolMentDt.Text = strWkDate
            Me.txtComment.Text = .PropStrMnaiyo
        End With

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
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel

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
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRegist_Click(sender As Object, e As EventArgs) Handles btnRegist.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        With dataEXTZ0201
            '---必須チェック
            If Me.txtComment.Text = String.Empty Then
                MsgBox(String.Format(CommonEXT.E0001, "コメント"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If

            'データクラスに格納
            .PropStrMnaiyo = Me.txtComment.Text
        End With

        '休館メンテ情報更新処理
        If dataEXTZ0201.PropStrStudioKbn = Nothing Then
            '複数スタジオを更新する場合
            If logicEXTZ0201.DelInsHolmentInfo(dataEXTZ0201) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
        Else
            '個別スタジオを更新する場合
            If logicEXTZ0201.UpdateHolmentInfo(dataEXTZ0201) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
        End If

        '戻り値をOKに設定
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

   
End Class
