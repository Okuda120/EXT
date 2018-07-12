Imports Common
Imports CommonEXT

Public Class EXTZ0211

    Private commonLogicEXT As New CommonLogicEXT      '共通ロジッククラス
    Private ppStrDeleteKbn As String              '取消区分

#Region "「キャンセル画面」ロード時処理 "

    ''' <summary>
    ''' 「キャンセル画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「利用日画面」ロード時の処理を行う 
    ''' <para>作成情報： 2015.11.18 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0211_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region


    ''' <summary>
    ''' 確定ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        If Me.rdoCancel.Checked Then
            PropStrDeleteKbn = CommonDeclareEXT.TORIKESHI_KBN_CANCEL
        End If
        If Me.rdoDelete.Checked Then
            PropStrDeleteKbn = CommonDeclareEXT.TORIKESHI_KBN_DELETE
        End If
        '画面を閉じる
        '戻り値をOKにする
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
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
        '戻り値をOKにする
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' プロパティセット【取消区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDeleteKbn()
        Get
            Return ppStrDeleteKbn
        End Get
        Set(ByVal value)
            ppStrDeleteKbn = value
        End Set
    End Property

End Class
