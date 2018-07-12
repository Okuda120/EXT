Imports Common
Imports CommonEXT
Imports System.Net

Public Class EXTZ0213

    Private ppStrIraiOrKiroku As String '1=依頼,2=記録
    Private ppStrYoyakuNo As String '予約No
    Private ppDsApproval As DataSet 'DS
    Private ppBlnChangeFlg As Boolean
    Private commonLogicEXT As New CommonLogicEXT      '共通ロジッククラス

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0213_Load(sender As Object, e As EventArgs) Handles Me.Load
        PropBlnChangeFlg = False
        If Me.PropStrIraiOrKiroku = SHONIN_IRAI Then
            '承認依頼の場合
            Me.Text = "ちゃり～ん。　承認依頼追加"
            Me.lblTitle1.Text = "依頼者"
            Me.lblTitle2.Text = "依頼日　＊"
            Me.lblTitle3.Visible = False
            Me.cmbStatus.Visible = False
            Me.lblName.Text = CommonDeclareEXT.PropComStrUserName
            Dim dt As DateTime = System.DateTime.Now
            Me.dtpIraiCheck.txtDate.Text = dt.ToString("yyyy/MM/dd HH:mm")
        ElseIf Me.PropStrIraiOrKiroku = SHONIN_KIROKU Then
            '記録追加の場合
            Me.Text = "ちゃり～ん。　承認記録追加"
            Me.lblTitle1.Text = "承認者"
            Me.lblTitle2.Text = "確認日　＊"
            Me.lblTitle3.Visible = True
            Me.cmbStatus.Visible = True
            Me.cmbStatus.Items.Add("承認")
            Me.cmbStatus.Items.Add("差し戻し")
            Me.lblName.Text = CommonDeclareEXT.PropComStrUserName
            Dim dt As DateTime = System.DateTime.Now
            Me.dtpIraiCheck.txtDate.Text = dt.ToString("yyyy/MM/dd HH:mm")
        End If

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click
        If inputCheckMain() = False Then
            Exit Sub
        End If

        Dim index As New Integer
        Dim table As DataTable
        Dim row As DataRow
        index = 0
        If Me.PropStrIraiOrKiroku = SHONIN_IRAI Then
            table = PropDsApproval.Tables("IRAI_RIREKI_TBL")
            row = table.NewRow
            row("seq") = DBNull.Value
            Dim dt As DateTime = Me.dtpIraiCheck.txtDate.Text
            row("irai_dt") = dt.ToString("yyyy/MM/dd HH:mm")
            row("com") = Me.txtCom.Text
            row("user_nm") = CommonDeclareEXT.PropComStrUserName
            table.Rows.Add(row)
        ElseIf Me.PropStrIraiOrKiroku = SHONIN_KIROKU Then
            table = PropDsApproval.Tables("CHECK_RIREKI_TBL")
            row = table.NewRow
            row("seq") = DBNull.Value
            Dim dt As DateTime = Me.dtpIraiCheck.txtDate.Text
            row("check_dt") = dt.ToString("yyyy/MM/dd HH:mm")
            If Me.cmbStatus.Text = "承認" Then
                row("check_sts") = "1"
            Else
                row("check_sts") = "2"
            End If
            row("com") = Me.txtCom.Text
            row("user_nm") = CommonDeclareEXT.PropComStrUserName
            table.Rows.Add(row)
        End If
        PropBlnChangeFlg = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻るボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function inputCheckMain() As Boolean
        'ステータス
        If Me.PropStrIraiOrKiroku = SHONIN_KIROKU Then
            If String.IsNullOrEmpty(Me.cmbStatus.Text) Then
                MsgBox(String.Format(CommonDeclareEXT.E0002, "確認ステータス"), MsgBoxStyle.Exclamation, "エラー")
                Return False
            End If
        End If
        '日付
        If String.IsNullOrEmpty(Me.dtpIraiCheck.txtDate.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0001, Me.lblTitle2.Text), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' プロパティセット【承認依頼記録判定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrIraiOrKiroku</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrIraiOrKiroku()
        Get
            Return ppStrIraiOrKiroku
        End Get
        Set(ByVal value)
            ppStrIraiOrKiroku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:承認依頼／記録】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsApproval()
        Get
            Return ppDsApproval
        End Get
        Set(ByVal value)
            ppDsApproval = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
        End Set
    End Property

End Class