Imports Common
Imports CommonEXT
Imports System.Net

Public Class EXTZ0214

    Private commonLogicEXT As New CommonLogicEXT        '共通ロジッククラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Public CommonValidation As New CommonValidation     'CommonValidation
    Public dataEXTZ0214 As New DataEXTZ0214             'データクラス
    Public logicEXTZ0214 As New LogicEXTZ0214           'ロジッククラス
    Private strSetflg As String = ""                        ' 2016.02.02. ADD h.hagiwara

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0214_Load(sender As Object, e As EventArgs) Handles Me.Load
        
        
        Me.lblExasAite.Text = dataEXTZ0214.PropStrAitesakiCD
        Me.lblExasAiteNm.Text = dataEXTZ0214.PropStrAitesakiNm
        dataEXTZ0214.PropCmbSeikyusakiBusyo = Me.cmbSeikyusakiBusyo
        dataEXTZ0214.PropCmbSeikyuNaiyo = Me.cmbSeikyuNaiyo

        '請求先部署、請求内容コンボボックスのセット
        ' 請求先部署、請求内容ファイルの取り込み
        If logicEXTZ0214.GetCmbCsvData(dataEXTZ0214) = False Then
            'エラーとしない
        End If

        strSetflg = ""                        ' 2016.02.02. ADD h.hagiwara

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
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
    Private Sub btnFileSyutu_Click(sender As Object, e As EventArgs) Handles btnFileSyutu.Click
        If inputCheckMain() = False Then
            Exit Sub
        End If
        '画面の入力値を設定
        dataEXTZ0214.PropStrSeikyusakiBusyoCD = Me.cmbSeikyusakiBusyo.SelectedValue
        dataEXTZ0214.PropStrSeikyusyoTantoCD = Me.txtTantoCD.Text
        dataEXTZ0214.PropStrSeikyuNaiyoCD = Me.cmbSeikyuNaiyo.SelectedValue
        'グループ請求情報出力用のフラグを設定
        dataEXTZ0214.PropBlnGSeikyuFlg = True

        strSetflg = "1"                        ' 2016.02.02. ADD h.hagiwara

        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻るボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        'Exit Sub用フラグの設定
        dataEXTZ0214.PropBlnGSeikyuCloseFlg = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2016/01/18 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function inputCheckMain() As Boolean
        'G請求先部署 
        If String.IsNullOrEmpty(Me.cmbSeikyusakiBusyo.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "G請求先部署"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If

        'G請求書担当者コード
        If String.IsNullOrEmpty(Me.txtTantoCD.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0001, "G請求書担当者コード"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        '半角英数チェック
        If CommonValidation.IsHalfChar(Me.txtTantoCD.Text) = False Then
            puErrMsg = String.Format(CommonDeclareEXT.E0005, "G請求書担当者コード")
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return False
        End If
        ' 2016.04.04 UPD START↓ h.hagiwara 桁チェック修正
        ''桁数チェック（6桁）
        'If Len(Me.txtTantoCD.Text) <> 6 Then
        '    puErrMsg = String.Format(CommonDeclareEXT.E0010, "G請求書担当者コード", "6")
        '    MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
        '    Return False
        'End If
        '桁数チェック
        If Len(Me.txtTantoCD.Text) > 6 Then
            puErrMsg = String.Format(CommonDeclareEXT.E0024, "G請求書担当者コード", "6")
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return False
        End If
        ' 2016.04.04 UPD END↑ h.hagiwara 桁チェック修正

        'G請求内容 
        If String.IsNullOrEmpty(Me.cmbSeikyuNaiyo.Text) Then
            MsgBox(String.Format(CommonDeclareEXT.E0002, "G請求内容"), MsgBoxStyle.Exclamation, "エラー")
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 閉じるButton処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>未更新時無効
    ''' <para>作成情報： 2016.02.02 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0214_FormClosing(ByVal sender As Object, _
            ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing

        If strSetflg <> "1" Then
            'Exit Sub用フラグの設定
            dataEXTZ0214.PropBlnGSeikyuCloseFlg = True
        End If

    End Sub

End Class