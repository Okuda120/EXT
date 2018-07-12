Imports Common
Imports CommonEXT
Imports EXTM

''' <summary>
''' EXTZ0104
''' </summary>
''' <remarks>使用状況一覧画面
''' <para>作成情報：2015/09/24 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
''' 
Public Class EXTZ0104

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTZ0104 As New LogicEXTZ0104     'ロジッククラス
    Public dataEXTZ0104 As New DataEXTZ0104         'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「使用状況一覧画面」ロード時処理 "

    ''' <summary>
    ''' 「使用状況一覧画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「使用状況一覧画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0104_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '画面初期設定
        Me.rdoTheatre.Checked = True
        Me.txtKikanNen.Text = String.Empty
        Me.txtKikanTsuki.Text = String.Empty
        For intRow As Integer = 0 To Me.vwRiyoJokyoThaetre.ActiveSheet.Rows.Count - 1
            For intCol As Integer = 0 To Me.vwRiyoJokyoThaetre.ActiveSheet.Columns.Count - 1
                If intRow = 0 Then
                    If intCol = 12 Then
                        Me.vwRiyoJokyoThaetre.ActiveSheet.ColumnHeader.Cells(0, intCol).Value = "合計"
                    Else
                        Me.vwRiyoJokyoThaetre.ActiveSheet.ColumnHeader.Cells(0, intCol).Value = String.Empty
                    End If
                End If
                Me.vwRiyoJokyoThaetre.ActiveSheet.Cells(intRow, intCol).Value = 0
            Next
        Next
        For intRow As Integer = 0 To Me.vwRiyoJokyoStudio.ActiveSheet.Rows.Count - 1
            For intCol As Integer = 0 To Me.vwRiyoJokyoStudio.ActiveSheet.Columns.Count - 1
                If intRow = 0 Then
                    If intCol = 12 Then
                        Me.vwRiyoJokyoStudio.ActiveSheet.ColumnHeader.Cells(0, intCol).Value = "合計"
                    Else
                        Me.vwRiyoJokyoStudio.ActiveSheet.ColumnHeader.Cells(0, intCol).Value = String.Empty
                    End If
                End If
                Me.vwRiyoJokyoStudio.ActiveSheet.Cells(intRow, intCol).Value = 0
            Next
        Next
        Me.vwRiyoJokyoStudio.Visible = False

        'プロパティに設定
        dataEXTZ0104.PropvwRiyoJokyoTheatre = Me.vwRiyoJokyoThaetre
        dataEXTZ0104.PropvwRiyoJokyoStudio = Me.vwRiyoJokyoStudio

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

#Region "「表示」ボタン押下時処理 "

    ''' <summary>
    ''' 「表示」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「表示」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        '---必須チェック
        If Me.txtKikanNen.Text = String.Empty Then
            MsgBox(String.Format(CommonEXT.E0001, "年"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If
        If Me.txtKikanTsuki.Text = String.Empty Then
            MsgBox(String.Format(CommonEXT.E0001, "月"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        '---属性チェック(半角数字)
        If commonValidate.IsHalfNmb(Me.txtKikanNen.Text) = False Then
            MsgBox(String.Format(CommonEXT.E0003, "年"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If
        If commonValidate.IsHalfNmb(Me.txtKikanTsuki.Text) = False Then
            MsgBox(String.Format(CommonEXT.E0003, "月"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        '---桁数チェック
        If Me.txtKikanNen.Text.Length < 4 Then
            MsgBox(String.Format(CommonEXT.E0010, "年", "4"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        '存在しない年月の場合、エラーメッセージを出力
        If IsDate(Me.txtKikanNen.Text & "/" & Me.txtKikanTsuki.Text & "/" & "01") = False Then
            MsgBox(String.Format(CommonEXT.E2026), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        '検索条件をデータクラスへ格納
        With dataEXTZ0104
            If Me.rdoTheatre.Checked Then
                'シアターをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_THEATER
            Else
                'スタジオをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_STUDIO
            End If
            .PropStrKikanNen = Me.txtKikanNen.Text
            .PropStrKikanTsuki = Integer.Parse(Me.txtKikanTsuki.Text).ToString("00")
        End With

        'スプレッドの作成
        If logicEXTZ0104.MakeSpread(dataEXTZ0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
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
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "検索対象ラジオボタンのチェック状態が変わった場合の処理 "

    ''' <summary>
    ''' 検索対象ラジオボタンのチェック状態が変わった場合の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示する一覧を切り替える 
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoTheatre_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTheatre.CheckedChanged

        If Me.rdoTheatre.Checked Then
            'シアターがチェック状態
            Me.vwRiyoJokyoThaetre.Visible = True
            Me.vwRiyoJokyoStudio.Visible = False
        Else
            'スタジオがチェック状態
            Me.vwRiyoJokyoStudio.Visible = True
            Me.vwRiyoJokyoThaetre.Visible = False
        End If

    End Sub

#End Region

End Class
