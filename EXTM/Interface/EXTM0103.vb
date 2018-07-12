Imports Common
Imports CommonEXT

''' <summary>
''' EXシアター消費税マスタメンテInterfaceクラス
''' </summary>
''' <remarks>EXシアター消費税マスタメンテの設定を行う
''' <para>作成情報：2015/08/26 hayabuchi
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTM0103

    Public logicEXTM0103 As New LogicEXTM0103    'ロジッククラス
    Public dataEXTM0103 As New DataEXTM0103      'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    ''' <summary>
    ''' フォーム読み込み時の処理
    ''' </summary>
    ''' <remarks>フォームを読み込んだ際に行われる処理
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' </para></remarks>
    Public Sub EXTM0103_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'データをセット
            With dataEXTM0103
                .PropVwList = Me.ppVwList                              '画面.一覧シート
                .PropBtnReg = Me.BtnReg                                'フッタ.登録ボタン
                .PropBtnBack = Me.BtnBack                              'フッタ.戻るボタン
            End With

            'データ取得
            If logicEXTM0103.InitSearch(dataEXTM0103) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '画面表示
            If logicEXTM0103.SetSheet(dataEXTM0103) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            ' 背景色設定
            Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
            If CommonEXT.PropConfigrationFlg = "1" Then
                ' 検証機の場合には背景イメージも表示しない
                Me.BackgroundImage = Nothing
            End If

        Catch ex As Exception

            'ログ出力設定
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

        End Try

    End Sub
    ''' <summary>
    ''' 登録ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>登録ボタン押下時処理
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub BtnReg_Click(sender As Object, e As EventArgs) Handles BtnReg.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        If logicEXTM0103.CheckInputSheet(dataEXTM0103) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If
        '確認表示
        If MsgBox(puErrMsg, MsgBoxStyle.YesNo, TITLE_INFO) = MsgBoxResult.No Then
            Exit Sub
        End If
        '登録処理
        If logicEXTM0103.ChkTaxmstData(dataEXTM0103) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If
        '完了メッセージの設定
        MsgBox(puErrMsg, MsgBoxStyle.Information, TITLE_INFO)
        '画面表示
        If logicEXTM0103.InitSearch(dataEXTM0103) = False Then
            Return
        End If
        If logicEXTM0103.SetSheet(dataEXTM0103) = False Then
            Return
        End If

    End Sub
    ''' <summary>
    ''' 戻るボタン押下時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>前画面へ遷移する
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click

        Me.Close()

    End Sub


End Class