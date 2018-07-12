Imports Common
Imports CommonEXT

''' <summary>
''' EXシアターユーザーマスタメンテ画面Interfaceクラス
''' </summary>
''' <remarks>EXシアターユーザーマスタメンテ画面の設定を行う
''' <para>作成情報：2015/08/11 hayabuchi
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTM0101

    Public logicEXTM0101 As New LogicEXTM0101    'ロジッククラス
    Public dataEXTM0101 As New DataEXTM0101      'データクラス
    Public commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    ''' <summary>
    ''' フォーム読み込み時の処理
    ''' </summary>
    ''' <remarks>フォームを読み込んだ際に行われる処理
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' </para></remarks>
    Public Sub EXTM0101_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'マウスポインタ設定（通常⇒砂時計）
            Me.Cursor = Cursors.WaitCursor

            'データをセット
            With dataEXTM0101
                .PropTxtSearchUserID = Me.txtUserId                    '画面.ユーザーID
                .PropTxtSearchKanjiName = Me.txtUserKanjiName          '画面.漢字氏名
                .PropTxtSearchMail = Me.txtMail                        '画面.メールアドレス
                .PropCmbSearchBushoName = Me.cmbBushoName              '画面.部署名
                .PropChkShoninFlg = Me.chkShoninFlg                    '画面.承認者のみ表示
                .PropChkMstFlg = Me.chkMstFlg                          '画面.マスタ操作可能者のみ表示
                .PropChkStsFlg = Me.chkStsFlg                          '画面.無効なユーザーも含めて表示
                .PropBtnSearch = Me.BtnSearch                          '画面.検索ボタン
                .PropVwList = Me.ppVwList                              '画面.一覧シート
                .PropBtnReg = Me.BtnReg                                'フッタ.登録ボタン
                .PropBtnBack = Me.BtnBack                              'フッタ.戻るボタン
            End With

            'フォームの初期化
            If logicEXTM0101.InitForm(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '部署名情報取得
            If logicEXTM0101.GetComboBusho(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '部署名情報をフォームオブジェクトのコンボボックスに設定
            If logicEXTM0101.ComboBushoSet(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            'ユーザー取得
            If logicEXTM0101.InitSearch(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '画面表示
            If logicEXTM0101.SetSheet(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            ' 背景色設定
            Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
            If CommonEXT.PropConfigrationFlg = "1" Then
                ' 検証機の場合には背景イメージも表示しない
                Me.BackgroundImage = Nothing
            End If

        Catch ex As Exception
            'ログ出力設定
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

        Finally
            'マウスポインタ設定（砂時計⇒通常）
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    ''' <summary>
    ''' 検索ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>検索ボタン押下処理
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click

        Try
            'マウスポインタ設定（通常⇒砂時計）
            Me.Cursor = Cursors.WaitCursor

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'ユーザー取得
            If logicEXTM0101.Search(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '画面表示
            If logicEXTM0101.SetSheet(dataEXTM0101) = False Then
                'エラーメッセージ表示
                MsgBox(M0101_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

        Catch ex As Exception
            'ログ出力設定
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

        Finally
            'マウスポインタ設定（砂時計⇒通常）
            Me.Cursor = Cursors.Default

        End Try
    End Sub
    ''' <summary>
    ''' 登録ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>登録ボタン押下時処理
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Sub BtnReg_Click(sender As Object, e As EventArgs) Handles BtnReg.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        If logicEXTM0101.CheckInputSheet(dataEXTM0101) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If

        '確認表示
        If MsgBox(puErrMsg, MsgBoxStyle.YesNo, TITLE_INFO) = MsgBoxResult.No Then
            Exit Sub
        End If
        '重複チェック
        If logicEXTM0101.SelectCountSameUserData(dataEXTM0101) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If
        '登録処理
        If logicEXTM0101.RegUsermstData(dataEXTM0101) = False Then
            'エラーメッセージ表示
            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If
        '完了メッセージの設定
        MsgBox(puErrMsg, MsgBoxStyle.Information, TITLE_INFO)
        '画面表示
        If logicEXTM0101.InitForm(dataEXTM0101) = False Then
            Return
        End If
        If logicEXTM0101.InitSearch(dataEXTM0101) = False Then
            Return
        End If
        If logicEXTM0101.SetSheet(dataEXTM0101) = False Then
            Return
        End If
    End Sub
    ''' <summary>
    ''' 戻るボタン押下時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>前画面へ遷移する
    ''' <para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click

        Me.Close()

    End Sub

End Class