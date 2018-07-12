Imports Common
Imports CommonEXT
Imports Npgsql
Imports FarPoint.Win.Spread


''' <summary>利用者情報登録/詳細Interfaceクラス
''' </summary>
''' <remarks>利用者情報登録/詳細画面の設定を行う
''' <para>作成情報：2015/08/25 ozawa
''' <p>改訂情報：</p>
''' </para>
''' </remarks>
Public Class EXTM0202
    'インスタンス生成
    Public DataEXTM0201 As New DataEXTM0201         '利用者一覧画面データクラス
    Public dataEXTM0202 As New DataEXTM0202         '利用者情報/詳細画面データクラス
    Public dataEXTM0203 As New DataEXTM0203         'EXAS相手先一覧画面データクラス
    'Private EXTM0203 As New EXTM0203
    Private LogicEXTM0202 As New LogicEXTM0202      'Logicクラス
    Private CommonLogic As New CommonLogic          '共通ロジッククラス
    Private SqlEXTN0201 As New SqlEXTM0201          'Sqlクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    ''' <summary>初期表示時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>読み込み時に行われる処理
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報</p>
    ''' </para>
    ''' </remarks>
    Private Sub EXTM0202_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With dataEXTM0202
            Me.Lbl_RiyoCd.Text = dataEXTM0202.PropParamRiyoCd            '引渡しデータをセット

            .PropTxtRiyo_nm = Me.Txt_RiyoNm         '利用者名
            .PropTxtRiyo_kana = Me.Txt_RiyoKana     '利用者カナ
            .PropTxtDaihyo_nm = Me.Txt_Daihyo       '代表者名
            .PropTxtTel1 = Me.Txt_Tel1              '電話番号1
            .PropTxtTel2 = Me.Txt_Tel2              '電話番号2
            .PropTxtTel3 = Me.Txt_Tel3              '電話番号3
            .PropTxtNaisen = Me.Txt_Naisen          '内線番号
            .PropTxtFax1 = Me.Txt_Fax1              'Fax番号1
            .PropTxtFax2 = Me.Txt_Fax2              'Fax番号2
            .PropTxtFax3 = Me.Txt_Fax3              'Fax番号3
            .PropTxtYubin1 = Me.Txt_Yubin1          '郵便番号1
            .PropTxtYubin2 = Me.Txt_Yubin2          '郵便番号2
            .PropCmbTodo = Me.Cmb_Todo              '都道府県
            .PropTxtShiku = Me.Txt_Shiku            '市区町村
            .PropTxtBanchi = Me.Txt_Banchi          '番地
            .PropTxtBuild = Me.Txt_Build            'ビル名
            .PropTxtCom = Me.Txt_Com                'コメント
            .PropRdoTujyo = Me.Rdo_Tujyo            'ラジオボタン（通常）
            .PropRdoChui = Me.Rdo_Chui              'ラジオボタン（要注意）
            .PropRdoHuka = Me.Rdo_Huka              'ラジオボタン（利用不可）
            .PropLblAite_cd = Me.Lbl_AiteCd         '相手先コード
            .PropLblAite_nm = Me.Lbl_AiteNm         '相手先名

            .PropVwList = Me.FpSpread1                    '一覧シート

        End With

        'コンボボックス編集
        If CommonLogic.SetCmbBox(strCmbToDo, Cmb_Todo) = False Then
            Return
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 初期表示時の検索処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>条件に応じて初期検索を行う
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報</p>
    ''' </para></remarks>
    Private Sub EXTM0202_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown

        '遷移元画面から利用者番号がパラメータとして渡ってきている場合のみ行う。
        Try

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'データをセット
            With dataEXTM0202
                .PropVwList = FpSpread1              '一覧シート
                .PropLblRiyo_cd = Lbl_RiyoCd         '表示されている利用者番号
            End With

            'パラメータが渡ってきていない場合は検索しない
            If dataEXTM0202.PropParamRiyoCd = "" Then
                '新規用の引き渡し項目を編集
                With dataEXTM0202
                    .PropTxtRiyo_nm.Text = .PropParamRiyoNm            ' 遷移元画面から引き継がれる利用者名
                    .PropTxtRiyo_kana.Text = .PropParamRiyoKana        ' 遷移元画面から引き継がれる利用者名カナ
                    .PropTxtDaihyo_nm.Text = .PropParamDaihyoNm        ' 遷移元画面から引き継がれる代表者名
                    .PropTxtTel1.Text = .PropParamRiyoTel11            ' 遷移元画面から引き継がれる利用者電話番号1
                    .PropTxtTel2.Text = .PropParamRiyoTel12            ' 遷移元画面から引き継がれる利用者電話番号2
                    .PropTxtTel3.Text = .PropParamRiyoTel13            ' 遷移元画面から引き継がれる利用者電話番号3
                    .PropTxtNaisen.Text = .PropParamRiyoNaisen         ' 遷移元画面から引き継がれる利用者内線番号
                    .PropTxtFax1.Text = .PropParamRiyoFax11            ' 遷移元画面から引き継がれる利用者FAX1
                    .PropTxtFax2.Text = .PropParamRiyoFax12            ' 遷移元画面から引き継がれる利用者FAX2
                    .PropTxtFax3.Text = .PropParamRiyoFax13            ' 遷移元画面から引き継がれる利用者FAX3
                    .PropTxtYubin1.Text = .PropParamRiyoYubin1         ' 遷移元画面から引き継がれる利用者郵便番号1
                    .PropTxtYubin2.Text = .PropParamRiyoYubin2         ' 遷移元画面から引き継がれる利用者郵便番号2
                    .PropCmbTodo.Text = .PropParamRiyoTodo             ' 遷移元画面から引き継がれる利用者都道府県
                    .PropTxtShiku.Text = .PropParamRiyoShiku           ' 遷移元画面から引き継がれる利用者市区町村
                    .PropTxtBanchi.Text = .PropParamRiyoBan            ' 遷移元画面から引き継がれる利用者番地
                    .PropTxtBuild.Text = .PropParamRiyoBuild           ' 遷移元画面から引き継がれる利用者ビル名
                End With
                Exit Sub
            End If

            '待機カーソル
            Me.Cursor = Cursors.WaitCursor

            ' 利用者情報の取得
            If LogicEXTM0202.GetRiyoshaData(dataEXTM0202) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '利用コメントの取得
            If LogicEXTM0202.GetComment(dataEXTM0202) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            'スプレッドのプロパティー設定
            SpreadConfig()

        Catch ex As Exception
            Common.CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, ex.Message, ex, Nothing)
            'エラーメッセージ表示()
            MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    ''' <summary>更新ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>更新ボタン押下時、利用者情報の更新/新規登録を行う
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub Btn_Update_Click(sender As Object, e As EventArgs) Handles Btn_Update.Click
        'マウスポインタ変更(通常→砂時計)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If dataEXTM0202.PropLblRiyo_cd.Text <> "" Then
            '利用者番号がある場合：更新

            '入力チェックを行う
            If LogicEXTM0202.CheckInputValueMain(dataEXTM0202) = False Then
                'MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If
            '確認ダイアログ表示
            If MessageBox.Show(String.Format(EXTM0102_C0011, "利用者情報"), "登録確認", MessageBoxButtons.YesNo) = DialogResult.No Then
                Exit Sub
            End If
            '更新処理
            If LogicEXTM0202.UpdatetNewData(dataEXTM0202) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If
            '完了メッセージ
            MsgBox(String.Format(M0202_I0002, "利用者情報の登録処理"))

        ElseIf dataEXTM0202.PropLblRiyo_cd.Text = "" Then
            '利用者番号がない場合：新規登録

            '入力チェックを行う
            If LogicEXTM0202.CheckInputValueMain(dataEXTM0202) = False Then
                'MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If
            '確認ダイアログ表示 
            If MessageBox.Show(String.Format(EXTM0102_C0011, "利用者情報"), "登録確認", MessageBoxButtons.YesNo) = DialogResult.No Then
                Exit Sub
            End If
            '更新処理
            If LogicEXTM0202.InsertNewData(dataEXTM0202) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '採番した利用者コードを表示
            Me.Lbl_RiyoCd.Text = dataEXTM0202.PropNewRiyo_cd
            '完了メッセージ
            MsgBox(String.Format(M0202_I0002, "利用者情報の登録処理"))

        End If
    End Sub


    ''' <summary>EXAS相手先マスタ検索ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>EXAS相手先マスタ検索ボタン押下時、利用者とEXAS相手先マスタの紐付けを行う
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub Btn_AitesakiSearch_Click(sender As Object, e As EventArgs) Handles Btn_AitesakiSearch.Click
        '戻り値を格納している配列
        Dim ArrayList As New ArrayList

        'データを引き渡す
        With EXTM0203.dataEXTM0203
            .PropStrMenuID = "EXTM0202"
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '画面遷移
        ArrayList = EXTM0203.ShowDialog()

        If ArrayList Is Nothing Then
            dataEXTM0202.PropLblAite_cd = Lbl_AiteCd          '相手先コード
            dataEXTM0202.PropLblAite_nm = Lbl_AiteNm          '相手先名
        Else
            If ArrayList.Count <> 0 Then
                Me.Lbl_AiteCd.Text = ArrayList(0)
                Me.Lbl_AiteNm.Text = ArrayList(1)

            Else
                dataEXTM0202.PropLblAite_cd = Lbl_AiteCd          '相手先コード
                dataEXTM0202.PropLblAite_nm = Lbl_AiteNm          '相手先名
            End If
        End If
        'If ArrayList.Count <> 0 Then
        '    Me.Lbl_AiteCd.Text = ArrayList(0)
        '    Me.Lbl_AiteNm.Text = ArrayList(1)

        'Else
        '    dataEXTM0202.PropLblAite_cd = Lbl_AiteCd          '相手先コード
        '    dataEXTM0202.PropLblAite_nm = Lbl_AiteNm          '相手先名

        'End If

        Me.Show()
    End Sub

    ''' <summary>戻るボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>戻るボタン押下時、遷移元によって処理を切り替える
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub Btn_Modoru_Click(sender As Object, e As EventArgs) Handles Btn_Modoru.Click
        '遷移元画面によって処理を切り替える(単体テスト時はなし)
        'メインウインドウで開いている場合
        '利用者ＩＤが設定されている場合OK戻り、未設定はCanselで戻り
        If dataEXTM0202.PropLblRiyo_cd.Text = "" Then
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End If
        'モーダルウインドウで開いている場合、ウインドウを閉じる
        Me.Close()

    End Sub


    ''' <summary>画面プロパティ関数
    ''' </summary>
    ''' <remarks>画面のプロパティーを再設定する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub SpreadConfig()
        'ヘッダーの幅を再セット
        FpSpread1.ActiveSheet.Columns(0).Width = 93
        FpSpread1.ActiveSheet.Columns(1).Width = 93
        FpSpread1.ActiveSheet.Columns(2).Width = 239
        FpSpread1.ActiveSheet.Columns(3).Width = 389
    End Sub

End Class