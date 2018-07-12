Imports Common
Imports CommonEXT
Imports Npgsql
Imports FarPoint.Win.Spread
'Imports EXTA

'''<summary>利用者一覧画面Interfaceクラス画面
''' </summary>
''' <remarks>利用者一覧画面の設定を行う
''' <para>作成情報：2015/08/10 ozawa
''' <p>改訂情報：</p>
''' </para></remarks>

Public Class EXTM0201
    'インスタンス生成
    Public DataEXTM0201 As New DataEXTM0201         'データクラス
    Private LogicEXTM0201 As New LogicEXTM0201      'Logicクラス
    Private CommonLogic As New CommonLogic          '共通ロジッククラス
    Private SqlEXTN0201 As New SqlEXTM0201          'Sqlクラス
    Private commonLogicEXT As New CommonLogicEXT


    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>読み込み時に行われる処理
    ''' <para>作成情報：2015/08/10
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub EXTM0201_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If DataEXTM0201.PropParamValue <> "" And DataEXTM0201.PropTxtRiyo_kana IsNot Nothing Then
            Txt_RiyoKana.Text = DataEXTM0201.PropTxtRiyo_kana.Text
        End If

        With DataEXTM0201
            .PropVwList = FpSpread1       '一覧シート
            .PropRdoTujyo = Rdo_Tujyo     'ラジオボタン（通常）
            .PropRdoChui = Rdo_chui       'ラジオボタン（要注意）
            .PropRdoHuka = Rdo_Huka       'ラジオボタン（利用不可）

            .PropTxtRiyo_cd = Txt_RiyoCd         '利用者番号
            .PropTxtRiyo_nm = Txt_RiyoNm         '利用者名
            .PropTxtRiyo_kana = Txt_RiyoKana     '利用者カナ
            .PropTxtAite_cd = Txt_AitesakiCd     '相手先コード
            .PropTxtAite_nm = Txt_AitesakiNm     '相手先名
            .PropTxtTel1 = Txt_Tel1              '電話番号①
            .PropTxtTel2 = Txt_Tel2              '電話番号②
            .PropTxtTel3 = Txt_Tel3              '電話番号③
        End With


        ' スプレッドの初期化処理
        If LogicEXTM0201.InitView(DataEXTM0201) = False Then
            'エラーメッセージ表示
            MsgBox(M0201_E0000, MsgBoxStyle.Critical, TITLE_ERROR)
            Return
        End If

        ' 2015.12.03 ADD h.hagiwara 検索表示フラグ   0:表示なし
        DataEXTM0201.PropInitDspFlg = "0"

        'メニュー画面からの遷移時、選択確定ボタンを非活性にする
        'フラグ(0:通常モード　1:要注意者・利用停止者を初期表示　それ以外:メインウインドウで表示)

        'DataEXTM0201.PropParamValue = ""      '単体テストのために記述

        '0:通常レベルの利用者を初期表示する場合
        If DataEXTM0201.PropParamValue = "0" Then
            Btn_Kakutei.Enabled = True   '選択確定ボタンを活性化する
            Btn_Touroku.Visible = False  '新規登録ボタンを非表示にする
            DataEXTM0201.PropVwList.Sheets(0).Columns(13).Visible = False    '編集ボタンの列を非表示にする

            '1:要注意者・利用停止者を初期表示する場合
        ElseIf DataEXTM0201.PropParamValue = "1" Then
            'Btn_Kakutei.Enabled = True    '選択確定ボタンを活性化する
            Btn_Kakutei.Enabled = False    '選択確定ボタンを活性化する
            Btn_Touroku.Visible = False    '新規登録ボタンを非表示にする
            Rdo_chui.Checked = True        'ラジオボタン：要注意にデフォルトチェック
            DataEXTM0201.PropVwList.Sheets(0).Columns(13).Visible = False   '編集ボタンの列を非表示にする
            DataEXTM0201.PropVwList.Sheets(0).Columns(0).Locked = True   '選択の入力を不可とする

            'それ以外（メニュー画面から遷移してきた場合）
        Else
            Btn_Kakutei.Enabled = False   '選択確定ボタンを非活性にする
            DataEXTM0201.PropVwList.Sheets(0).Columns(0).Locked = True   '選択の入力を不可とする
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 初期検索を行う
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報</p>
    ''' </para></remarks>
    Private Sub EXTM0201_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Try

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'データをセット
            With DataEXTM0201
                .PropVwList = FpSpread1       '一覧シート
            End With

            'パラメータ値が設定されていない場合は初期検索しない
            If DataEXTM0201.PropParamValue <> "0" And DataEXTM0201.PropParamValue <> "1" Then
                Exit Sub
            End If

            '待機カーソル
            Me.Cursor = Cursors.WaitCursor

            ' 一覧の取得
            If LogicEXTM0201.InitSearch(DataEXTM0201) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            ' 一覧の表示処理
            If LogicEXTM0201.ViewRow(DataEXTM0201) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '画面プロパティー設定
            SpreadConfig()

            ' 0件の場合（暫定）
            If DataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then

            End If

        Catch ex As Exception
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'エラーメッセージ表示
            MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>検索ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>検索ボタン押下時、SQLを発行し結果を表示する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>

    Private Sub Btn_Kensaku_Click(sender As Object, e As EventArgs) Handles Btn_Kensaku.Click

        'SQL発行
        Try

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '引数の設定
            With DataEXTM0201
                .PropVwList = FpSpread1       '一覧シート
                .PropRdoTujyo = Rdo_Tujyo     'ラジオボタン（通常）
                .PropRdoChui = Rdo_chui       'ラジオボタン（要注意）
                .PropRdoHuka = Rdo_Huka       'ラジオボタン（利用不可）

                .PropTxtRiyo_cd = Txt_RiyoCd         '利用者番号
                .PropTxtRiyo_nm = Txt_RiyoNm         '利用者名
                .PropTxtRiyo_kana = Txt_RiyoKana     '利用者カナ
                .PropTxtAite_cd = Txt_AitesakiCd     '相手先コード
                .PropTxtAite_nm = Txt_AitesakiNm     '相手先名
                .PropTxtTel1 = Txt_Tel1              '電話番号①
                .PropTxtTel2 = Txt_Tel2              '電話番号②
                .PropTxtTel3 = Txt_Tel3              '電話番号③

            End With

            '待機カーソル
            Me.Cursor = Cursors.WaitCursor

            'DBから取得
            If LogicEXTM0201.Search(DataEXTM0201) = False Then
                'エラーメッセージを表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            ' 一覧の表示処理
            If LogicEXTM0201.ViewRow(DataEXTM0201) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return
            End If

            '画面プロパティー設定
            SpreadConfig()

            ' 0件の場合（暫定）
            If DataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then

            End If

            ' 2015.12.03 ADD h.hagiwara 検索表示フラグ   1:表示あり
            DataEXTM0201.PropInitDspFlg = "1"

        Catch ex As Exception
            Common.CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, ex.Message, ex, Nothing)
            'エラーメッセージ表示()
            MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    ''' <summary>新規登録ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>新規登録ボタン押下時、画面遷移する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>

    Private Sub Btn_Touroku_Click(sender As Object, e As EventArgs) Handles Btn_Touroku.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0202

        Me.Hide()

        frm.ShowDialog()

        'Me.Close()

        Me.Show()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        If DataEXTM0201.PropInitDspFlg = "1" Then                                     ' 2015.12.03 ADD h.hagiwara
            Try
                '引数の設定
                With DataEXTM0201
                    .PropVwList = FpSpread1       '一覧シート
                    .PropRdoTujyo = Rdo_Tujyo     'ラジオボタン（通常）
                    .PropRdoChui = Rdo_chui       'ラジオボタン（要注意）
                    .PropRdoHuka = Rdo_Huka       'ラジオボタン（利用不可）

                    .PropTxtRiyo_cd = Txt_RiyoCd         '利用者番号
                    .PropTxtRiyo_nm = Txt_RiyoNm         '利用者名
                    .PropTxtRiyo_kana = Txt_RiyoKana     '利用者カナ
                    .PropTxtAite_cd = Txt_AitesakiCd     '相手先コード
                    .PropTxtAite_nm = Txt_AitesakiNm     '相手先名
                    .PropTxtTel1 = Txt_Tel1              '電話番号①
                    .PropTxtTel2 = Txt_Tel2              '電話番号②
                    .PropTxtTel3 = Txt_Tel3              '電話番号③

                End With

                '待機カーソル
                Me.Cursor = Cursors.WaitCursor

                'DBから取得
                If LogicEXTM0201.Search(DataEXTM0201) = False Then
                    'エラーメッセージを表示
                    MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                    Return
                End If

                ' 一覧の表示処理
                If LogicEXTM0201.ViewRow(DataEXTM0201) = False Then
                    'エラーメッセージ表示
                    MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                    Return
                End If

                '画面プロパティー設定
                SpreadConfig()

                ' 0件の場合（暫定）
                If DataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then

                End If

            Catch ex As Exception
                Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
                'エラーメッセージ表示
                MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
            Finally
                Me.Cursor = Cursors.Default
            End Try
        End If                                                                                     ' 2015.12.03 ADD h.hagiwara

    End Sub


    ''' <summary>確認・編集ボタン ／スプレッドシートクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>確認・編集ボタン押下時、利用者番号を詳細画面に引き渡す／チェックボックスの制御を行う
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub FpSpread1_ButtonClicked(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles FpSpread1.ButtonClicked

        '確認・編集ボタンの列
        DataEXTM0201.PropIndex = 13

        '確認・編集ボタンの列(13)がクリックされた場合
        If e.Column = DataEXTM0201.PropIndex Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim EXTM0202 As New EXTM0202
            Dim DataEXTM0202 As New DataEXTM0202

            '引渡すデータ（利用者番号）を格納
            With EXTM0202.dataEXTM0202
                .PropParamRiyoCd = FpSpread1.ActiveSheet.GetValue(e.Row, 1)
            End With

            '画面遷移
            Me.Hide()
            EXTM0202.ShowDialog()
            'Me.Close()
            Me.Show()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            If DataEXTM0201.PropInitDspFlg = "1" Then                                     ' 2015.12.03 ADD h.hagiwara
                Try

                    '引数の設定
                    With DataEXTM0201
                        .PropVwList = FpSpread1       '一覧シート
                        .PropRdoTujyo = Rdo_Tujyo     'ラジオボタン（通常）
                        .PropRdoChui = Rdo_chui       'ラジオボタン（要注意）
                        .PropRdoHuka = Rdo_Huka       'ラジオボタン（利用不可）

                        .PropTxtRiyo_cd = Txt_RiyoCd         '利用者番号
                        .PropTxtRiyo_nm = Txt_RiyoNm         '利用者名
                        .PropTxtRiyo_kana = Txt_RiyoKana     '利用者カナ
                        .PropTxtAite_cd = Txt_AitesakiCd     '相手先コード
                        .PropTxtAite_nm = Txt_AitesakiNm     '相手先名
                        .PropTxtTel1 = Txt_Tel1              '電話番号①
                        .PropTxtTel2 = Txt_Tel2              '電話番号②
                        .PropTxtTel3 = Txt_Tel3              '電話番号③

                    End With

                    '待機カーソル
                    Me.Cursor = Cursors.WaitCursor

                    'DBから取得
                    If LogicEXTM0201.Search(DataEXTM0201) = False Then
                        'エラーメッセージを表示
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return
                    End If

                    ' 一覧の表示処理
                    If LogicEXTM0201.ViewRow(DataEXTM0201) = False Then
                        'エラーメッセージ表示
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return
                    End If

                    '画面プロパティー設定
                    SpreadConfig()

                    ' 0件の場合（暫定）
                    If DataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then

                    End If

                Catch ex As Exception
                    Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
                    'エラーメッセージ表示
                    MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical, TITLE_ERROR)
                Finally
                    Me.Cursor = Cursors.Default
                End Try

            End If                                                                                     ' 2015.12.03 ADD h.hagiwara

        ElseIf e.Column = 0 Then  '選択行をクリックした場合

            'イベントを発生させないようにする
            RemoveHandler FpSpread1.ButtonClicked, AddressOf FpSpread1_ButtonClicked

            '選択行番号取得
            DataEXTM0201.PropCheckIndex = e.Row

            '選択行にチェックをつけ、それ以外はチェックを外す
            If LogicEXTM0201.ClickVwCellMain(DataEXTM0201) = False Then
                'エラーメッセージ表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Exit Sub
            End If

            'イベントを発生させるようにする
            AddHandler FpSpread1.ButtonClicked, AddressOf FpSpread1_ButtonClicked

        End If
    End Sub

    ''' <summary>選択確定ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>選択確定ボタン押下時、利用者情報を引渡し呼び出し元画面へ遷移する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub Btn_Kakutei_Click(sender As Object, e As EventArgs) Handles Btn_Kakutei.Click
        '変数宣言
        Dim EXTA0101 As New EXTM0101

        Try

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'データをセット
            With DataEXTM0201
                Me.vwList = .PropVwList
            End With

            ' チェックされた行のインデックス取得
            Dim Index As Integer() = GetCheckRowIndex(vwList)

            ' 選択されていない場合
            If Index.Length = 0 Then
                'エラーメッセージ表示
                puErrMsg = String.Format(M0201_E0002, "利用者情報")
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return

                '複数行選択している場合
            ElseIf Index.Length > 1 Then
                'エラーメッセージ表示
                puErrMsg = String.Format(M0201_E2035, "利用者情報")
                MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                Return

            ElseIf Index.Length = 1 Then

                ' 戻り値設定
                With DataEXTM0201
                    .PropParamRiyoCd = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_CD).Value              ' 遷移元画面に引き渡す利用者番号
                    .PropParamRiyoNm = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_NM).Value              ' 遷移元画面に引き渡す利用者名
                    .PropParamRiyoKana = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_KNM).Value           ' 遷移元画面に引き渡す利用者名カナ
                    .PropParamDaihyoNm = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SDAIHYO_NM).Value    ' 遷移元画面に引き渡す代表者名
                    .PropParamRiyoTel11 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_STEL11).Value       ' 遷移元画面に引き渡す利用者電話番号1
                    .PropParamRiyoTel12 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_STEL12).Value       ' 遷移元画面に引き渡す利用者電話番号2
                    .PropParamRiyoTel13 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_STEL13).Value       ' 遷移元画面に引き渡す利用者電話番号3
                    .PropParamRiyoNaisen = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SNAISEN).Value     ' 遷移元画面に引き渡す利用者内線番号
                    .PropParamRiyoFax11 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SFAX11).Value       ' 遷移元画面に引き渡す利用者FAX1
                    .PropParamRiyoFax12 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SFAX12).Value       ' 遷移元画面に引き渡す利用者FAX2
                    .PropParamRiyoFax13 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SFAX13).Value       ' 遷移元画面に引き渡す利用者FAX3
                    .PropParamRiyoYubin1 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SYUBIN1).Value     ' 遷移元画面に引き渡す利用者郵便番号1
                    .PropParamRiyoYubin2 = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SYUBIN2).Value     ' 遷移元画面に引き渡す利用者郵便番号2
                    .PropParamRiyoTodo = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_ADD1).Value          ' 遷移元画面に引き渡す利用者都道府県
                    .PropParamRiyoShiku = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_ADD2).Value         ' 遷移元画面に引き渡す利用者市区町村
                    .PropParamRiyoBan = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_ADD3).Value           ' 遷移元画面に引き渡す利用者番地
                    .PropParamRiyoBuild = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_ADD4).Value         ' 遷移元画面に引き渡す利用者ビル名
                    .PropParamRiyoLvl = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_SLVL).Value           ' 遷移元画面に引き渡す利用者レベル
                    .PropParamAiteCd = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_AITE_CD).Value         ' 遷移元画面に引き渡す相手先コード
                    .PropParamAiteNm = .PropVwList.ActiveSheet.Cells(Index(0), M0201_COL_RIYO_AITE_NM).Value         ' 遷移元画面に引き渡す相手先名

                End With

                '戻り値をOKにする
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                'フォームを閉じる
                Me.Close()

            End If

        Catch ex As Exception
            'エラーメッセージ表示
            MsgBox(M0201_E0000 & ex.Message, MsgBoxStyle.Critical)
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
        End Try

    End Sub

    ''' <summary>戻るボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>戻るボタン押下時
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub Btn_Modoru_Click(sender As Object, e As EventArgs) Handles Btn_Modoru.Click
        'モーダルウインドウで開いていた場合
        If DataEXTM0201.PropParamValue = "0" And DataEXTM0201.PropParamValue = "1" Then
            '戻り値をキャンセルにする
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End If

        'フォームを閉じる
        Me.Close()
    End Sub


    ''' <summary>画面プロパティ関数
    ''' </summary>
    ''' <remarks>画面のプロパティーを再設定する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub SpreadConfig()
        'ヘッダーの幅を再セット
        FpSpread1.ActiveSheet.Columns(0).Width = 34
        FpSpread1.ActiveSheet.Columns(1).Width = 81
        FpSpread1.ActiveSheet.Columns(2).Width = 69
        FpSpread1.ActiveSheet.Columns(3).Width = 165
        FpSpread1.ActiveSheet.Columns(4).Width = 130
        FpSpread1.ActiveSheet.Columns(5).Width = 93
        FpSpread1.ActiveSheet.Columns(6).Width = 82
        FpSpread1.ActiveSheet.Columns(7).Width = 172
        FpSpread1.ActiveSheet.Columns(8).Width = 86
        FpSpread1.ActiveSheet.Columns(9).Width = 86
        FpSpread1.ActiveSheet.Columns(10).Width = 86
        FpSpread1.ActiveSheet.Columns(11).Width = 77
        FpSpread1.ActiveSheet.Columns(12).Width = 278
        FpSpread1.ActiveSheet.Columns(13).Width = 75

    End Sub

    ''' <summary>チェックされた行インデックスを取得する
    ''' </summary>
    ''' <param name="spread"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Function GetCheckRowIndex(ByVal spread As FarPoint.Win.Spread.FpSpread) As Integer()
        Dim indexList As New List(Of Integer)

        For i As Integer = 0 To spread.ActiveSheet.RowCount - 1
            If spread.ActiveSheet.GetValue(i, 0) = True Then
                indexList.Add(i)
            End If
        Next

        Return indexList.ToArray()

    End Function

End Class