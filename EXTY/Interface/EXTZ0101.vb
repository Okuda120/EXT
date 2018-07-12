Imports Common
Imports CommonEXT
Imports EXTB
Imports EXTC
Imports EXTM
Imports EXTY

''' <summary>
''' EXTZ0101
''' </summary>
''' <remarks>予約一覧画面
''' <para>作成情報：2015/09/09 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
''' 
Public Class EXTZ0101

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTZ0101 As New LogicEXTZ0101     'ロジッククラス
    Public dataEXTZ0101 As New DataEXTZ0101         'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「予約一覧画面」ロード時処理 "

    ''' <summary>
    ''' 「予約一覧画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「予約一覧画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0101_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '画面初期設定
        Me.rdoTheatre.Checked = True
        Me.dtpRiyoDt_From.txtDate.Text = String.Empty
        Me.dtpRiyoDt_To.txtDate.Text = String.Empty
        Me.txtRiyoNm.Text = String.Empty
        Me.txtRiyoNm_Kana.Text = String.Empty
        Me.txtSaijiShutsuenNm.Text = String.Empty
        Me.txtYoyakuNo.Text = String.Empty
        Me.chkMikanryoOnly.Checked = False
        If Me.vwYoyakuTheatre.ActiveSheet.Rows.Count > 0 Then
            Me.vwYoyakuTheatre.Sheets(0).RemoveRows(0, Me.vwYoyakuTheatre.ActiveSheet.Rows.Count)
        End If
        If Me.vwYoyakuStudio.ActiveSheet.Rows.Count > 0 Then
            Me.vwYoyakuStudio.Sheets(0).RemoveRows(0, Me.vwYoyakuStudio.ActiveSheet.Rows.Count)
        End If
        Me.vwYoyakuTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwYoyakuTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwYoyakuStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwYoyakuStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwYoyakuStudio.Visible = False
        Me.btnDecision.Enabled = False

        'プロパティに設定
        dataEXTZ0101.PropvwYoyakuTheatre = Me.vwYoyakuTheatre
        dataEXTZ0101.PropvwYoyakuStudio = Me.vwYoyakuStudio
        dataEXTZ0101.PropBtnDecision = Me.btnDecision
        dataEXTZ0101.PropFrmYoyaku = Me

        '選択確定ボタンの制御
        If dataEXTZ0101.PropStrSelectMode = SELECTMODE_ALSOKNYUKIN Then
            '「ALSOK現金入金機データ登録」画面から遷移した場合
            dataEXTZ0101.PropBtnDecision.Enabled = True
            '選択確定ボタンを活性
            dataEXTZ0101.PropvwYoyakuTheatre.Sheets(0).Columns(0).Locked = False    '選択の入力を可とする
            dataEXTZ0101.PropvwYoyakuStudio.Sheets(0).Columns(0).Locked = False   '選択の入力を可とする
            ' 2015.12.15 ADD START↓ h.hagiwara
            If dataEXTZ0101.PropStrSelectShisetu = "2" Then
                Me.rdoTheatre.Enabled = False
                Me.rdoStudio.Enabled = True
                Me.rdoStudio.Checked = True
                Me.vwYoyakuTheatre.Visible = False
                Me.vwYoyakuStudio.Visible = True
            Else
                Me.rdoStudio.Enabled = False
                Me.vwYoyakuTheatre.Visible = True
                Me.vwYoyakuStudio.Visible = False
            End If
            ' 2015.12.15 ADD END↑ h.hagiwara
        Else
            'メニュー画面から遷移した場合
            dataEXTZ0101.PropBtnDecision.Enabled = False
            '選択確定ボタンを非活性
            dataEXTZ0101.PropvwYoyakuTheatre.Sheets(0).Columns(0).Locked = True   '選択の入力を不可とする
            dataEXTZ0101.PropvwYoyakuStudio.Sheets(0).Columns(0).Locked = True   '選択の入力を不可とする
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

#Region "「利用者検索」ボタン押下時処理 "

    ''' <summary>
    ''' 「利用者検索」ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>利用者名カナの取得を行う</remarks>
    Private Sub btnRiyoshaSearch_Click(sender As Object, e As EventArgs) Handles btnRiyoshaSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0201

        '利用者一覧画面を表示
        ' 201.03.03 UPD START↓ h.hagiwara 
        'If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '    '利用者名カナを設定
        '    Me.txtRiyoNm_Kana.Text = frm.DataEXTM0201.PropStrRiyoKana
        'End If
        frm.DataEXTM0201.PropTxtRiyo_kana = Me.txtRiyoNm_Kana
        frm.DataEXTM0201.PropParamValue = "0"
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '利用者名カナを設定
            Me.txtRiyoNm_Kana.Text = frm.DataEXTM0201.PropParamRiyoKana
        End If
        ' 201.03.03 UPD START↓ h.hagiwara 

    End Sub

#End Region

#Region "「検索」ボタン押下時処理 "

    ''' <summary>
    ''' 「検索」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「戻る」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        '変数宣言
        Dim regexHalfAlpNum As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]+$")     '半角英数字チェックの正規表現パターン 

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        '＜半角英数字チェック＞
        '---予約番号
        If Me.txtYoyakuNo.Text <> String.Empty Then
            If regexHalfAlpNum.IsMatch(Me.txtYoyakuNo.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "予約NO"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '＜大小チェック＞
        '---利用日
        If Me.dtpRiyoDt_From.txtDate.Text <> String.Empty AndAlso Me.dtpRiyoDt_To.txtDate.Text <> String.Empty Then
            If commonValidate.IsDateFromTo(Me.dtpRiyoDt_From.txtDate.Text, Me.dtpRiyoDt_To.txtDate.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0022, "利用日（To）", "利用日（From）"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '検索条件をデータクラスへ格納
        With dataEXTZ0101
            If Me.rdoTheatre.Checked Then
                'シアターをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_THEATER
            Else
                'スタジオをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_STUDIO
            End If
            .PropStrRiyoDtFrom = Me.dtpRiyoDt_From.txtDate.Text
            .PropStrRiyoDtTo = Me.dtpRiyoDt_To.txtDate.Text
            .PropStrRiyoNm = Me.txtRiyoNm.Text
            .PropStrRiyokana = Me.txtRiyoNm_Kana.Text
            .PropStrSaijiNm = Me.txtSaijiShutsuenNm.Text
            .PropStrYoyakuNo = Me.txtYoyakuNo.Text
            If Me.chkMikanryoOnly.Checked Then
                '「未完了の予約のみを表示」をチェック
                .PropBlnMikanryo = True
            Else
                '「未完了の予約のみを表示」を未チェック
                .PropBlnMikanryo = False
            End If
        End With

        'スプレッドの作成
        If logicEXTZ0101.MakeSpread(dataEXTZ0101) = False Then
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
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "「選択確定」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択確定」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択確定」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/09/21 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnDecision_Click(sender As Object, e As EventArgs) Handles btnDecision.Click

        '変数宣言
        Dim intCheckCnt As Integer = 0                  'チェック件数
        Dim strYoyakuNo As String = String.Empty        '予約NO
        Dim strSaijiNm As String = String.Empty         '催事・アーティスト名

        If Me.rdoTheatre.Checked Then
            'シアターがチェック状態
            For intDataRow As Integer = 0 To Me.vwYoyakuTheatre.ActiveSheet.Rows.Count - 1
                If Me.vwYoyakuTheatre.ActiveSheet.Cells(intDataRow, SpreadIndex_Theatre_Check).Value = "1" Then
                    strYoyakuNo = Me.vwYoyakuTheatre.ActiveSheet.Cells(intDataRow, SpreadIndex_Theatre_YoyakuNo).Value
                    strSaijiNm = Me.vwYoyakuTheatre.ActiveSheet.Cells(intDataRow, SpreadIndex_Theatre_SaijiNm).Value
                    intCheckCnt += 1
                End If
            Next
        Else
            'スタジオがチェック状態
            For intDataRow As Integer = 0 To Me.vwYoyakuStudio.ActiveSheet.Rows.Count - 1
                If Me.vwYoyakuStudio.ActiveSheet.Cells(intDataRow, SpreadIndex_Studio_Check).Value = "1" Then
                    strYoyakuNo = Me.vwYoyakuStudio.ActiveSheet.Cells(intDataRow, SpreadIndex_Studio_YoyakuNo).Value
                    'strSaijiNm = Me.vwYoyakuTheatre.ActiveSheet.Cells(intDataRow, SpreadIndex_Studio_ShutsuenNm).Value                                 ' 2015.12.09 UPD h.hagiwara
                    strSaijiNm = Me.vwYoyakuStudio.ActiveSheet.Cells(intDataRow, SpreadIndex_Studio_ShutsuenNm).Value                                   ' 2015.12.09 UPD h.hagiwara
                    intCheckCnt += 1
                End If
            Next
        End If

        If intCheckCnt = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "予約"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        ElseIf intCheckCnt > 1 Then
            MsgBox(String.Format(CommonEXT.E2035, "予約"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        '遷移元に返却するパラメータをデータクラスに格納
        dataEXTZ0101.PropStrRtnYoyakuNo = strYoyakuNo
        dataEXTZ0101.PropStrRtnSaijiNm = strSaijiNm

        '戻り値をOKに設定
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

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
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoTheatre_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTheatre.CheckedChanged

        If Me.rdoTheatre.Checked Then
            'シアターがチェック状態
            Me.vwYoyakuTheatre.Visible = True
            Me.vwYoyakuStudio.Visible = False
        Else
            'スタジオがチェック状態
            Me.vwYoyakuStudio.Visible = True
            Me.vwYoyakuTheatre.Visible = False
        End If

    End Sub

#End Region

#Region "予約一覧（シアター）のボタンをクリックした場合"

    ''' <summary>
    ''' 予約一覧（シアター）のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>予約一覧（シアター）のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/09/10 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwYoyakuTheatre_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwYoyakuTheatre.ButtonClicked

        '「確認・編集」ボタンが押下された場合
        If e.Column = SpreadIndex_Theatre_Button Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim strStsCd = vwYoyakuTheatre.ActiveSheet.Cells(e.Row, SpreadIndex_Theatre_StsCd).Value
            If strStsCd = YOYAKU_STS_KARI_MI Or strStsCd = YOYAKU_STS_KARI Or strStsCd = YOYAKU_STS_CANCEL_KARI Then
                Dim frm As New EXTB0102
                Me.Hide()

                'パラメータに予約番号を設定
                frm.dataEXTB0102.PropStrYoyakuNo = vwYoyakuTheatre.ActiveSheet.Cells(e.Row, SpreadIndex_Theatre_YoyakuNo).Value

                '「仮予約登録／詳細画面(シアター)」画面を表示
                frm.ShowDialog()
            ElseIf strStsCd = YOYAKU_STS_SEISHIKI Or strStsCd = YOYAKU_STS_SEISHIKI_COMP Or strStsCd = YOYAKU_STS_CANCEL_SEISHIKI Then
                Dim frm As New EXTB0103
                Me.Hide()

                'パラメータに予約番号を設定
                frm.dataEXTB0102.PropStrYoyakuNo = vwYoyakuTheatre.ActiveSheet.Cells(e.Row, SpreadIndex_Theatre_YoyakuNo).Value

                '「正式予約登録／詳細画面(シアター)」画面を表示
                frm.ShowDialog()
            End If
            Me.Show()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'スプレッドの再作成
            If logicEXTZ0101.MakeSpread(dataEXTZ0101) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If


            '選択チェックボックスがクリックされた場合
        ElseIf e.Column = 0 Then

            RemoveHandler vwYoyakuTheatre.ButtonClicked, AddressOf vwYoyakuTheatre_ButtonClicked

            '選択行番号取得
            dataEXTZ0101.PropIntCheckRow = e.Row

            '選択行にチェックをつけ、それ以外はチェックを外す
            If logicEXTZ0101.ClickVwCellMain(dataEXTZ0101) = False Then
                'エラーメッセージを表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Exit Sub
            End If

            AddHandler vwYoyakuTheatre.ButtonClicked, AddressOf vwYoyakuTheatre_ButtonClicked

        End If
    End Sub

#End Region

#Region "予約一覧（スタジオ）のボタンをクリックした場合"

    ''' <summary>
    ''' 予約一覧（スタジオ）のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>予約一覧（スタジオ）のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/09/10 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwYoyakuStudio_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwYoyakuStudio.ButtonClicked

        '「確認・編集」ボタンが押下された場合
        If e.Column = SpreadIndex_Studio_Button Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            Dim strStsCd = vwYoyakuStudio.ActiveSheet.Cells(e.Row, SpreadIndex_Studio_StsCd).Value
            If strStsCd = YOYAKU_STS_KARI_MI Or strStsCd = YOYAKU_STS_KARI Or strStsCd = YOYAKU_STS_CANCEL_KARI Then
                Dim frm As New EXTC0102
                Me.Hide()

                'パラメータに予約番号を設定
                frm.dataEXTC0102.PropStrYoyakuNo = vwYoyakuStudio.ActiveSheet.Cells(e.Row, SpreadIndex_Studio_YoyakuNo).Value

                '「仮予約登録／詳細画面(スタジオ)」画面を表示
                frm.ShowDialog()
            ElseIf strStsCd = YOYAKU_STS_SEISHIKI Or strStsCd = YOYAKU_STS_SEISHIKI_COMP Or strStsCd = YOYAKU_STS_CANCEL_SEISHIKI Then
                Dim frm As New EXTC0103
                Me.Hide()

                'パラメータに予約番号を設定
                frm.dataEXTC0102.PropStrYoyakuNo = vwYoyakuStudio.ActiveSheet.Cells(e.Row, SpreadIndex_Studio_YoyakuNo).Value

                '「正式予約登録／詳細画面(スタジオ)」画面を表示
                frm.ShowDialog()
            End If
            Me.Show()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'スプレッドの再作成
            If logicEXTZ0101.MakeSpread(dataEXTZ0101) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If


            '選択チェックボックスがクリックされた場合
        ElseIf e.Column = 0 Then

            RemoveHandler vwYoyakuStudio.ButtonClicked, AddressOf vwYoyakuStudio_ButtonClicked

            '選択行番号取得
            dataEXTZ0101.PropIntCheckRow = e.Row

            '選択行にチェックをつけ、それ以外はチェックを外す
            If logicEXTZ0101.ClickVwCellMain(dataEXTZ0101) = False Then
                'エラーメッセージを表示
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Exit Sub
            End If

            AddHandler vwYoyakuStudio.ButtonClicked, AddressOf vwYoyakuStudio_ButtonClicked

        End If

    End Sub

#End Region

End Class
