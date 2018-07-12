Imports Common
Imports CommonEXT
Imports EXTB
Imports EXTC
Imports EXTM

''' <summary>
''' EXTZ0102
''' </summary>
''' <remarks>請求一覧画面
''' <para>作成情報：2015/09/15 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0102

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTZ0102 As New LogicEXTZ0102     'ロジッククラス
    Public dataEXTZ0102 As New DataEXTZ0102         'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「予約一覧画面」ロード時処理 "

    ''' <summary>
    ''' 「予約一覧画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「予約一覧画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0102_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '画面初期設定
        Me.rdoTheatre.Checked = True
        Me.dtpSeikyuDt_From.txtDate.Text = String.Empty
        Me.dtpSeikyuDt_To.txtDate.Text = String.Empty
        Me.dtpNyukinYoteiDt_From.txtDate.Text = String.Empty
        Me.dtpNyukinYoteiDt_To.txtDate.Text = String.Empty
        Me.dtpNyukinDt_From.txtDate.Text = String.Empty
        Me.dtpNyukinDt_To.txtDate.Text = String.Empty
        Me.txtAitesakiNm.Text = String.Empty
        Me.txtAitesakiNm_Kana.Text = String.Empty
        Me.txtRiyoNm.Text = String.Empty
        Me.txtSaijiShutsuenNm.Text = String.Empty
        Me.txtSeikyuIraiNo.Text = String.Empty
        Me.chkMinyukinOnly.Checked = False
        If Me.vwSeikyuTheatre.ActiveSheet.Rows.Count > 0 Then
            Me.vwSeikyuTheatre.Sheets(0).RemoveRows(0, Me.vwSeikyuTheatre.ActiveSheet.Rows.Count)
        End If
        If Me.vwSeikyuStudio.ActiveSheet.Rows.Count > 0 Then
            Me.vwSeikyuStudio.Sheets(0).RemoveRows(0, Me.vwSeikyuStudio.ActiveSheet.Rows.Count)
        End If
        Me.vwSeikyuTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwSeikyuTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwSeikyuStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwSeikyuStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwSeikyuStudio.Visible = False

        'プロパティに設定
        dataEXTZ0102.PropvwSeikyuTheatre = Me.vwSeikyuTheatre
        dataEXTZ0102.PropvwSeikyuStudio = Me.vwSeikyuStudio

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

#Region "「相手先検索」ボタン押下時処理 "

    ''' <summary>
    ''' 「相手先検索」ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>相手先名カナの取得を行う
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnAitesaakiSearch_Click(sender As Object, e As EventArgs) Handles btnAitesaakiSearch.Click

        '戻り値を格納している配列
        Dim ArrayList As New ArrayList
        Dim frm As New EXTM0203

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'パラメータ設定
        frm.dataEXTM0203.PropStrMenuID = "EXTZ0102"

        ''利用者一覧画面を表示
        ArrayList = frm.ShowDialog()

        If ArrayList Is Nothing Then
        Else
            '相手先名カナを設定
            Me.txtAitesakiNm_Kana.Text = ArrayList(2)
        End If

    End Sub

#End Region

#Region "「検索」ボタン押下時処理 "

    ''' <summary>
    ''' 「検索」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「戻る」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '変数宣言
        Dim regexHalfAlpNum As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]+$")     '半角英数字チェックの正規表現パターン 

        '入力チェック
        '＜半角英数字チェック＞
        '---請求依頼番号
        If Me.txtSeikyuIraiNo.Text <> String.Empty Then
            If regexHalfAlpNum.IsMatch(Me.txtSeikyuIraiNo.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "請求依頼番号"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '＜大小チェック＞
        '---請求日
        If Me.dtpSeikyuDt_From.txtDate.Text <> String.Empty AndAlso Me.dtpSeikyuDt_To.txtDate.Text <> String.Empty Then
            If commonValidate.IsDateFromTo(Me.dtpSeikyuDt_From.txtDate.Text, Me.dtpSeikyuDt_To.txtDate.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0022, "請求日（To）", "請求日（From）"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---入金予定日
        If Me.dtpNyukinYoteiDt_From.txtDate.Text <> String.Empty AndAlso Me.dtpNyukinYoteiDt_To.txtDate.Text <> String.Empty Then
            If commonValidate.IsDateFromTo(Me.dtpNyukinYoteiDt_From.txtDate.Text, Me.dtpNyukinYoteiDt_To.txtDate.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0022, "入金予定日（To）", "入金予定日（From）"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---入金日
        If Me.dtpNyukinDt_From.txtDate.Text <> String.Empty AndAlso Me.dtpNyukinDt_To.txtDate.Text <> String.Empty Then
            If commonValidate.IsDateFromTo(Me.dtpNyukinDt_From.txtDate.Text, Me.dtpNyukinDt_To.txtDate.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0022, "入金日（To）", "入金日（From）"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '検索条件をデータクラスへ格納
        With dataEXTZ0102
            If Me.rdoTheatre.Checked Then
                'シアターをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_THEATER
            Else
                'スタジオをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_STUDIO
            End If
            .PropStrSeikyuDtFrom = Me.dtpSeikyuDt_From.txtDate.Text
            .PropStrSeikyuDtTo = Me.dtpSeikyuDt_To.txtDate.Text
            .PropStrNyukinYoteiDtFrom = Me.dtpNyukinYoteiDt_From.txtDate.Text
            .PropStrNyukinYoteiDtTo = Me.dtpNyukinYoteiDt_To.txtDate.Text
            .PropStrNyukinDtFrom = Me.dtpNyukinDt_From.txtDate.Text
            .PropStrNyukinDtTo = Me.dtpNyukinDt_To.txtDate.Text
            .PropStrAiteNm = Me.txtAitesakiNm.Text
            .PropStrAiteNmKana = Me.txtAitesakiNm_Kana.Text
            .PropStrRiyoNm = Me.txtRiyoNm.Text
            .PropStrSaijiNm = Me.txtSaijiShutsuenNm.Text
            .PropStrSeikyuIraiNo = Me.txtSeikyuIraiNo.Text
            If Me.chkMinyukinOnly.Checked Then
                '「入金予定日を過ぎた未入金請求のみ表示」をチェック
                .PropBlnMinyukin = True
            Else
                '「入金予定日を過ぎた未入金請求のみ表示」を未チェック
                .PropBlnMinyukin = False
            End If
        End With

        'スプレッドの作成
        If logicEXTZ0102.MakeSpread(dataEXTZ0102) = False Then
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
    ''' <para>作成情報：2015/09/15 h.endo
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
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoTheatre_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTheatre.CheckedChanged

        If Me.rdoTheatre.Checked Then
            'シアターがチェック状態
            Me.vwSeikyuTheatre.Visible = True
            Me.vwSeikyuStudio.Visible = False
        Else
            'スタジオがチェック状態
            Me.vwSeikyuStudio.Visible = True
            Me.vwSeikyuTheatre.Visible = False
        End If

    End Sub

#End Region

#Region "請求一覧（シアター）のボタンをクリックした場合"

    ''' <summary>
    ''' 請求一覧（シアター）のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>請求一覧（シアター）のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwSeikyuTheatre_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwSeikyuTheatre.ButtonClicked

        If e.Column = SpreadSeikyuIndex_Theatre_Button Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「確認」ボタンが押下された場合
            Dim frm As New EXTB0103
            Me.Hide()

            'パラメータに請求依頼番号を設定
            frm.dataEXTB0102.PropStrYoyakuNo = vwSeikyuTheatre.ActiveSheet.Cells(e.Row, SpreadSeikyuIndex_Theatre_YoyakuNo).Value

            '「正式予約登録／詳細画面(シアター)」画面を表示
            frm.ShowDialog()
            Me.Show()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'スプレッドの作成
            If logicEXTZ0102.MakeSpread(dataEXTZ0102) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
        End If

    End Sub

#End Region

#Region "請求一覧（スタジオ）のボタンをクリックした場合"

    ''' <summary>
    ''' 請求一覧（スタジオ）のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>請求一覧（スタジオ）のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwSeikyuStudio_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwSeikyuStudio.ButtonClicked

        If e.Column = SpreadSeikyuIndex_Studio_Button Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「確認」ボタンが押下された場合
            Dim frm As New EXTC0103
            Me.Hide()

            'パラメータに予約番号を設定
            frm.dataEXTC0102.PropStrYoyakuNo = vwSeikyuStudio.ActiveSheet.Cells(e.Row, SpreadSeikyuIndex_Studio_YoyakuNo).Value

            '「正式予約登録／詳細画面(スタジオ)」画面を表示
            frm.ShowDialog()
            Me.Show()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            'スプレッドの作成
            If logicEXTZ0102.MakeSpread(dataEXTZ0102) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
        End If

    End Sub

#End Region

End Class
