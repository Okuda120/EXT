Imports Common
Imports CommonEXT
Imports EXTM

''' <summary>
''' EXTZ0103
''' </summary>
''' <remarks>日別売上一覧画面
''' <para>作成情報：2015/09/16 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
''' 
Public Class EXTZ0103

    '変数宣言
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonValidate As New CommonValidation  '共通バリデーションクラス
    Private logicEXTZ0103 As New LogicEXTZ0103     'ロジッククラス
    Public dataEXTZ0103 As New DataEXTZ0103         'データクラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「日別売上一覧画面」ロード時処理 "

    ''' <summary>
    ''' 「日別売上一覧画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「日別売上一覧画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0103_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '画面初期設定
        Me.rdoTheatre.Checked = True
        Me.txtShiyoNen_From.Text = String.Empty
        Me.txtShiyoTsuki_From.Text = String.Empty
        Me.txtShiyoNen_To.Text = String.Empty
        Me.txtShiyoTsuki_To.Text = String.Empty
        Me.txtRiyoNm.Text = String.Empty
        Me.txtRiyoNm_Kana.Text = String.Empty
        If Me.vwDayUriageTheatre.ActiveSheet.Rows.Count > 0 Then
            Me.vwDayUriageTheatre.Sheets(0).RemoveRows(0, Me.vwDayUriageTheatre.ActiveSheet.Rows.Count)
        End If
        If Me.vwSeikyuChoseiTheatre.ActiveSheet.Rows.Count > 0 Then
            Me.vwSeikyuChoseiTheatre.Sheets(0).RemoveRows(0, Me.vwSeikyuChoseiTheatre.ActiveSheet.Rows.Count)
        End If
        If Me.vwDayUriageStudio.ActiveSheet.Rows.Count > 0 Then
            Me.vwDayUriageStudio.Sheets(0).RemoveRows(0, Me.vwDayUriageStudio.ActiveSheet.Rows.Count)
        End If
        If Me.vwSeikyuChoseiStudio.ActiveSheet.Rows.Count > 0 Then
            Me.vwSeikyuChoseiStudio.Sheets(0).RemoveRows(0, Me.vwSeikyuChoseiStudio.ActiveSheet.Rows.Count)
        End If
        Me.vwDayUriageTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Always
        Me.vwDayUriageTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwSeikyuChoseiTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwSeikyuChoseiTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwDayUriageStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Always
        Me.vwDayUriageStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwSeikyuChoseiStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwSeikyuChoseiStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwDayUriageStudio.Visible = False
        Me.vwSeikyuChoseiStudio.Visible = False

        'プロパティに設定
        dataEXTZ0103.PropvwDayUriageTheatre = Me.vwDayUriageTheatre
        dataEXTZ0103.PropvwDayUriageStudio = Me.vwDayUriageStudio
        dataEXTZ0103.PropvwSeikyuChoseiTheatre = Me.vwSeikyuChoseiTheatre
        dataEXTZ0103.PropvwSeikyuChoseiStudio = Me.vwSeikyuChoseiStudio

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
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

        'パラメータ設定  2015.11.25 ADD h.hagiwara
        frm.DataEXTM0201.PropTxtRiyo_kana = Me.txtRiyoNm_Kana
        frm.DataEXTM0201.PropParamValue = "0"

        '利用者一覧画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '利用者名カナを設定
            'Me.txtRiyoNm_Kana.Text = frm.dataEXTM0201.PropStrRiyoKana         ' 2015.11.25 UPD h.hagiwara
            Me.txtRiyoNm_Kana.Text = frm.DataEXTM0201.PropParamRiyoKana        ' 2015.11.25 UPD h.hagiwara
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
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        '変数宣言
        Dim strNengetsuFrom As String = String.Empty        '開始年月ワーク
        Dim strNengetsuTo As String = String.Empty          '終了年月ワーク

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        '＜半角英数字チェック＞
        '---使用年（From）
        If Me.txtShiyoNen_From.Text <> String.Empty Then
            If commonValidate.IsHalfNmb(txtShiyoNen_From.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "年(From)"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---使用月（From）
        If Me.txtShiyoTsuki_From.Text <> String.Empty Then
            If commonValidate.IsHalfNmb(txtShiyoTsuki_From.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "月(From)"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---使用年（To）
        If Me.txtShiyoNen_To.Text <> String.Empty Then
            If commonValidate.IsHalfNmb(txtShiyoNen_To.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "年(To)"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---使用月（To）
        If Me.txtShiyoTsuki_To.Text <> String.Empty Then
            If commonValidate.IsHalfNmb(txtShiyoTsuki_To.Text) = False Then
                MsgBox(String.Format(CommonEXT.E0005, "月(To)"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '＜最小値チェック＞
        '---使用年（From）
        If Me.txtShiyoNen_From.Text <> String.Empty Then
            If Me.txtShiyoNen_From.Text.Length < 4 Then
                MsgBox(String.Format(CommonEXT.E0010, "年(From)", "4"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---使用年（To）
        If Me.txtShiyoNen_To.Text <> String.Empty Then
            If Me.txtShiyoNen_To.Text.Length < 4 Then
                MsgBox(String.Format(CommonEXT.E0010, "年(To)", "4"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '＜存在しない年月の場合、エラーメッセージを出力＞
        '---使用年月（From）
        If Me.txtShiyoNen_From.Text <> String.Empty OrElse Me.txtShiyoTsuki_From.Text <> String.Empty Then
            If IsDate(Me.txtShiyoNen_From.Text & "/" & Me.txtShiyoTsuki_From.Text & "/" & "01") = False Then
                MsgBox(String.Format(CommonEXT.E2026), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If
        '---使用年月（To）
        If Me.txtShiyoNen_To.Text <> String.Empty OrElse Me.txtShiyoTsuki_To.Text <> String.Empty Then
            If IsDate(Me.txtShiyoNen_To.Text & "/" & Me.txtShiyoTsuki_To.Text & "/" & "01") = False Then
                MsgBox(String.Format(CommonEXT.E2026), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '＜大小チェック＞
        '---使用年月
        If Me.txtShiyoNen_From.Text <> String.Empty And Me.txtShiyoTsuki_From.Text <> String.Empty _
            And Me.txtShiyoNen_To.Text <> String.Empty And Me.txtShiyoTsuki_To.Text <> String.Empty Then
            strNengetsuFrom = Me.txtShiyoNen_From.Text & Integer.Parse(Me.txtShiyoTsuki_From.Text).ToString("00")
            strNengetsuTo = Me.txtShiyoNen_To.Text & Integer.Parse(Me.txtShiyoTsuki_To.Text).ToString("00")
            If commonValidate.IsYMFromTo(strNengetsuFrom, strNengetsuTo) = False Then
                MsgBox(String.Format(CommonEXT.E0022, "年月（To）", "年月（From）"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                Return
            End If
        End If

        '検索条件をデータクラスへ格納
        With dataEXTZ0103
            If Me.rdoTheatre.Checked Then
                'シアターをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_THEATER
            Else
                'スタジオをチェック
                .PropStrShisetsuKbn = SHISETU_KBN_STUDIO
            End If
            .PropStrShiyoNenFrom = Me.txtShiyoNen_From.Text
            If Me.txtShiyoTsuki_From.Text <> String.Empty Then
                .PropStrShiyoTsukiFrom = Integer.Parse(Me.txtShiyoTsuki_From.Text).ToString("00")
            Else
                .PropStrShiyoTsukiFrom = Me.txtShiyoTsuki_From.Text
            End If
            .PropStrShiyoNenTo = Me.txtShiyoNen_To.Text
            If Me.txtShiyoTsuki_To.Text <> String.Empty Then
                .PropStrShiyoTsukiTo = Integer.Parse(Me.txtShiyoTsuki_To.Text).ToString("00")
            Else
                .PropStrShiyoTsukiTo = Me.txtShiyoTsuki_To.Text
            End If
            .PropStrShiyoTsukiTo = Me.txtShiyoTsuki_To.Text
            .PropStrRiyoNm = Me.txtRiyoNm.Text
            .PropStrRiyoNmKana = Me.txtRiyoNm_Kana.Text
        End With

        'スプレッドの作成
        If logicEXTZ0103.MakeSpread(dataEXTZ0103) = False Then
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

#Region "検索対象ラジオボタンのチェック状態が変わった場合の処理 "

    ''' <summary>
    ''' 検索対象ラジオボタンのチェック状態が変わった場合の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示する一覧を切り替える 
    ''' <para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoTheatre_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTheatre.CheckedChanged

        If Me.rdoTheatre.Checked Then
            'シアターがチェック状態
            Me.vwDayUriageTheatre.Visible = True
            Me.vwSeikyuChoseiTheatre.Visible = True
            Me.vwDayUriageStudio.Visible = False
            Me.vwSeikyuChoseiStudio.Visible = False
        Else
            'スタジオがチェック状態
            Me.vwDayUriageStudio.Visible = True
            Me.vwSeikyuChoseiStudio.Visible = True
            Me.vwDayUriageTheatre.Visible = False
            Me.vwSeikyuChoseiTheatre.Visible = False
        End If

    End Sub

#End Region

End Class
