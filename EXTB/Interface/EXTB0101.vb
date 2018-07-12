Imports Common
Imports CommonEXT
Imports EXTM
Imports EXTZ

''' <summary>
''' EXTB0101
''' </summary>
''' <remarks>予約カレンダー（シアター）画面
''' <para>作成情報：2015/08/04 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>

Public Class EXTB0101

    '変数宣言
    Private commonLogic As New CommonLogic          '共通クラス
    Private logicEXTB0101 As New LogicEXTB0101      'ロジッククラス
    Public dataEXTB0101 As New DataEXTB0101         'データクラス
    Public commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス
    '---コンテキストメニュー
    Private c11Menu As ContextMenuStrip  'シアター行（上段）
    Private c12Menu As ContextMenuStrip  'シアター行（下段）
    Private c21Menu As ContextMenuStrip  'キャンセル待ち行（上段）
    Private c22Menu As ContextMenuStrip  'キャンセル待ち行（下段）
    Private c31Menu As ContextMenuStrip  'その他行（上段）
    Private c32Menu As ContextMenuStrip  'その他行（下段）

    Private intSelectYear As Integer
    Private intSelectMonth As Integer

#Region "「予約カレンダー（シアター）画面」ロード時処理 "

    ''' <summary>
    ''' 「予約カレンダー（シアター）画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「予約カレンダー（シアター）画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTB0101_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''共通設定値取得
        'If commonLogic.InitCommonSetting(Nothing) = False Then
        '    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, TITLE)
        '    Return
        'End If

        '年月の設定
        'Me.txtMonth.Text = Now.Year.ToString                          ' 2016.01.04 UPD h.hagiwara
        Me.txtYear.Text = Now.Year.ToString                            ' 2016.01.04 UPD h.hagiwara
        Me.txtMonth.Text = Now.Month.ToString
        Me.intSelectYear = Now.Year
        Me.intSelectMonth = Now.Month

        'コンテキストメニューの作成
        '<予約>
        '---上段---
        c11Menu = New ContextMenuStrip()
        c11Menu.Items.Add("この日を仮予約登録する", Nothing, New System.EventHandler(AddressOf RegistKariYoyaku))
        c11Menu.Items.Add("この日を休館日とする", Nothing, New System.EventHandler(AddressOf ChangeKyukan))
        c11Menu.Items.Add("この日をメンテ日とする", Nothing, New System.EventHandler(AddressOf ChangeMainte))
        c11Menu.Items.Add("この日を営業日とする", Nothing, New System.EventHandler(AddressOf ChangeEigyo))
        c11Menu.Items.Add("この日を祝日（休日）とする", Nothing, New System.EventHandler(AddressOf ChangeHoliday))
        c11Menu.Items.Add("この日を平日とする", Nothing, New System.EventHandler(AddressOf ChangeWeekday))
        c12Menu = New ContextMenuStrip()
        '---下段---
        c12Menu.Items.Add("この日を仮予約登録する", Nothing, New System.EventHandler(AddressOf RegistKariYoyaku2))
        c12Menu.Items.Add("この日を休館日とする", Nothing, New System.EventHandler(AddressOf ChangeKyukan2))
        c12Menu.Items.Add("この日をメンテ日とする", Nothing, New System.EventHandler(AddressOf ChangeMainte2))
        c12Menu.Items.Add("この日を営業日とする", Nothing, New System.EventHandler(AddressOf ChangeEigyo2))
        c12Menu.Items.Add("この日を祝日（休日）とする", Nothing, New System.EventHandler(AddressOf ChangeHoliday2))
        c12Menu.Items.Add("この日を平日とする", Nothing, New System.EventHandler(AddressOf ChangeWeekday2))
        '<キャンセル>
        '---上段---
        c21Menu = New ContextMenuStrip()
        c21Menu.Items.Add("この日のキャンセル待ち（事前）を登録する", Nothing, New System.EventHandler(AddressOf RegistCancelWait))
        '---下段---
        c22Menu = New ContextMenuStrip()
        c22Menu.Items.Add("この日のキャンセル待ち（事前）を登録する", Nothing, New System.EventHandler(AddressOf RegistCancelWait2))
        '<その他>
        '---上段---
        c31Menu = New ContextMenuStrip()
        c31Menu.Items.Add("この日の予定を登録する", Nothing, New System.EventHandler(AddressOf RegistYotei))
        '---下段---
        c32Menu = New ContextMenuStrip()
        c32Menu.Items.Add("この日の予定を登録する", Nothing, New System.EventHandler(AddressOf RegistYotei2))

        'プロパティに設定
        dataEXTB0101.PropTxtYear = Me.txtYear
        dataEXTB0101.PropTxtMonth = Me.txtMonth
        dataEXTB0101.PropvwCalandarFirst = Me.vwCalandarFirst
        dataEXTB0101.PropvwCalandarSecond = Me.vwCalendarSecond
        dataEXTB0101.PropvwCancelWait = Me.vwCancelWait
        dataEXTB0101.PropvwCancelDone = Me.vwCancelDone

        'スプレッドスクロールバー設定
        Me.vwCancelWait.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Always
        Me.vwCancelWait.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwCancelDone.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        'スプレッドの作成
        If logicEXTB0101.MakeSpread(dataEXTB0101) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

#Region "表示ボタン押下時処理 "

    ''' <summary>
    ''' 表示ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        '年月入力チェック
        If logicEXTB0101.NengetuInputCheck(dataEXTB0101) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'スプレッドの作成
        If logicEXTB0101.MakeSpread(dataEXTB0101) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

#End Region

#Region "「前月を表示」ボタン押下時処理 "

    ''' <summary>
    ''' 「前月を表示」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「前月を表示」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearchPrevMonth_Click(sender As Object, e As EventArgs) Handles btnSearchPrevMonth.Click

        '年月を表示
        intSelectYear = Integer.Parse(dataEXTB0101.PropTxtYear.Text)
        intSelectMonth = Integer.Parse(dataEXTB0101.PropTxtMonth.Text)
        Dim datNewDate As Date = DateSerial(intSelectYear, intSelectMonth - 1, 1)
        Me.txtYear.Text = datNewDate.Year.ToString
        Me.txtMonth.Text = datNewDate.Month.ToString
        Me.intSelectYear = datNewDate.Year
        Me.intSelectMonth = datNewDate.Month

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'スプレッドの作成
        If logicEXTB0101.MakeSpread(dataEXTB0101) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

#End Region

#Region "「次月を表示」ボタン押下時処理 "

    ''' <summary>
    ''' 「次月を表示」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「次月を表示」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearchNextMonth_Click(sender As Object, e As EventArgs) Handles btnSearchNextMonth.Click

        '年月を表示
        intSelectYear = Integer.Parse(dataEXTB0101.PropTxtYear.Text)
        intSelectMonth = Integer.Parse(dataEXTB0101.PropTxtMonth.Text)
        Dim datNewDate As Date = DateSerial(intSelectYear, intSelectMonth + 1, 1)
        Me.txtYear.Text = datNewDate.Year.ToString
        Me.txtMonth.Text = datNewDate.Month.ToString
        Me.intSelectYear = datNewDate.Year
        Me.intSelectMonth = datNewDate.Month

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'スプレッドの作成
        If logicEXTB0101.MakeSpread(dataEXTB0101) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

#End Region

#Region "「利用者要注意一覧を見る」ボタン押下時処理 "

    ''' <summary>
    ''' 「利用者要注意一覧を見る」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「利用者一覧画面」に遷移し、レベルが要注意と利用不可の内容を一覧表示する
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnUserList_Click(sender As Object, e As EventArgs) Handles btnUserList.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        Dim frm As New EXTM0201
        Me.Hide()

        '「利用者一覧」画面を表示
        frm.dataEXTM0201.PropParamValue = "1"
        frm.ShowDialog()
        Me.Show()

    End Sub

#End Region

#Region "「メニューへ」ボタン押下時処理 "

    ''' <summary>
    ''' 「メニューへ」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「メニューへ」ボタン押下時の処理を行う 
    ''' <para>作成情報：2015/08/04 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnToMenu_Click(sender As Object, e As EventArgs) Handles btnToMenu.Click

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日を仮予約登録する」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を仮予約登録する」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を仮予約登録する」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistKariYoyaku(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList         '日付リスト
        Dim strDate As String = String.Empty    '一覧から選択した日付

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKariyoyaku(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTB0101.SetInsertUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTB0102
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTB0102.PropAryStrRiyoDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を仮予約登録する」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を仮予約登録する」押下時の処理（カレンダー下段）  
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistKariYoyaku2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList         '日付リスト
        Dim strDate As String = String.Empty    '一覧から選択した日付

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKariyoyaku(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTB0101.SetInsertUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTB0102
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTB0102.PropAryStrRiyoDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日を休館日とする」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を休館日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を休館日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeKyukan(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_KYUKAN) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "1"

        '---「休館日登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を休館日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を休館日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeKyukan2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_KYUKAN) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "1"

        '---「休館日登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日をメンテ日とする」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日をメンテ日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日をメンテ日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeMainte(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_MAINTE) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "2"

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日をメンテ日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日をメンテ日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeMainte2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_MAINTE) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "2"

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日を営業日とする」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を営業日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を営業日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeEigyo(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_EIGYO) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日に設定する
        If logicEXTB0101.SetEigyo(dataEXTB0101, arrDateLst(0)) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を営業日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を営業日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeEigyo2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_EIGYO) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日に設定する
        If logicEXTB0101.SetEigyo(dataEXTB0101, arrDateLst(0)) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を祝日（休日）とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/31 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeHoliday(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, arrDateLst(0), HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を祝日（休日）とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/31 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeHoliday2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, arrDateLst(0), HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

#End Region

#Region "「シアター」行上の右クリックメニュー「この日を平日とする」押下時処理"

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を平日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を平日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/31 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub Changeweekday(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, arrDateLst(0), HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

    ''' <summary>
    ''' 「シアター」行上の右クリックメニュー「この日を平日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を平日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/31 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub Changeweekday2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, arrDateLst(0), HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

#End Region

#Region "「キャンセル」行上の右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時処理"

    ''' <summary>
    ''' 「キャンセル」行上の右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistCancelWait(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim arrWakujun As New ArrayList             'キャンセル枠順
        Dim strDate As String = String.Empty        '一覧から選択した日付

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' 2016.02.19 ADD START↓ h.hagiwara 登録済み枠選択時のチェック追加
        If Me.vwCalandarFirst_Sheet1.ActiveCell.Value <> "" Then
            MsgBox(CommonDeclareEXT.E2051, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        ' 2016.02.19 ADD END↑ h.hagiwara 登録済み枠選択時のチェック追加

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'キャンセル待ち数チェック
        If logicEXTB0101.CheckCancelWait(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTB0101.SetInsertCancelUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTB0104
        Me.Hide()

        'パラメータに日付リスト・枠順を設定
        frm.dataEXTB0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTB0104.PropAryStrCancelDate.add({strDate, (Me.vwCalandarFirst_Sheet1.ActiveRowIndex - 2).ToString})

        '「キャンセル待ち登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

    ''' <summary>
    ''' 「キャンセル」行上の右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistCancelWait2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim arrWakujun As New ArrayList             'キャンセル枠順
        Dim strDate As String = String.Empty        '一覧から選択した日付

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' 2016.02.19 ADD START↓ h.hagiwara 登録済み枠選択時のチェック追加
        If Me.vwCalendarSecond_Sheet1.ActiveCell.Value <> "" Then
            MsgBox(CommonDeclareEXT.E2051, MsgBoxStyle.Exclamation, "エラー")
            Return
        End If
        ' 2016.02.19 ADD END↑ h.hagiwara 登録済み枠選択時のチェック追加

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'キャンセル待ち数チェック
        If logicEXTB0101.CheckCancelWait(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTB0101.SetInsertCancelUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        '画面遷移
        Dim frm As New EXTB0104
        Me.Hide()

        'パラメータに日付リスト・枠順を設定
        frm.dataEXTB0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTB0104.PropAryStrCancelDate.add({strDate, (Me.vwCalendarSecond_Sheet1.ActiveRowIndex - 2).ToString})

        '「キャンセル待ち登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

#End Region

#Region "「その他」行上の右クリックメニュー「この日予定を登録する」押下時処理"

    ''' <summary>
    ''' 「その他」行上の右クリックメニュー「この日予定を登録する」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日予定を登録する」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistYotei(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim strDate As String = String.Empty        '一覧から選択した日付

        '日付取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '画面遷移
        Dim frm As New EXTB0105
        Me.Hide()

        '---パラメータに日付を設定
        frm.dataEXTB0105.PropStrRiyoDt = strDate

        '---「その他利用施設登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '更新されていた場合は、最新状態を表示する
            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「その他」行上の右クリックメニュー「この日予定を登録する」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日予定を登録する」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistYotei2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim strDate As String = String.Empty        '一覧から選択した日付

        '日付取得
        '---マウスでのセル選択より取得
        strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '画面遷移
        Dim frm As New EXTB0105
        Me.Hide()

        '---パラメータに日付を設定
        frm.dataEXTB0105.PropStrRiyoDt = strDate

        '---「その他利用施設登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '更新されていた場合は、最新状態を表示する
            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "カレンダーのセル上でマウスをクリックした場合"

    ''' <summary>
    ''' カレンダー（上段）のセル上でマウスをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックした場合にメニューを表示する 
    ''' <para>作成情報：2015/08/11 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCalandarFirst_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwCalandarFirst.CellClick

        '変数宣言
        Dim strYoyakuNo As String = String.Empty            '予約番号
        Dim strCancelWaitNo As String = String.Empty        'キャンセル待ち番号 
        Dim strYoyakuKbn As String = String.Empty           '予約区分
        '定数宣言
        Const YOYAKUKBN_KARI = "1"                          '仮予約
        Const YOYAKUKBN_SEISHIKI = "2"                      '正式予約

        ' 列ヘッダのクリック時
        If e.ColumnHeader = True Then
            e.Cancel = True
            Return
        End If

        ' 行ヘッダのクリック時
        If e.RowHeader = True Then
            e.Cancel = True
            Return
        End If

        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            '■右クリックの場合

            '＜シアター行の場合＞
            If e.Row >= 1 And e.Row <= 2 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c11Menu.Show(e.X, e.Y + 100)

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 3 And e.Row <= 5 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c21Menu.Show(e.X, e.Y + 150)

            End If

            '＜その他行の場合＞
            If e.Row >= 6 And e.Row <= 7 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c31Menu.Show(e.X, e.Y + 150)

            End If
        Else
            '■左クリックの場合
            '＜シアター行の場合＞
            If e.Row >= 1 And e.Row <= 2 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                '仮予約済または正式予約済であるかの確認
                If dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "1") Then
                    '仮予約済（未確認）の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "1")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "2") Then
                    '仮予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "2")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "3") Then
                    '正式予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "3")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "4") Then
                    '正式予約済（完了）の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & e.Row.ToString & "_" & "4")
                End If

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                If strYoyakuKbn = YOYAKUKBN_KARI Then
                    '仮予約済の場合 
                    Dim frm As New EXTB0102
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTB0102.PropStrYoyakuNo = strYoyakuNo

                    '「仮予約登録」画面を表示
                    frm.ShowDialog()

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()
                ElseIf strYoyakuKbn = YOYAKUKBN_SEISHIKI Then
                    '正式予約済の場合 
                    Dim frm As New EXTB0103
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTB0102.PropStrYoyakuNo = strYoyakuNo

                    '「正式予約登録」画面を表示
                    frm.ShowDialog()

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()
                End If

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 3 And e.Row <= 5 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'キャンセル待ち情報であるかの確認
                If dataEXTB0101.PropDicCancelInfo.ContainsKey((e.Column + 1).ToString & "_" & (e.Row - 2).ToString) Then
                    'キャンセル待ち情報である場合 
                    strCancelWaitNo = dataEXTB0101.PropDicCancelInfo((e.Column + 1).ToString & "_" & (e.Row - 2).ToString)

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    Dim frm As New EXTB0104
                    Me.Hide()

                    'パラメータにキャンセル番号を設定
                    frm.dataEXTB0104.PropStrYoyakuNo = strCancelWaitNo

                    '「キャンセル待ち登録」画面を表示
                    frm.ShowDialog()

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()

                End If

            End If
        End If

    End Sub

    ''' <summary>
    ''' カレンダー（下段）のセル上でマウスをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックした場合にメニューを表示する 
    ''' <para>作成情報：2015/08/11 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCalandarSecond_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwCalendarSecond.CellClick

        '変数宣言
        Dim strYoyakuNo As String = String.Empty            '予約番号
        Dim strCancelWaitNo As String = String.Empty        'キャンセル待ち番号 
        Dim strYoyakuKbn As String = String.Empty           '予約区分
        Dim intMonthDayCnt As Integer = 0                   '指定年月に属する日数
        '定数宣言
        Const YOYAKUKBN_KARI = "1"                          '仮予約
        Const YOYAKUKBN_SEISHIKI = "2"                      '正式予約

        ' 列ヘッダのクリック時
        If e.ColumnHeader = True Then
            e.Cancel = True
            Return
        End If

        ' 行ヘッダのクリック時
        If e.RowHeader = True Then
            e.Cancel = True
            Return
        End If

        '指定月の日数を取得
        intMonthDayCnt = DateTime.DaysInMonth(Integer.Parse(dataEXTB0101.PropStrYear), Integer.Parse(dataEXTB0101.PropStrMonth))

        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            '■右クリックの場合

            '＜シアター行の場合＞
            If e.Row >= 1 And e.Row <= 2 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c12Menu.Show(e.X, e.Y + 500)

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 3 And e.Row <= 5 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c22Menu.Show(e.X, e.Y + 550)

            End If

            '＜その他行の場合＞
            If e.Row >= 6 And e.Row <= 7 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c32Menu.Show(e.X, e.Y + 550)

            End If

        Else
            '■左クリックの場合
            '＜シアター行の場合＞
            If e.Row >= 1 And e.Row <= 2 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                '仮予約済または正式予約済であるかの確認
                If dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "1") Then
                    '仮予約済（未確認）の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "1")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "2") Then
                    '仮予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "2")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "3") Then
                    '正式予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "3")
                ElseIf dataEXTB0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "4") Then
                    '正式予約済（完了）の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTB0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & e.Row.ToString & "_" & "4")
                End If

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                If strYoyakuKbn = YOYAKUKBN_KARI Then
                    '仮予約済の場合 
                    Dim frm As New EXTB0102
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTB0102.PropStrYoyakuNo = strYoyakuNo

                    '「仮予約登録」画面を表示
                    frm.ShowDialog()

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()

                ElseIf strYoyakuKbn = YOYAKUKBN_SEISHIKI Then
                    '正式予約済の場合 
                    Dim frm As New EXTB0103
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTB0102.PropStrYoyakuNo = strYoyakuNo

                    '「正式予約登録」画面を表示
                    frm.ShowDialog()

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()

                End If

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 3 And e.Row <= 5 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'キャンセル待ち情報であるかの確認
                If dataEXTB0101.PropDicCancelInfo.ContainsKey((e.Column + 16).ToString & "_" & (e.Row - 2).ToString) Then
                    'キャンセル待ち情報である場合 
                    strCancelWaitNo = dataEXTB0101.PropDicCancelInfo((e.Column + 16).ToString & "_" & (e.Row - 2).ToString)

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    Dim frm As New EXTB0104
                    Me.Hide()

                    'パラメータにキャンセル番号を設定
                    frm.dataEXTB0104.PropStrYoyakuNo = strCancelWaitNo

                    '「キャンセル待ち登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTB0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With
                    logicEXTB0101.MakeSpread(dataEXTB0101)

                    Me.Show()

                End If

            End If

        End If

    End Sub

#End Region

#Region "キャンセル待ち一覧のセル上のボタンをクリックした場合"

    ''' <summary>
    ''' キャンセル待ち一覧のセル上のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>キャンセル待ち一覧のセル上のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCancelWait_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwCancelWait.ButtonClicked

        If e.Column = 4 Or e.Column = 5 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「詳細」ボタンまたは「日付未定キャンセル待ち登録」が押下された場合
            Dim frm As New EXTB0104
            Me.Hide()

            If e.Column = 4 Then
                '「詳細」ボタンが押下された場合はパラメータにキャンセル待ちNOを設定
                frm.dataEXTB0104.PropStrYoyakuNo = vwCancelWait.ActiveSheet.Cells(e.Row, 6).Value
            End If

            '「キャンセル待ち登録」画面を表示
            frm.ShowDialog()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '前回表示時の値でカレンダーを再表示する
            With dataEXTB0101
                .PropTxtYear.Text = .PropStrYear
                .PropTxtMonth.Text = .PropStrMonth
            End With
            logicEXTB0101.MakeSpread(dataEXTB0101)

            Me.Show()
        End If

    End Sub

#End Region

#Region "キャンセル済一覧のセル上のボタンをクリックした場合"

    ''' <summary>
    ''' キャンセル済一覧のセル上のボタンをクリックした場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>キャンセル済一覧のセル上のボタンをクリックした場合の処理を行う
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCancelDone_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwCancelDone.ButtonClicked

        If e.Column = 3 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '「詳細」ボタンが押下された場合
            Dim frm As New EXTB0104
            Me.Hide()

            'パラメータにキャンセル待ちNOを設定
            frm.dataEXTB0104.PropStrYoyakuNo = vwCancelDone.ActiveSheet.Cells(e.Row, 4).Value

            '「キャンセル待ち登録」画面を表示
            frm.ShowDialog()

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '前回表示時の値でカレンダーを再表示する
            With dataEXTB0101
                .PropTxtYear.Text = .PropStrYear
                .PropTxtMonth.Text = .PropStrMonth
            End With
            logicEXTB0101.MakeSpread(dataEXTB0101)

            Me.Show()
        End If

    End Sub

#End Region

#Region "「選択日を仮予約登録」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日を仮予約登録」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日を仮予約登録」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRegistKariYoyaku_Click(sender As Object, e As EventArgs) Handles btnRegistKariYoyaku.Click

        '変数宣言
        Dim arrDateLst As New ArrayList         '日付リスト
        Dim intDayCount As Integer = 0          '選択した日数
        Dim strDate As String = String.Empty    '選択日    

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKariyoyaku(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTB0101.SetInsertUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTB0102
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTB0102.PropAryStrRiyoDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

#End Region

#Region "「選択日をキャンセル待ち登録」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日をキャンセル待ち登録」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日をキャンセル待ち登録」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnRegistCancelWait_Click(sender As Object, e As EventArgs) Handles btnRegistCancelWait.Click

        '変数宣言
        Dim arrDateLst As New ArrayList         '選択毎の日付
        Dim intDayCount As Integer = 0          '選択した日数
        Dim strDate As String = String.Empty    '選択日    

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'キャンセル待ち数チェック
        If logicEXTB0101.CheckCancelWait(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTB0101.SetInsertCancelUketukeSeigyo(dataEXTB0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        '画面遷移
        Dim frm As New EXTB0104
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTB0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTB0104.PropAryStrCancelDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '前回表示時の値でカレンダーを再表示する
        With dataEXTB0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With
        logicEXTB0101.MakeSpread(dataEXTB0101)

        Me.Show()

    End Sub

#End Region

#Region "「選択日を祝日（休日）とする」ボタン押下時処理 "

    ''' <summary>
    ''' 「択日を祝日（休日）とする」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「択日を祝日（休日）とする」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnChangeHoliday_Click(sender As Object, e As EventArgs) Handles btnChangeHoliday.Click

        '変数宣言
        Dim strDate As String = String.Empty    '選択毎の日付
        Dim intDayCount As Integer = 0          '選択した日数
        Dim arrDateLst As New ArrayList         '日付リスト

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        For Each strDate In arrDateLst
            If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, strDate, HOLIYDAY_ON) = False Then
                If puErrMsg <> String.Empty Then
                    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                End If
                Return
            End If
        Next

    End Sub

#End Region

#Region "「選択日を平日とする」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日を平日とする」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日を平日とする」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnChangeWeekday_Click(sender As Object, e As EventArgs) Handles btnChangeWeekday.Click

        '変数宣言
        Dim strDate As String = String.Empty        '選択毎日付
        Dim intDayCount As Integer = 0              '選択した日数
        Dim arrDateLst As New ArrayList             '日付リスト

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTB0101.CheckHoliday(dataEXTB0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        For Each strDate In arrDateLst
            If logicEXTB0101.SetHolidayWeekday(dataEXTB0101, strDate, HOLIYDAY_OFF) = False Then
                If puErrMsg <> String.Empty Then
                    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                End If
                Return
            End If
        Next

    End Sub

#End Region

#Region "「選択日を休館日とする」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日を休館日とする」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日を休館日とする」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnChangeKyukanbi_Click(sender As Object, e As EventArgs) Handles btnChangeKyukanbi.Click

        '変数宣言
        Dim strDate As String = String.Empty       '選択毎日付
        Dim intDayCount As Integer = 0             '選択した日数
        Dim arrDateLst As New ArrayList            '日付リスト

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_KYUKAN) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "1"

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「選択日をメンテ日とする」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日をメンテ日とする」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日をメンテ日とする」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnChangeMaintebi_Click(sender As Object, e As EventArgs) Handles btnChangeMaintebi.Click

        '変数宣言
        Dim strDate As String = String.Empty       '選択毎日付
        Dim intDayCount As Integer = 0             '選択した日数
        Dim arrDateLst As New ArrayList            '日付リスト

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_MAINTE) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '画面遷移
        Dim frm As New EXTZ0201
        Me.Hide()

        '---パラメータに選択日付、施設区分を設定
        frm.dataEXTZ0201.PropStrHolmentDt = arrDateLst(0)
        frm.dataEXTZ0201.PropStrShisetuKbn = "1"
        frm.dataEXTZ0201.PropStrStudioKbn = "0"
        frm.dataEXTZ0201.PropHolmentKbn = "2"

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTB0101.MakeSpread(dataEXTB0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「選択日を営業日とする」ボタン押下時処理 "

    ''' <summary>
    ''' 「選択日を営業日とする」ボタン押下時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「選択日を営業日とする」ボタン押下時の処理 
    ''' <para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnChangeEigyobi_Click(sender As Object, e As EventArgs) Handles btnChangeEigyobi.Click

        '変数宣言
        Dim strDate As String = String.Empty       '選択毎日付
        Dim intDayCount As Integer = 0             '選択した日数
        Dim arrDateLst As New ArrayList            '日付リスト

        '日付リスト取得
        '---チェックボックスでのセル選択より取得
        For intCol As Integer = 0 To Me.vwCalandarFirst.ActiveSheet.Columns.Count - 1
            If Me.vwCalandarFirst.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTB0101.PropStrYear & "/" & dataEXTB0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next

        '日付の選択が1件もない場合はエラー
        If intDayCount = 0 Then
            MsgBox(String.Format(CommonEXT.E0002, "いずれかの日付"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
            Return
        End If

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTB0101.CheckSelectDateKyukanMainteEigyo(dataEXTB0101, arrDateLst, SETDATE_EIGYO) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日の設定
        For Each strDate In arrDateLst
            If logicEXTB0101.SetEigyo(dataEXTB0101, strDate) = False Then
                If puErrMsg <> String.Empty Then
                    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                End If
                Return
            End If
        Next

    End Sub

#End Region

End Class