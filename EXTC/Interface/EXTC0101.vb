Imports Common
Imports CommonEXT
Imports EXTM
Imports EXTZ

''' <summary>
''' EXTC0101
''' </summary>
''' <remarks>予約カレンダー（スタジオ）画面
''' <para>作成情報：2015/09/01 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTC0101

    '変数宣言
    Private commonLogic As New CommonLogic          '共通クラス
    Private logicEXTC0101 As New LogicEXTC0101      'ロジッククラス
    Public dataEXTC0101 As New DataEXTC0101         'データクラス
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

#Region "「予約カレンダー（スタジオ）画面」ロード時処理 "

    ''' <summary>
    ''' 「予約カレンダー（スタジオ）画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「予約カレンダー（スタジオ）画面」ロード時の処理を行う 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTC0101_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        c11Menu.Items.Add("この日を平日とする", Nothing, New System.EventHandler(AddressOf Changeweekday))
        c12Menu = New ContextMenuStrip()
        '---下段---
        c12Menu.Items.Add("この日を仮予約登録する", Nothing, New System.EventHandler(AddressOf RegistKariYoyaku2))
        c12Menu.Items.Add("この日を休館日とする", Nothing, New System.EventHandler(AddressOf ChangeKyukan2))
        c12Menu.Items.Add("この日をメンテ日とする", Nothing, New System.EventHandler(AddressOf ChangeMainte2))
        c12Menu.Items.Add("この日を営業日とする", Nothing, New System.EventHandler(AddressOf ChangeEigyo2))
        c12Menu.Items.Add("この日を祝日（休日）とする", Nothing, New System.EventHandler(AddressOf ChangeHoliday2))
        c12Menu.Items.Add("この日を平日とする", Nothing, New System.EventHandler(AddressOf Changeweekday2))
        '<キャンセル>
        '---上段---
        c21Menu = New ContextMenuStrip()
        c21Menu.Items.Add("この日のキャンセル待ち（事前）を登録する", Nothing, New System.EventHandler(AddressOf RegistCancelWait))
        '---下段---
        c22Menu = New ContextMenuStrip()
        c22Menu.Items.Add("この日のキャンセル待ち（事前）を登録する", Nothing, New System.EventHandler(AddressOf RegistCancelWait2))

        'プロパティに設定
        dataEXTC0101.PropTxtYear = Me.txtYear
        dataEXTC0101.PropTxtMonth = Me.txtMonth
        dataEXTC0101.PropvwCalandarFirst = Me.vwCalandarFirst
        dataEXTC0101.PropvwCalandarSecond = Me.vwCalendarSecond
        dataEXTC0101.PropvwCancelWait = Me.vwCancelWait
        dataEXTC0101.PropvwCancelDone = Me.vwCancelDone

        'スプレッドスクロールバー設定
        Me.vwCancelWait.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Always
        Me.vwCancelWait.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        Me.vwCancelDone.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Always
        Me.vwCancelDone.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

        'スプレッドの作成
        If logicEXTC0101.MakeSpread(dataEXTC0101) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '年月入力チェック
        If logicEXTC0101.NengetuInputCheck(dataEXTC0101) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        'スプレッドの作成
        If logicEXTC0101.MakeSpread(dataEXTC0101) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearchPrevMonth_Click(sender As Object, e As EventArgs) Handles btnSearchPrevMonth.Click

        '年月を表示
        intSelectYear = Integer.Parse(dataEXTC0101.PropTxtYear.Text)
        intSelectMonth = Integer.Parse(dataEXTC0101.PropTxtMonth.Text)
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
        If logicEXTC0101.MakeSpread(dataEXTC0101) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnSearchNextMonth_Click(sender As Object, e As EventArgs) Handles btnSearchNextMonth.Click

        '年月を表示
        intSelectYear = Integer.Parse(dataEXTC0101.PropTxtYear.Text)
        intSelectMonth = Integer.Parse(dataEXTC0101.PropTxtMonth.Text)
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
        If logicEXTC0101.MakeSpread(dataEXTC0101) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
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
        frm.dataEXTM0201.PropParamValue = RIYOSHA_LV1
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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnToMenu_Click(sender As Object, e As EventArgs) Handles btnToMenu.Click

        '画面を閉じる
        Me.Close()

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日を仮予約登録する」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を仮予約登録する」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を仮予約登録する」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistKariYoyaku(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList         '日付リスト
        Dim strDate As String = String.Empty    '一覧から選択した日付
        Dim strStudioKbn As String = String.Empty    'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 1 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 7 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKariyoyaku(dataEXTC0101, arrDateLst, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTC0101.SetInsertUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTC0102
        Me.Hide()

        'パラメータに日付リスト、スタジオ区分を設定
        frm.dataEXTC0102.PropAryStrRiyoDate = arrDateLst
        frm.dataEXTC0102.PropStrStudioKbn = strStudioKbn

        '「仮予約登録」画面を表示
        frm.ShowDialog()
        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        logicEXTC0101.MakeSpread(dataEXTC0101)

        Me.Show()

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を仮予約登録する」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を仮予約登録する」押下時の処理（カレンダー下段）  
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistKariYoyaku2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList         '日付リスト
        Dim strDate As String = String.Empty    '一覧から選択した日付
        Dim strStudioKbn As String = String.Empty    'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 1 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 7 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKariyoyaku(dataEXTC0101, arrDateLst, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTC0101.SetInsertUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTC0102
        Me.Hide()

        'パラメータに日付リスト、スタジオ区分を設定
        frm.dataEXTC0102.PropAryStrRiyoDate = arrDateLst
        frm.dataEXTC0102.PropStrStudioKbn = strStudioKbn

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        logicEXTC0101.MakeSpread(dataEXTC0101)

        Me.Show()

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日を休館日とする」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を休館日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を休館日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeKyukan(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList                 '日付リスト
        Dim strDate As String = String.Empty            '一覧から選択した日付
        Dim strStudioKbn As String = String.Empty       'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 1 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 7 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_KYUKAN, OPERATE_RIGHTMENU) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropStrStudioKbn = strStudioKbn
        frm.dataEXTZ0201.PropHolmentKbn = "1"

        '---「休館日登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            logicEXTC0101.MakeSpread(dataEXTC0101)
        End If
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を休館日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を休館日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeKyukan2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim strStudioKbn As String = String.Empty   'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 1 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 7 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_KYUKAN, OPERATE_RIGHTMENU) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropStrStudioKbn = strStudioKbn
        frm.dataEXTZ0201.PropHolmentKbn = "1"

        '---「休館日登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日をメンテ日とする」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日をメンテ日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日をメンテ日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeMainte(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim strStudioKbn As String = String.Empty   'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 1 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 7 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_MAINTE, OPERATE_RIGHTMENU) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropStrStudioKbn = strStudioKbn
        frm.dataEXTZ0201.PropHolmentKbn = "2"
        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)
        End If
        Me.Show()

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日をメンテ日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日をメンテ日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeMainte2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim strStudioKbn As String = String.Empty   'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 1 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 7 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_MAINTE, OPERATE_RIGHTMENU) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropStrStudioKbn = strStudioKbn
        frm.dataEXTZ0201.PropHolmentKbn = "2"

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)
        End If
        Me.Show()

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日を営業日とする」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を営業日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を営業日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeEigyo(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim strStudioKbn As String = String.Empty   'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 1 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 7 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_EIGYO, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日に設定する
        If logicEXTC0101.SetEigyo(dataEXTC0101, arrDateLst(0)) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を営業日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を営業日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeEigyo2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim strStudioKbn As String = String.Empty   'スタジオ区分

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 1 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 3 Then
            '201stの場合
            strStudioKbn = STUDIO_201
        ElseIf Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 7 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 9 Then
            '202stの場合
            strStudioKbn = STUDIO_202
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '選択日のチェック
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_EIGYO, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日に設定する
        If logicEXTC0101.SetEigyo(dataEXTC0101, arrDateLst(0)) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
        End If

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を祝日（休日）とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
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
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, arrDateLst(0), HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を祝日（休日）とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を祝日（休日）とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub ChangeHoliday2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, arrDateLst(0), HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

#End Region

#Region "「スタジオ」行上の右クリックメニュー「この日を平日とする」押下時処理"

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を平日とする」押下時処理（カレンダー上段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を平日とする」押下時の処理（カレンダー上段） 
    ''' <para>作成情報：2015/09/01 h.endo
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
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, arrDateLst(0), HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

    End Sub

    ''' <summary>
    ''' 「スタジオ」行上の右クリックメニュー「この日を平日とする」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日を平日とする」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub Changeweekday2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList             '日付リスト
        Dim strDate As String = String.Empty        '一覧から選択した日付
        Dim strUpdMode As String = String.Empty     '更新モード
        Dim dtMainte As New DataTable               'メンテ日情報

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '祝祭日マスタのチェック
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, arrDateLst(0), HOLIYDAY_OFF) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistCancelWait(ByVal sender As Object, ByVal e As System.EventArgs)

        '変数宣言
        Dim arrDateLst As New ArrayList                 '日付リスト
        Dim arrWakujun As New ArrayList                 'キャンセル枠順
        Dim strDate As String = String.Empty            '一覧から選択した日付
        Dim strStudioKbn As String = String.Empty       'スタジオ区分
        Dim strWakuNo As String = String.Empty          '枠番号

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalandarFirst_Sheet1.ActiveCell.Column.Index + 1).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 4 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 6 Then
            '201stの場合
            strStudioKbn = STUDIO_201
            strWakuNo = (Me.vwCalandarFirst_Sheet1.ActiveRowIndex - 3).ToString
        ElseIf Me.vwCalandarFirst_Sheet1.ActiveRowIndex >= 10 And Me.vwCalandarFirst_Sheet1.ActiveRowIndex <= 12 Then
            '202stの場合
            strStudioKbn = STUDIO_202
            strWakuNo = (Me.vwCalandarFirst_Sheet1.ActiveRowIndex - 9).ToString
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

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
        If logicEXTC0101.CheckCancelWait(dataEXTC0101, arrDateLst, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTC0101.SetInsertCancelUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTC0104
        Me.Hide()

        'パラメータに日付リスト・枠順、スタジオ区分を設定
        frm.dataEXTC0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTC0104.PropAryStrCancelDate.add({strDate, strWakuNo})
        frm.dataEXTC0104.PropStrStudioKbn = strStudioKbn

        '「キャンセル待ち登録」画面を表示
        frm.ShowDialog()

        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        logicEXTC0101.MakeSpread(dataEXTC0101)

        Me.Show()

    End Sub

    ''' <summary>
    ''' 「キャンセル」行上の右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時処理（カレンダー下段） 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>右クリックメニュー「この日のキャンセル待ち（事前）を登録する」押下時の処理（カレンダー下段） 
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub RegistCancelWait2(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim arrDateLst As New ArrayList                 '日付リスト
        Dim arrWakujun As New ArrayList                 'キャンセル枠順
        Dim strDate As String = String.Empty            '一覧から選択した日付
        Dim strStudioKbn As String = String.Empty       'スタジオ区分
        Dim strWakuNo As String = String.Empty          '枠番号

        '日付リスト取得
        '---マウスでのセル選択より取得
        strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (Me.vwCalendarSecond_Sheet1.ActiveCell.Column.Index + 16).ToString("00")
        arrDateLst.Add(strDate)

        '選択行よりスタジオ区分を取得
        If Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 4 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 6 Then
            '201stの場合
            strStudioKbn = STUDIO_201
            strWakuNo = (Me.vwCalendarSecond_Sheet1.ActiveRowIndex - 3).ToString
        ElseIf Me.vwCalendarSecond_Sheet1.ActiveRowIndex >= 10 And Me.vwCalendarSecond_Sheet1.ActiveRowIndex <= 12 Then
            '202stの場合
            strStudioKbn = STUDIO_202
            strWakuNo = (Me.vwCalendarSecond_Sheet1.ActiveRowIndex - 9).ToString
        End If

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

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
        If logicEXTC0101.CheckCancelWait(dataEXTC0101, arrDateLst, OPERATE_RIGHTMENU) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTC0101.SetInsertCancelUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        '画面遷移
        Dim frm As New EXTC0104
        Me.Hide()

        'パラメータに日付リスト・枠順、スタジオ区分を設定
        frm.dataEXTC0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTC0104.PropAryStrCancelDate.add({strDate, strWakuNo})
        frm.dataEXTC0104.PropStrStudioKbn = strStudioKbn

        'スタジオ区分をデータクラスに設定
        dataEXTC0101.PropStrStudioKbn = strStudioKbn

        '「キャンセル待ち登録」画面を表示
        frm.ShowDialog()

        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCalandarFirst_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwCalandarFirst.CellClick

        '変数宣言
        Dim strYoyakuNo As String = String.Empty            '予約番号
        Dim strCancelWaitNo As String = String.Empty        'キャンセル待ち番号 
        Dim strYoyakuKbn As String = String.Empty           '予約区分
        Dim strStudioKbn As String = String.Empty           'スタジオ区分
        Dim intYoyakuRow As Integer = 0                     '予約行
        Dim intCancelRow As Integer = 0                     'キャンセル行

        '定数宣言
        '---予約区分
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

            '＜スタジオ行の場合＞
            If e.Row >= 1 And e.Row <= 3 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c11Menu.Show(e.X, e.Y + 100)

            ElseIf e.Row >= 7 And e.Row <= 9 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c11Menu.Show(e.X, e.Y + 130)

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 4 And e.Row <= 6 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c21Menu.Show(e.X, e.Y + 140)

            ElseIf e.Row >= 10 And e.Row <= 12 And e.Column >= 0 And e.Column <= 14 Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c21Menu.Show(e.X, e.Y + 150)

            End If

        Else
            '■左クリックの場合
            '＜スタジオ行の場合＞
            If ((e.Row >= 1 And e.Row <= 3) Or (e.Row >= 7 And e.Row <= 9)) And _
               (e.Column >= 0 And e.Column <= 14) Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'スタジオ区分取得
                If e.Row >= 1 And e.Row <= 3 Then
                    strStudioKbn = STUDIO_201
                    intYoyakuRow = e.Row
                Else
                    strStudioKbn = STUDIO_202
                    intYoyakuRow = e.Row - 6
                End If

                '仮予約済または正式予約済であるかの確認
                If dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_KARI) Then
                    '仮予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_KARI)
                ElseIf dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI) Then
                    '正式予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI)
                ElseIf dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI_COMP) Then
                    '正式予約済（完了）の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI_COMP)
                End If

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                If strYoyakuKbn = YOYAKUKBN_KARI Then
                    '仮予約済の場合 
                    Dim frm As New EXTC0102
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTC0102.PropStrYoyakuNo = strYoyakuNo

                    '「仮予約登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

                    Me.Show()
                ElseIf strYoyakuKbn = YOYAKUKBN_SEISHIKI Then
                    '正式予約済の場合 
                    Dim frm As New EXTC0103
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTC0102.PropStrYoyakuNo = strYoyakuNo

                    '「正式予約登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

                    Me.Show()
                End If

            End If

            '＜キャンセル待ち行の場合＞
            If ((e.Row >= 4 And e.Row <= 6) Or (e.Row >= 10 And e.Row <= 12)) And _
                (e.Column >= 0 And e.Column <= 14) Then

                'アクティブセルの設定
                vwCalandarFirst.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'スタジオ区分取得
                If e.Row >= 4 And e.Row <= 6 Then
                    strStudioKbn = STUDIO_201
                    intCancelRow = e.Row - 3
                Else
                    strStudioKbn = STUDIO_202
                    intCancelRow = e.Row - 9
                End If

                'キャンセル待ち情報であるかの確認
                If dataEXTC0101.PropDicCancelInfo.ContainsKey((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intCancelRow.ToString) Then
                    'キャンセル待ち情報である場合 
                    strCancelWaitNo = dataEXTC0101.PropDicCancelInfo((e.Column + 1).ToString & "_" & strStudioKbn & "_" & intCancelRow.ToString)

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    Dim frm As New EXTC0104
                    Me.Hide()

                    'パラメータにキャンセル番号を設定
                    frm.dataEXTC0104.PropStrYoyakuNo = strCancelWaitNo

                    '「キャンセル待ち登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCalandarSecond_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwCalendarSecond.CellClick

        '変数宣言
        Dim strYoyakuNo As String = String.Empty            '予約番号
        Dim strCancelWaitNo As String = String.Empty        'キャンセル待ち番号 
        Dim strYoyakuKbn As String = String.Empty           '予約区分
        Dim intMonthDayCnt As Integer = 0                   '指定年月に属する日数
        Dim strStudioKbn As String = String.Empty           'スタジオ区分
        Dim intYoyakuRow As Integer = 0                     '予約行
        Dim intCancelRow As Integer = 0                     'キャンセル行

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
        intMonthDayCnt = DateTime.DaysInMonth(Integer.Parse(dataEXTC0101.PropStrYear), Integer.Parse(dataEXTC0101.PropStrMonth))

        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            '■右クリックの場合

            '＜スタジオ行の場合＞
            If e.Row >= 1 And e.Row <= 3 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c12Menu.Show(e.X, e.Y + 500)

            ElseIf e.Row >= 7 And e.Row <= 9 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c12Menu.Show(e.X, e.Y + 530)

            End If

            '＜キャンセル待ち行の場合＞
            If e.Row >= 4 And e.Row <= 6 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c22Menu.Show(e.X, e.Y + 550)

            ElseIf e.Row >= 10 And e.Row <= 12 And e.Column >= 0 And e.Column <= intMonthDayCnt - 16 Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'コンテキストメニューの表示
                c22Menu.Show(e.X, e.Y + 560)

            End If

        Else
            '■左クリックの場合
            '＜スタジオ行の場合＞
            If ((e.Row >= 1 And e.Row <= 3) Or (e.Row >= 7 And e.Row <= 9)) And _
                (e.Column >= 0 And e.Column <= intMonthDayCnt - 16) Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'スタジオ区分取得
                If e.Row >= 1 And e.Row <= 3 Then
                    strStudioKbn = STUDIO_201
                    intYoyakuRow = e.Row
                Else
                    strStudioKbn = STUDIO_202
                    intYoyakuRow = e.Row - 6
                End If

                '仮予約済または正式予約済であるかの確認
                If dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_KARI) Then
                    '仮予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_KARI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_KARI)
                ElseIf dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI) Then
                    '正式予約済の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI)
                ElseIf dataEXTC0101.PropDicYoyakuInfo.ContainsKey((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI_COMP) Then
                    '正式予約済（完了）の場合 
                    strYoyakuKbn = YOYAKUKBN_SEISHIKI
                    strYoyakuNo = dataEXTC0101.PropDicYoyakuInfo((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intYoyakuRow.ToString & "_" & YOYAKU_STS_SEISHIKI_COMP)
                End If

                ' DBレプリケーション
                If commonLogicEXT.CheckDBCondition() = False Then
                    'メッセージを出力 
                    MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                    Exit Sub
                End If

                If strYoyakuKbn = YOYAKUKBN_KARI Then
                    '仮予約済の場合 
                    Dim frm As New EXTC0102
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTC0102.PropStrYoyakuNo = strYoyakuNo

                    '「仮予約登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

                    Me.Show()

                ElseIf strYoyakuKbn = YOYAKUKBN_SEISHIKI Then
                    '正式予約済の場合 
                    Dim frm As New EXTC0103
                    Me.Hide()

                    'パラメータに予約番号を設定
                    frm.dataEXTC0102.PropStrYoyakuNo = strYoyakuNo

                    '「正式予約登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

                    Me.Show()

                End If

            End If

            '＜キャンセル待ち行の場合＞
            If ((e.Row >= 4 And e.Row <= 6) Or (e.Row >= 10 And e.Row <= 12)) And _
                (e.Column >= 0 And e.Column <= intMonthDayCnt - 16) Then

                'アクティブセルの設定
                vwCalendarSecond.ActiveSheet.SetActiveCell(e.Row, e.Column)

                'スタジオ区分取得
                If e.Row >= 4 And e.Row <= 6 Then
                    strStudioKbn = STUDIO_201
                    intCancelRow = e.Row - 3
                Else
                    strStudioKbn = STUDIO_202
                    intCancelRow = e.Row - 9
                End If

                'キャンセル待ち情報であるかの確認
                If dataEXTC0101.PropDicCancelInfo.ContainsKey((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intCancelRow.ToString) Then
                    'キャンセル待ち情報である場合 
                    strCancelWaitNo = dataEXTC0101.PropDicCancelInfo((e.Column + 16).ToString & "_" & strStudioKbn & "_" & intCancelRow.ToString)

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If

                    Dim frm As New EXTC0104
                    Me.Hide()

                    'パラメータにキャンセル番号を設定
                    frm.dataEXTC0104.PropStrYoyakuNo = strCancelWaitNo

                    '「キャンセル待ち登録」画面を表示
                    frm.ShowDialog()

                    '前回表示時の値でカレンダーを再表示する
                    With dataEXTC0101
                        .PropTxtYear.Text = .PropStrYear
                        .PropTxtMonth.Text = .PropStrMonth
                    End With

                    ' DBレプリケーション
                    If commonLogicEXT.CheckDBCondition() = False Then
                        'メッセージを出力 
                        MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                        Exit Sub
                    End If
                    logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCancelWait_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwCancelWait.ButtonClicked

        If e.Column = 5 Or e.Column = 6 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            '「詳細」ボタンまたは「日付未定キャンセル待ち登録」が押下された場合
            Dim frm As New EXTC0104
            Me.Hide()

            If e.Column = 5 Then
                '「詳細」ボタンが押下された場合はパラメータにキャンセル待ちNOを設定
                frm.dataEXTC0104.PropStrYoyakuNo = vwCancelWait.ActiveSheet.Cells(e.Row, 7).Value
            End If

            '「キャンセル待ち登録」画面を表示
            frm.ShowDialog()

            '前回表示時の値でカレンダーを再表示する
            With dataEXTC0101
                .PropTxtYear.Text = .PropStrYear
                .PropTxtMonth.Text = .PropStrMonth
            End With

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwCancelDone_ButtonClicked(sender As Object, e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vwCancelDone.ButtonClicked

        If e.Column = 4 Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            '「詳細」ボタンが押下された場合
            Dim frm As New EXTC0104
            Me.Hide()

            'パラメータにキャンセル待ちNOを設定
            frm.dataEXTC0104.PropStrYoyakuNo = vwCancelDone.ActiveSheet.Cells(e.Row, 5).Value

            '「キャンセル待ち登録」画面を表示
            frm.ShowDialog()

            '前回表示時の値でカレンダーを再表示する
            With dataEXTC0101
                .PropTxtYear.Text = .PropStrYear
                .PropTxtMonth.Text = .PropStrMonth
            End With

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        dataEXTC0101.PropStrStudioKbn = Nothing
        If logicEXTC0101.CheckSelectDateKariyoyaku(dataEXTC0101, arrDateLst, OPERATE_BUTTON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' 受付制御表の登録
        If logicEXTC0101.SetInsertUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        Dim frm As New EXTC0102
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTC0102.PropAryStrRiyoDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        dataEXTC0101.PropStrStudioKbn = Nothing
        If logicEXTC0101.CheckCancelWait(dataEXTC0101, arrDateLst, OPERATE_BUTTON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        ' 2015.11.19 ADD START↓ h.hagiwara
        ' キャンセル受付制御表の登録
        If logicEXTC0101.SetInsertCancelUketukeSeigyo(dataEXTC0101, arrDateLst) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If
        ' 2015.11.19 ADD END↑ h.hagiwara

        '画面遷移
        Dim frm As New EXTC0104
        Me.Hide()

        'パラメータに日付リストを設定
        frm.dataEXTC0104.PropAryStrCancelDate = New ArrayList
        frm.dataEXTC0104.PropAryStrCancelDate = arrDateLst

        '「仮予約登録」画面を表示
        frm.ShowDialog()

        '前回表示時の値でカレンダーを再表示する
        With dataEXTC0101
            .PropTxtYear.Text = .PropStrYear
            .PropTxtMonth.Text = .PropStrMonth
        End With

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        logicEXTC0101.MakeSpread(dataEXTC0101)

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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_ON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '祝日の設定
        For Each strDate In arrDateLst
            If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, strDate, HOLIYDAY_ON) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        If logicEXTC0101.CheckHoliday(dataEXTC0101, arrDateLst, HOLIYDAY_OFF) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '平日の設定
        For Each strDate In arrDateLst
            If logicEXTC0101.SetHolidayWeekday(dataEXTC0101, strDate, HOLIYDAY_OFF) = False Then
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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        dataEXTC0101.PropStrStudioKbn = Nothing
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_KYUKAN, OPERATE_BUTTON) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropHolmentKbn = HOLMENT_KBN_HOLIDAY

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)
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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        dataEXTC0101.PropStrStudioKbn = Nothing
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_MAINTE, OPERATE_BUTTON) = False Then
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
        frm.dataEXTZ0201.PropStrShisetuKbn = SHISETU_KBN_STUDIO
        frm.dataEXTZ0201.PropHolmentKbn = HOLMENT_KBN_MENTDAY

        '---「メンテ登録」画面を表示
        If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            logicEXTC0101.MakeSpread(dataEXTC0101)
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
    ''' <para>作成情報：2015/09/01 h.endo
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
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 1).ToString("00")
                If arrDateLst.Contains(strDate) = False Then
                    arrDateLst.Add(strDate)
                    intDayCount += 1
                End If
            End If
        Next
        For intCol As Integer = 0 To Me.vwCalendarSecond.ActiveSheet.Columns.Count - 1
            If Me.vwCalendarSecond.ActiveSheet.Cells(0, intCol).Value = "1" Then
                strDate = dataEXTC0101.PropStrYear & "/" & dataEXTC0101.PropStrMonth & "/" & (intCol + 16).ToString("00")
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
        dataEXTC0101.PropStrStudioKbn = Nothing
        If logicEXTC0101.CheckSelectDateKyukanMainteEigyo(dataEXTC0101, arrDateLst, SETDATE_EIGYO, OPERATE_BUTTON) = False Then
            If puErrMsg <> String.Empty Then
                MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            End If
            Return
        End If

        '営業日の設定
        For Each strDate In arrDateLst
            If logicEXTC0101.SetEigyo(dataEXTC0101, strDate) = False Then
                If puErrMsg <> String.Empty Then
                    MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                End If
                Return
            End If
        Next

    End Sub

#End Region

End Class