Imports Npgsql
Imports Common
Imports FarPoint.Win.Spread.CellType
Imports CommonEXT

Public Class EXTM0102

    Private sqlEXTM0102 As New SqlEXTM0102
    Private dataEXTM0102 As New DataEXTM0102
    Private logicEXTM0102 As New LogicEXTM0102
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    ''' <summary>
    ''' 付帯設備マスタメンテ、初期表示処理
    ''' </summary>
    ''' <remarks>現在の日付が含まれる期間をコンボボックスに表示。
    ''' その期間に対応する分類、付帯設備の表を取得し、表示する。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTM0102_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With dataEXTM0102
            .PropNewBtn = Me.rdoBtnNew
            .PropFinishedBtn = Me.rdoBtnFinished
            .PropFinishedFromTo = Me.cmbBoxFnishedFromTo
            .PropYearFrom = Me.txtYearFrom
            .PropMonthFrom = Me.txtMonthFrom
            .PropYearTo = Me.txtYearTo
            .PropMonthTo = Me.txtMonthTo
            .PropTheaterBtn = Me.rdoBtnTheater
            .PropStudioBtn = Me.rdoBtnStudio
            .PropVwGroupingSheet = Me.vwGroupingSheet
            .PropVwFutaiSheet = Me.vwFutaiSheet
            .PropBackBtn = Me.btnBack
            .PropNewEntryBtn = Me.btnNewEntry
            .PropEntryBtn = Me.btnEntry

        End With

        '分類表表示
        If logicEXTM0102.InitBunrui(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        '付帯シート記述,もし分類の表に何もなければ実行しない
        If logicEXTM0102.InitFutai(dataEXTM0102, 0) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, 0).Value
        ''取得したデータを付帯設備シートに代入
        If logicEXTM0102.SetFutaiData(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、fromテキストロストフォーカス
    ''' </summary>
    ''' <remarks>入力終了時、１桁のみの場合、頭に0を付ける
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub txtMonthFrom_LostFocus(sender As Object, e As EventArgs) Handles txtMonthFrom.LostFocus
        'テキストが空でなく、文字数が１なら頭に０を付ける
        If dataEXTM0102.PropMonthFrom.Text <> Nothing Then
            If dataEXTM0102.PropMonthFrom.Text.Length = 1 Then
                dataEXTM0102.PropMonthFrom.Text = "0" + dataEXTM0102.PropMonthFrom.Text
            End If
        End If
    End Sub


    ''' <summary>
    ''' 付帯設備マスタメンテ、Toテキストロストフォーカス時
    ''' </summary>
    ''' <remarks>入力終了時、１桁のみの場合、頭に0を付ける
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub txtMonthTo_LostFocus(sender As Object, e As EventArgs) Handles txtMonthTo.LostFocus
        'テキストが空でなく、文字数が１なら頭に０を付ける
        If dataEXTM0102.PropMonthTo.Text <> Nothing Then
            If dataEXTM0102.PropMonthTo.Text.Length = 1 Then
                dataEXTM0102.PropMonthTo.Text = "0" + dataEXTM0102.PropMonthTo.Text
            End If
        End If
    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、戻るボタン押下時処理
    ''' </summary>
    ''' <remarks>前画面に遷移する。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        Me.Close()
    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、表示内容を元に新規登録ボタン押下時処理
    ''' </summary>
    ''' <remarks>既に登録された内容を元に、新規登録を行う。
    ''' 新規に料金を設定するがチェックされ、対象期間コンボボックスが活性化。
    ''' 期間From、期間Toがそれぞれクリアされる。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnNewEntry_Click(sender As Object, e As EventArgs) Handles btnNewEntry.Click
        '期間をクリアし、対象期間を活性化
        logicEXTM0102.PushNewEntryBtn(dataEXTM0102)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、登録ボタン押下時処理
    ''' </summary>
    ''' <remarks>入力されたデータを元に、登録処理を行う。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub btnEntry_Click(sender As Object, e As EventArgs) Handles btnEntry.Click

        '変数宣言
        Dim noData As Integer = 0       '分類
        Dim noDataFutai As Integer = 0  '付帯設備

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェックをし、エラーがなければ登録
        '期間入力チェック
        If logicEXTM0102.CheckErrerKikan(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        '分類表入力チェック
        If logicEXTM0102.CheckErrerBunrui(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If

        '付帯設備表入力チェック
        If logicEXTM0102.CheckErrerFutai(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If

        ' エラーがなければデータ領域に格納    20151023
        If logicEXTM0102.RyokinInfSet(dataEXTM0102) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        'ダイアログ表示
        If MessageBox.Show(String.Format(EXTM0102_C0011, "付帯設備分類マスタ、付帯設備マスタ"), "登録確認", MessageBoxButtons.YesNo) = DialogResult.No Then
            Exit Sub
        End If

        If logicEXTM0102.InsertUpdateDB(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        ''分類表表示行数で、値が全て空欄のもの以外実行
        'For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1

        '    '変数を初期化
        '    noData = 0

        '    '空の行を判断
        '    For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
        '        If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then

        '        Else
        '            noData += 1
        '        End If
        '    Next

        '    '更新、もしくは挿入実行、全て空の行では実行しない
        '    If noData <> 14 Then
        '        If logicEXTM0102.EntryBunrui(i, dataEXTM0102, noData) = False Then
        '            MsgBox(puErrMsg)
        '            Exit Sub
        '        End If

        '        '付帯設備表表示行数で、値が全て空欄のもの以外実行
        '        For x = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
        '            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(x, 7).Value = i Then
        '                '変数を初期化
        '                noDataFutai = 0
        '                '値が
        '                For y = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 2
        '                    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(x, y).Value <> Nothing Then

        '                    Else
        '                        noDataFutai += 1
        '                    End If
        '                Next
        '                '更新、もしくは挿入実行、全て空の行では実行しない
        '                If noDataFutai <> 7 Then

        '                    If logicEXTM0102.EntryFutai(x, dataEXTM0102, noDataFutai) = False Then
        '                        MsgBox(puErrMsg)
        '                        Exit Sub
        '                    End If
        '                End If
        '            End If
        '        Next
        '    End If
        'Next

        ''付帯設備表表示行数で、値が全て空欄のもの以外実行
        'For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1

        '    '変数を初期化
        '    noDataFutai = 0

        '    '値が
        '    For j = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
        '        If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then

        '        Else
        '            noDataFutai += 1
        '        End If
        '    Next
        '    '更新、もしくは挿入実行、全て空の行では実行しない
        '    If noDataFutai <> 7 Then
        '        If logicEXTM0102.EntryFutai(i, dataEXTM0102, noDataFutai) = False Then
        '            MsgBox(puErrMsg)
        '            Exit Sub
        '        End If
        '    End If
        'Next

        MsgBox(String.Format(EXTM0102_C0012, "付帯設備分類マスタ、付帯設備マスタ"))
        '分類表表示

        dataEXTM0102.PropInitCmbFlg = False
        dataEXTM0102.PropInitFlg = False
        dataEXTM0102.PropFinishedFromTo.Enabled = True
        dataEXTM0102.PropNewEntryBtn.Enabled = True                                                     ' 2016.04.14 ADD h.hagiwara

        If logicEXTM0102.InitBunrui(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        '付帯シート記述,もし分類の表に何もなければ実行しない
        If logicEXTM0102.InitFutai(dataEXTM0102, 0) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If
        dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, 0).Value
        '取得したデータを付帯設備シートに代入
        If logicEXTM0102.SetFutaiData(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、対象期間選択処理
    ''' </summary>
    ''' <remarks>既に登録された期間を選択し、それを元に分類、付帯設備の表を取得する。
    ''' 表には、登録情報+５件の登録可能枠を表示
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub cmbBoxFnishedFromTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBoxFnishedFromTo.SelectedIndexChanged
        '初期表示以外で実行
        If dataEXTM0102.PropInitFlg = True Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If
            '分類シート記述
            If logicEXTM0102.InitBunrui(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            '付帯シート記述
            If logicEXTM0102.InitFutai(dataEXTM0102, 0) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, 0).Value
            '取得したデータを付帯設備シートに代入
            If logicEXTM0102.SetFutaiData(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、対象期間選択処理
    ''' </summary>
    ''' <remarks>既に登録された期間を選手入力時、それを元に分類、付帯設備の表を取得する。
    ''' 表には、登録情報+５件の登録可能枠を表示
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub cmbBoxFnishedFromTo_TextChange(sender As Object, e As EventArgs) Handles cmbBoxFnishedFromTo.TextChanged
        '変数宣言
        Dim txtFinishedFromTo As String

        '初期表示以外で実行
        If dataEXTM0102.PropInitFlg = True Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '21文字の入力で実行
            If dataEXTM0102.PropFinishedFromTo.Text.Length = 21 Then
                'コンボボックスのテキストを変数に代入(対象期間コードの形)
                txtFinishedFromTo = cmbBoxFnishedFromTo.Text.Substring(0, 10) + cmbBoxFnishedFromTo.Text.Substring(11, 10)
                'ボックスのテキストを選択
                dataEXTM0102.PropFinishedFromTo.SelectedValue = txtFinishedFromTo
            End If

            '分類シート記述
            If logicEXTM0102.InitBunrui(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            '付帯シート記述
            If logicEXTM0102.InitFutai(dataEXTM0102, 0) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, 0).Value
            '取得したデータを付帯設備シートに代入
            If logicEXTM0102.SetFutaiData(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、新規に料金を設定するラジオボタン押下時処理
    ''' </summary>
    ''' <remarks>分類、付帯設備、期間をクリアし、対象期間と表示内容を元に新規登録ボタンを非活性にする。
    ''' 分類、付帯設備の表には、５件の登録可能枠を表示
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoBtnNew_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBtnNew.CheckedChanged
        'クリア、非活性処理
        If dataEXTM0102.PropNewBtn.Checked = True And dataEXTM0102.PropNewEntryBtnFlg = False Then
            logicEXTM0102.CheckNewBtn(dataEXTM0102)
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、設定済みの料金を設定するラジオボタンチェック時処理
    ''' </summary>
    ''' <remarks>対象期間、表示内容を元に新規登録ボタンを活性化する。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoBtnFinished_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBtnFinished.CheckedChanged

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        'クリア、活性処理
        If dataEXTM0102.PropInitFlg = True Then
            logicEXTM0102.CheckFinishedBtn(dataEXTM0102)
        End If

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、施設（シアター）ラジオボタンチェック時処理
    ''' </summary>
    ''' <remarks>対象期間、施設をキーに、ＳＱＬを発行し、返す値に対応する分類、付帯設備の表を表示する。
    ''' 表示された表には、登録件数＋５行の登録枠を表示する。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub rdoBtnTheater_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBtnTheater.CheckedChanged


        '初期表示処理では実行しない
        If dataEXTM0102.PropInitFlg = True Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            dataEXTM0102.PropInitCmbFlg = False

            '分類表取得
            If logicEXTM0102.InitBunrui(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If

            '付帯シート記述,もし分類の表に何もなければ実行しない
            If logicEXTM0102.InitFutai(dataEXTM0102, 0) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            '取得したデータを付帯設備シートに代入
            dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, 0).Value
            If logicEXTM0102.SetFutaiData(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If

            For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
                If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).Value <> Nothing Then
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Locked = False
                Else
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Locked = True
                End If
                If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Value <> Nothing Then
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Locked = False
                Else
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Locked = True
                End If
                If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Value <> Nothing Then
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Locked = False
                Else
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Locked = True
                End If
                ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
                'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value <> Nothing Then
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Locked = False
                'Else
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Locked = True
                'End If
                'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value <> Nothing Then
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Locked = False
                'Else
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Locked = True
                'End If
                'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value <> Nothing Then
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Locked = False
                'Else
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Locked = True
                'End If
                ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
            Next
        End If
    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、分類表クリック時処理
    ''' </summary>
    ''' <remarks>クリックしたセルの分類コードをキーに、付帯設備の表を表示
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwGrouping_CellClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwGroupingSheet.CellClick

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '付帯設備表入力チェック
        If logicEXTM0102.CheckErrerFutai(dataEXTM0102) = False Then
            MsgBox(puErrMsg)
            Exit Sub
        End If

        ' エラーがなければデータ領域に格納    20151023
        If logicEXTM0102.RyokinInfSet(dataEXTM0102) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        'If logicEXTM0102.InitFutai(dataEXTM0102, e.Row) = False Then
        '    MsgBox(puErrMsg)
        'End If
        dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(e.Row, 0).Value
        '取得したデータを付帯設備シートに代入
        If logicEXTM0102.SetFutaiData(dataEXTM0102, e.Row) = False Then
            MsgBox(puErrMsg)
        End If
    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、分類表チェンジ時処理
    ''' </summary>
    ''' <remarks>データが入力された場合、行数を１追加,データ行数が減った場合、行数を１減らす
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwGroupingSheet_Change(sender As Object, e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwGroupingSheet.Change

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        'データが増減した場合、それに応じて行数を調整
        logicEXTM0102.ChangeGroupingInsartRow(dataEXTM0102)

        '科目コードコンボボックスの値が変化した時、細目コードコンボボックスを活性化
        If e.Column = 4 Then
            logicEXTM0102.ChangeCmbKamokuCd(e.Row, dataEXTM0102)
        End If

        '細目コードコンボボックスの値が変化した時、内訳コードコンボボックスを活性化
        If e.Column = 5 Then
            logicEXTM0102.ChangeCmbSaimokuCd(e.Row, dataEXTM0102)
        End If

        '内訳コードコンボボックスの値が変化した時、内訳コードコンボボックスを活性化
        If e.Column = 6 Then
            logicEXTM0102.ChangeCmbUchiwakeCd(e.Row, dataEXTM0102)
        End If

        ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
        ''借方科目コードコンボボックスの値が変化した時、借方細目コードコンボボックスを活性化
        'If e.Column = 8 Then
        '    logicEXTM0102.ChangeCmbKarikamokuCd(e.Row, dataEXTM0102)
        'End If

        ''借方科目コードコンボボックスの値が変化した時、借方細目コードコンボボックスを活性化
        'If e.Column = 9 Then
        '    logicEXTM0102.ChangeCmbKarisaimokuCd(e.Row, dataEXTM0102)
        'End If

        ''借方細目コードコンボボックスの値が変化した時、借方内訳コードコンボボックスを活性化
        'If e.Column = 10 Then
        '    logicEXTM0102.ChangeCmbKariuchiwakeCd(e.Row, dataEXTM0102)
        'End If
        ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、付帯設備表チェンジ時処理
    ''' </summary>
    ''' <remarks>データが入力された場合、行数を１追加,減った場合、行数を１減らす
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub vwFutaiSheet_Change(sender As Object, e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwFutaiSheet.Change
        '表示行数を調整
        logicEXTM0102.ChangeFutaiInsartRow(dataEXTM0102)
    End Sub

End Class