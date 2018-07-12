Imports System.Text
Imports Common
Imports FarPoint.Win.Spread
Imports Npgsql
Imports CommonEXT

Public Class EXTM0104
    'インスタンス生成
    Public dataEXTM0104 As New DataEXTM0104         'データクラス
    Private logicEXTM0104 As New LogicEXTM0104      'ロジッククラス
    Private sql As New SqlEXTM0104
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>利用料マスタのデータをDBから取得し、画面にセットする
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub EXTM0104_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
        dataEXTM0104.propFlg = True
        With dataEXTM0104
            .PropRdoshinki = Me.rdoshinki
            .PropRdosumi = Me.rdosumi
            .PropCmbKikan = Me.cmbkikan
            .PropTxtKikanFromYear = Me.txtkikanfromyear
            .PropTxtKikanFromMonth = Me.txtkikanfrommonth
            .PropTxtKikanToYear = Me.txtkikantoyear
            .PropTxtkikanToMonth = Me.txtkikantomonth
            .PropVwBunrui = Me.vwbunrui
            .PropVwBairitu = Me.vwbairitu
            .PropVwRyokin = Me.vwryokin

        End With
        Me.rdosumi.Checked = True

        If logicEXTM0104.InitForm(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return

        End If


        If dataEXTM0104.PropCmbKikan.SelectedValue = "0" Then
            Me.btnsinki.Enabled = False
        End If
        dataEXTM0104.propFlg = False

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub


    ''' <summary>
    ''' 設定済みの料金を編集するラジオボタンがチェックされた時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>設定済みの料金を編集するラジオボタンをチェックした時に、期間のコンボボックスと
    ''' 登録内容を元に新規登録ボタンを活性化する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub Rdosumi_CheckedChanged(sender As Object, e As EventArgs) Handles rdosumi.CheckedChanged

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '初期表示フラグ
        If dataEXTM0104.propFlg = False Then
            '「対象期間」を活性化する
            Me.cmbkikan.Enabled = True
            '「登録内容を元に新規」ボタンを活性化する
            Me.btnsinki.Enabled = True

        End If
        '初期表示を行う
        If logicEXTM0104.InitForm(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

    End Sub


    ''' <summary>
    ''' 新規に料金を設定するラジオボタンがチェックされた時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>新規に料金を設定するラジオボタンをチェックした時に、期間のコンボボックスと
    ''' 登録内容をもとに新規登録ボタンを非活性にする
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rdoshinki.CheckedChanged
        '「新規に料金を設定する」にチェックされている場合
        If Me.rdoshinki.Checked Then
            '「対象期間」を非活性にする
            Me.cmbkikan.Enabled = False
            '「登録内容を元に新規」ボタンを非活性にする
            Me.btnsinki.Enabled = False
            '期間FROM年、期間FROM月、期間TO年、期間TO月をクリアする
            Me.txtkikanfromyear.Clear()
            Me.txtkikanfrommonth.Clear()
            Me.txtkikantoyear.Clear()
            Me.txtkikantomonth.Clear()
            '「対象期間」の表示値をクリアする
            Me.cmbkikan.Text = ""
            '分類ビューをクリアする
            If logicEXTM0104.ClearVwBunrui(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '料金ビューをクリアする
            If logicEXTM0104.ClearVwRyokin(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '倍率ビューをクリアする()
            If logicEXTM0104.ClearVwBairitu(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If


        End If
    End Sub


    ''' <summary>
    ''' 利用料(分類)をクリックした時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>クリックした行の料金分類コードを条件に利用料(料金)を取得し、画面に設定する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles vwbunrui.CellClick

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 表示中の内容をチェックする          20151023
        If logicEXTM0104.RyokinInputCheck(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        ' エラーがなければデータ領域に格納    20151023
        If logicEXTM0104.RyokinInfSet(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        '検索用分類コードをセット
        dataEXTM0104.PropStrBunruicd = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(e.Row, 0).Text
        '検索用施設をセット
        If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(e.Row, 1).Text = "シアター" Then
            dataEXTM0104.PropStrShisetu = "1"
        ElseIf dataEXTM0104.PropVwBunrui.Sheets(0).Cells(e.Row, 1).Text = "スタジオ" Then
            dataEXTM0104.PropStrShisetu = "2"
        End If

        '料金データを取得する
        '20151023 If logicEXTM0104.GetRyokinData(dataEXTM0104, e.Row) = False Then
        '20151023     MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
        '20151023     Return
        '20151023 End If
        '料金ビューをクリアする
        If logicEXTM0104.ClearVwRyokin(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        '料金データをセットする
        If logicEXTM0104.SetRyokinData(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

    End Sub

    ''' <summary>
    ''' 登録内容を元に新規登録ボタンが押下した時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>ラジオボタンを新規に料金を設定するにチェックする、期間のコンボボックスを非活性にする
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnsinki.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '新規に料金を設定するにチェック
        Me.rdoshinki.Checked = True
        '「対象期間」を非活性にする
        Me.cmbkikan.Enabled = False
        '期間FROM年、期間FROM月、期間TO年、期間TO月をクリアする
        Me.txtkikanfromyear.Clear()
        Me.txtkikanfrommonth.Clear()
        Me.txtkikantoyear.Clear()
        Me.txtkikantomonth.Clear()
        'ラジオボタンが変わったため分類ビュー、料金ビュー、倍率ビューを再セットする必要がある
        If logicEXTM0104.SetBunruiData(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        '検索用分類コードをセット
        dataEXTM0104.PropStrBunruicd = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(0, 0).Text
        '検索用施設をセット
        If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(0, 1).Text = "シアター" Then
            dataEXTM0104.PropStrShisetu = "1"
        ElseIf dataEXTM0104.PropVwBunrui.Sheets(0).Cells(0, 1).Text = "スタジオ" Then
            dataEXTM0104.PropStrShisetu = "2"
        End If
        If logicEXTM0104.SetRyokinData(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        If logicEXTM0104.SetBairituData(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

    End Sub


    ''' <summary>
    ''' 期間コンボボックスの内容が変化した場合の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>変化した値を元に、利用料(分類)、利用料(倍率)を再取得、画面に表示する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub kikan_comb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbkikan.TextChanged
        '「設定済みの料金を編集する」ラジオボタンにチェックしている場合
        If Me.rdosumi.Checked Then

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            '対象期間が空白以外を選択していた場合
            If dataEXTM0104.PropCmbKikan.Text <> "" Then
                '期間FROM年、期間FROM月、期間TO年、期間TO月を設定する
                If logicEXTM0104.SetTxtKikan(dataEXTM0104) = False Then
                    MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                    Return
                End If
            Else
                '期間FROM年、期間FROM月、期間TO年、期間TO月をクリアする
                'dataEXTM0104.PropTxtKikanFromYear.Clear()
                'dataEXTM0104.PropTxtKikanFromMonth.Clear()
                'dataEXTM0104.PropTxtKikanToYear.Clear()
                'dataEXTM0104.PropTxtkikanToMonth.Clear()
                dataEXTM0104.PropTxtKikanFromYear.Text = ""
                dataEXTM0104.PropTxtKikanFromMonth.Text = ""
                dataEXTM0104.PropTxtKikanToYear.Text = ""
                dataEXTM0104.PropTxtkikanToMonth.Text = ""
            End If

            '分類ビューをクリアする
            If logicEXTM0104.ClearVwBunrui(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '料金ビューをクリアする
            If logicEXTM0104.ClearVwRyokin(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '倍率ビューをクリアする
            If logicEXTM0104.ClearVwBairitu(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If

            '分類データを取得する
            If logicEXTM0104.GetBunruiData(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '分類データをセットする
            If logicEXTM0104.SetBunruiData(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '料金データを取得する
            If logicEXTM0104.GetRyokinData(dataEXTM0104, 0) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '料金データをセットする
            If logicEXTM0104.SetRyokinData(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '倍率データを取得する
            If logicEXTM0104.GetBairituData(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
            '倍率データをセットする
            If logicEXTM0104.SetBairituData(dataEXTM0104) = False Then
                MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
                Return
            End If
        End If

    End Sub



    ''' <summary>
    ''' 登録ボタンが押下した時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>入力チェックし、登録/更新処理を行う
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub insert_bt_Click(sender As Object, e As EventArgs) Handles btninsert.Click

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        '入力チェック
        If logicEXTM0104.KikanInputCheck(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        If logicEXTM0104.BunruiInputCheck(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        If logicEXTM0104.RyokinInputCheck(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        If logicEXTM0104.BairituInputCheck(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        ' エラーがなければデータ領域に格納    20151023
        If logicEXTM0104.RyokinInfSet(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If

        '確認ダイアログ出力
        Dim result As DialogResult = MessageBox.Show(String.Format(c0011, "利用料マスタ"), _
                                           "登録確認", _
                                           MessageBoxButtons.YesNo, _
                                           MessageBoxIcon.Exclamation, _
                                           MessageBoxDefaultButton.Button2)

        If result = DialogResult.No Then
            Return
        End If

        'システム日付取得（SELECT）
        If logicEXTM0104.GetSelectSysDate(dataEXTM0104) = False Then
            Return
        End If

        '登録/更新処理
        If logicEXTM0104.BtClickBunrui(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        If logicEXTM0104.BtClickRyokin(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        If logicEXTM0104.BtClickBairitu(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
        '完了メッセージを出力
        MsgBox(String.Format(I0002, "利用料マスタの登録/更新"))

        If logicEXTM0104.InitForm(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return

        End If

    End Sub

    Private Sub Back_bt_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Me.Close()

    End Sub

    ''' <summary>
    ''' 利用料(分類)に表示する行数制御
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示行数を常にデータ数+５に制御する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub FpSpread1_TextChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwbunrui.Change
        Dim dtct As Integer
        Dim ct As Integer = 0

        '最後の入力がある行を探す
        For index = 0 To dataEXTM0104.PropVwBunrui.Sheets(0).RowCount - 1
            ct = 0
            For i = 0 To 4
                If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, i).Text <> Nothing Then
                    ct += 1
                End If
            Next
            If ct <> 0 Then
                dtct = index
            End If
        Next
        '最後の入力がある行の行数+６
        dataEXTM0104.PropVwBunrui.Sheets(0).RowCount = dtct + 6

    End Sub

    ''' <summary>
    ''' 利用料(料金)に表示する行数制御
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示行数を常にデータ数+５に制御する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub ryokin_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwryokin.Change
        Dim dtct As Integer
        Dim ct As Integer = 0
        '最後の入力がある行を探す
        For index = 0 To dataEXTM0104.PropVwRyokin.Sheets(0).RowCount - 1
            ct = 0
            For i = 0 To 4
                If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, i).Text <> Nothing Then
                    ct += 1
                End If
            Next
            If ct = 5 Then
                dtct = index
            End If
        Next
        '最後の入力がある行の行数+６
        dataEXTM0104.PropVwRyokin.Sheets(0).RowCount = dtct + 6

    End Sub

    ''' <summary>
    ''' 利用料(倍率)に表示する行数制御
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>表示行数を常にデータ数+５に制御する
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>

    Private Sub bairitu_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles vwbairitu.Change
        Dim dtct As Integer
        Dim ct As Integer = 0
        '最後の入力がある行を探す
        For index = 0 To dataEXTM0104.PropVwBairitu.Sheets(0).RowCount - 1
            ct = 0
            For i = 0 To 4
                If dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, i).Text <> Nothing Then
                    ct += 1
                End If
            Next
            If ct = 5 Then
                dtct = index
            End If
        Next
        '最後の入力がある行の行数+６
        dataEXTM0104.PropVwBairitu.Sheets(0).RowCount = dtct + 6


    End Sub
    ''' <summary>
    '''期間(FROM)月の表示制御
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>期間(FROM)月に入力した数字が1桁の場合は0を前に追加して出力
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub txtkikanfrommonth_TextChanged(sender As Object, e As EventArgs) Handles txtkikanfrommonth.LostFocus
        '期間(FROM)月に入力した数字が1桁の場合は0を前に追加するロジック
        If logicEXTM0104.TxtChangeKikanFromMonth(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
    End Sub
    ''' <summary>
    '''期間(TO)月の表示制御
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>期間(TO)月に入力した数字が1桁の場合は0を前に追加して出力
    ''' <para>作成情報：2015/08/20 mori
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub txtkikantomonth_TextChanged(sender As Object, e As EventArgs) Handles txtkikantomonth.LostFocus
        '期間(TO)月に入力した数字が1桁の場合は0を前に追加するロジック
        If logicEXTM0104.TxtChangeKikanToMonth(dataEXTM0104) = False Then
            MsgBox(puErrMsg, MsgBoxStyle.Critical, "エラー")
            Return
        End If
    End Sub
End Class