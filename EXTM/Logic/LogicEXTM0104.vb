Imports Common
Imports Npgsql
Imports FarPoint.Win.Spread

''' <summary>
''' 利用料金マスタメンテ
''' </summary>
''' <remarks></remarks>
Public Class LogicEXTM0104

    Public SqlEXTM0104 As New SqlEXTM0104
    Public commonlogic As New CommonLogic

    Private Const M0104_RYOKIN_COL_BUNRUICD As Integer = 0
    Private Const M0104_RYOKIN_COL_RYOKINCD As Integer = 1
    Private Const M0104_RYOKIN_COL_RYOKINNM As Integer = 2
    Private Const M0104_RYOKIN_COL_RYOKINHOUR As Integer = 3
    Private Const M0104_RYOKIN_COL_RYOKIN As Integer = 4
    Private Const M0104_RYOKIN_COL_SORT As Integer = 5
    Private Const M0104_RYOKIN_COL_STS As Integer = 6
    Private Const M0104_RYOKIN_COL_SHISETU As Integer = 7
    Private Const M0104_RYOKIN_COL_UPKBN As Integer = 8

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>初期表示を作成
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitForm(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            '期間取得
            If GetKikanData(dataEXTM0104) = False Then
                Return False
            End If
            '期間のセット
            If SetKikanData(dataEXTM0104) = False Then
                Return False
            End If
            '分類取得
            If GetBunruiData(dataEXTM0104) = False Then
                Return False
            End If
            '画面の利用料(分類)をクリアする
            If ClearVwBunrui(dataEXTM0104) = False Then
                Return False
            End If
            '分類のセット
            If SetBunruiData(dataEXTM0104) = False Then
                Return False
            End If

            dataEXTM0104.PropStrBunruicd = dataEXTM0104.PropVwBunrui.Sheets(0).ActiveCell.Text
            '料金取得
            If GetRyokinData(dataEXTM0104, 0) = False Then
                Return False
            End If

            '画面の利用料(料金)をクリアする
            If ClearVwRyokin(dataEXTM0104) = False Then
                Return False
            End If

            '料金セット
            If SetRyokinData(dataEXTM0104) = False Then
                Return False
            End If

            '倍率取得
            If GetBairituData(dataEXTM0104) = False Then
                Return False
            End If

            '画面の利用料(分類)をクリアする
            If ClearVwBairitu(dataEXTM0104) = False Then
                Return False
            End If

            '倍率のセット
            If SetBairituData(dataEXTM0104) = False Then
                Return False
            End If

            '初期表示の料金ビューの行数を5行に設定
            '  dataEXTM0104.PropVwRyokin.Sheets(0).RowCount = 5
            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 分類取得
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>分類データをDBから取得
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetBunruiData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数を宣言
        Dim coon As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'データテーブル

        Try
            '検索用期間をセット
            dataEXTM0104.PropStrKikan = dataEXTM0104.PropCmbKikan.SelectedValue.ToString

            'コネクションを開く
            coon.Open()
            'sql文を設定
            If SqlEXTM0104.GetBunruiData(Adapter, coon, dataEXTM0104) = False Then
                Return False
            End If
            'sqlを実行
            Adapter.Fill(Table)
            'DBに問い合わせた値を格納する
            dataEXTM0104.PropDtBunrui = Table

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            coon.Close()
            coon.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 分類データセット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>DBから取得分類データを画面にセット
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetBunruiData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        Dim datact As Integer = dataEXTM0104.PropDtBunrui.Rows.Count

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try

            '分類ビューの表示行数をデータ個数+5行にする
            dataEXTM0104.PropVwBunrui.Sheets(0).RowCount = datact + 5

            'DB問い合わせた値を分類ビューに設定
            For i = 0 To datact - 1
                For index = 0 To 4
                    dataEXTM0104.PropVwBunrui.Sheets(0).Cells(i, index).Value = dataEXTM0104.PropDtBunrui.Rows(i)(index)
                Next

                '分類CDの列をロックする
                dataEXTM0104.PropVwBunrui.Sheets(0).Cells(i, 0).Locked = True
                '施設コンボボックスの列をロックする
                dataEXTM0104.PropVwBunrui.Sheets(0).Cells(i, 1).Locked = True
            Next

            '施設コンボボックスの設定
            Dim cmb As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
            'cmb.Items = New String() {"シアター", "スタジオ"}                      ' 2015.12.01 UPD h.hagiwara
            cmb.Items = New String() {"", "シアター", "スタジオ"}                   ' 2015.12.01 UPD h.hagiwara

            '行数分のコンボボックスを設定
            For index = 0 To datact + 4
                dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 1).CellType = cmb
            Next
            'DBの値によってコンボボックスの初期表示する値を設定
            For index = 0 To datact - 1

                If dataEXTM0104.PropDtBunrui.Rows(index)(1) = "2" Then
                    dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 1).Text = "スタジオ"
                End If
                If dataEXTM0104.PropDtBunrui.Rows(index)(1) = "1" Then
                    dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 1).Text = "シアター"
                End If

            Next
            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try
    End Function

    ''' <summary>
    ''' 期間取得
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>期間データをDBから取得
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetKikanData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数を宣言
        Dim coon As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'データテーブル

        Try
            'コネクションを開く
            coon.Open()
            'sql文を設定する
            If SqlEXTM0104.GetKikanData(Adapter, coon, dataEXTM0104) = False Then
                Return False
            End If
            'sqlを実行する
            Adapter.Fill(Table)
            '期間の表示値をキーに、期間のコードをバリューに設定するハッシュテーブルを宣言
            dataEXTM0104.PropHtKikan = New Hashtable
            'DB問い合わせた値を格納する
            dataEXTM0104.PropDtKikan = Table

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        Finally
            coon.Close()
            coon.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 期間データセット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>DBから取得期間データを利用料マスタメンテ画面にセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetKikanData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim datact As Integer = dataEXTM0104.PropDtKikan.Rows.Count
        Dim showid As String = "0"
        Dim showlabel As String = ""

        Try
            For index = 0 To datact - 1
                '期間(FROM）を切り出す
                Dim f As DateTime = DateTime.Parse(dataEXTM0104.PropDtKikan.Rows(index)(0).Substring(0, 10))
                '期間（TO）を切り出す
                Dim t As DateTime = DateTime.Parse(dataEXTM0104.PropDtKikan.Rows(index)(0).Substring(10, 10))
                '現在の日付を含む期間をコンボボックスの初期表示に設定する
                If (Now > f AndAlso Now < t) Then
                    showid = dataEXTM0104.PropDtKikan.Rows(index)(0)
                    showlabel = dataEXTM0104.PropDtKikan.Rows(index)(1)
                End If

            Next

            If showid <> "0" Then
                dataEXTM0104.PropTxtKikanFromYear.Text = showid.Substring(0, 4)
                dataEXTM0104.PropTxtKikanFromMonth.Text = showid.Substring(5, 2)
                dataEXTM0104.PropTxtKikanToYear.Text = showid.Substring(10, 4)
                dataEXTM0104.PropTxtkikanToMonth.Text = showid.Substring(15, 2)
                If commonlogic.SetCmbBox(dataEXTM0104.PropDtKikan, dataEXTM0104.PropCmbKikan, False, showid, showlabel) = False Then
                    Return False

                End If
            Else
                If commonlogic.SetCmbBox(dataEXTM0104.PropDtKikan, dataEXTM0104.PropCmbKikan) = False Then
                    Return False
                End If
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金取得
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>料金データをDBから取得
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetRyokinData(ByRef dataEXTM0104 As DataEXTM0104, ByVal rowindex As Integer) As Boolean
        '開始ログ
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数を宣言
        Dim coon As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'データテーブル

        Try
            '検索用期間をセット
            If dataEXTM0104.PropCmbKikan.SelectedValue.ToString <> "" Then
                dataEXTM0104.PropStrKikan = dataEXTM0104.PropCmbKikan.SelectedValue.ToString
            End If

            '検索用分類コードをセット
            'dataEXTM0104.PropStrBunruicd = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(dataEXTM0104.PropVwBunrui.Sheets(0).ActiveRowIndex, 0).Text
            dataEXTM0104.PropStrBunruicd = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(rowindex, 0).Text

            '検索用施設をセット
            'If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(dataEXTM0104.PropVwBunrui.Sheets(0).ActiveRowIndex, 1).Text = "シアター" Then
            '    dataEXTM0104.PropStrShisetu = "1"
            'ElseIf dataEXTM0104.PropVwBunrui.Sheets(0).Cells(dataEXTM0104.PropVwBunrui.Sheets(0).ActiveRowIndex, 1).Text = "スタジオ" Then
            '    dataEXTM0104.PropStrShisetu = "2"
            'End If
            If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(rowindex, 1).Text = "シアター" Then
                dataEXTM0104.PropStrShisetu = "1"
            ElseIf dataEXTM0104.PropVwBunrui.Sheets(0).Cells(rowindex, 1).Text = "スタジオ" Then
                dataEXTM0104.PropStrShisetu = "2"
            End If

            'コネクションを開く
            coon.Open()

            'sql文を設定する
            If SqlEXTM0104.GetRyokinData(Adapter, coon, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            Adapter.Fill(Table)
            'DB問い合わせた値を格納する
            dataEXTM0104.PropDtRyoukin = Table

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        Finally
            coon.Close()
            coon.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金データセット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>DBから取得料金データを利用料マスタメンテ画面にセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetRyokinData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '20151023 Dim datact As Integer = dataEXTM0104.PropDtRyoukin.Rows.Count
        'Dim datact As Integer
        Dim dtrow1 As DataRow
        Dim dtrow2 As DataRow

        Try
            If dataEXTM0104.PropDtRyokinDsp Is Nothing Then
                dataEXTM0104.PropDtRyokinDsp = dataEXTM0104.PropDtRyoukin.Clone
            Else
                dataEXTM0104.PropDtRyokinDsp.Clear()
            End If
            If dataEXTM0104.PropDtRyokinEscp Is Nothing Then
                dataEXTM0104.PropDtRyokinEscp = dataEXTM0104.PropDtRyoukin.Clone
            Else
                dataEXTM0104.PropDtRyokinEscp.Clear()
            End If

            ' 格納領域を表示用・退避用に分けて格納する。
            For j = 0 To dataEXTM0104.PropDtRyoukin.Rows.Count - 1
                If dataEXTM0104.PropDtRyoukin.Rows(j).Item(M0104_RYOKIN_COL_BUNRUICD) = dataEXTM0104.PropStrBunruicd And _
                   dataEXTM0104.PropDtRyoukin.Rows(j).Item(M0104_RYOKIN_COL_SHISETU) = dataEXTM0104.PropStrShisetu Then
                    dtrow1 = dataEXTM0104.PropDtRyokinDsp.NewRow
                    For k = 0 To 8
                        dtrow1(k) = dataEXTM0104.PropDtRyoukin.Rows(j).Item(k)
                    Next
                    dataEXTM0104.PropDtRyokinDsp.Rows.Add(dtrow1)
                Else
                    dtrow2 = dataEXTM0104.PropDtRyokinEscp.NewRow
                    For k = 0 To 8
                        dtrow2(k) = dataEXTM0104.PropDtRyoukin.Rows(j).Item(k)
                    Next
                    dataEXTM0104.PropDtRyokinEscp.Rows.Add(dtrow2)
                End If
            Next

            '料金ビューの行数をデータ個数+5行に設定する
            'dataEXTM0104.PropVwRyokin.Sheets(0).RowCount = datact + 5
            dataEXTM0104.PropVwRyokin.Sheets(0).RowCount = dataEXTM0104.PropDtRyokinDsp.Rows.Count + 5

            'リストにDBから取得した値を画面にセット
            For i As Integer = 0 To dataEXTM0104.PropDtRyokinDsp.Rows.Count - 1
                For index = 0 To 8
                    'dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, index).Value = dataEXTM0104.PropDtRyoukin.Rows(i).Item(index)
                    dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, index).Value = dataEXTM0104.PropDtRyokinDsp.Rows(i).Item(index)
                Next
                '分類CDと料金CDをロックする
                dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, 0).Locked = True
                dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, 1).Locked = True
            Next

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 倍率取得
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>倍率データをDBから取得
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetBairituData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim coon As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'データテーブル

        Try
            '検索用の期間を取得する
            dataEXTM0104.PropStrKikan = dataEXTM0104.PropCmbKikan.SelectedValue.ToString

            'コネクションを開く
            coon.Open()
            'sqlを設定する
            If SqlEXTM0104.GetBairituData(Adapter, coon, dataEXTM0104) = False Then
                Return False
            End If
            'sqlを実行する
            Adapter.Fill(Table)

            'DB問い合わせたデータを格納する
            dataEXTM0104.PropDtBairitu = Table

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        Finally
            coon.Close()
            coon.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 倍率データセット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>DBから取得倍率データを利用料マスタメンテ画面にセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetBairituData(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '行数を格納する変数を宣言
        Dim datact As Integer = dataEXTM0104.PropDtBairitu.Rows.Count

        Try
            '倍率ビューの表示行数をデータ個数+5行に設定する
            dataEXTM0104.PropVwBairitu.Sheets(0).RowCount = datact + 5

            'リストにDBから取得した値を格納する
            For i As Integer = 0 To datact - 1
                For index = 0 To 5
                    dataEXTM0104.PropVwBairitu.Sheets(0).Cells(i, index).Value = dataEXTM0104.PropDtBairitu.Rows(i)(index)
                Next
                '倍率CDをロックする
                dataEXTM0104.PropVwBairitu.Sheets(0).Cells(i, 0).Locked = True
                '施設をロックする
                dataEXTM0104.PropVwBairitu.Sheets(0).Cells(i, 1).Locked = True
            Next
            'コンボボックスを設定する
            Dim cmb As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
            'cmb.Items = New String() {"シアター", "スタジオ"}                      ' 2015.12.01 UPD h.hagiwara
            cmb.Items = New String() {"", "シアター", "スタジオ"}                   ' 2015.12.01 UPD h.hagiwara

            For index = 0 To datact + 4
                dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 1).CellType = cmb
            Next

            'コンボボックスの表示
            For index = 0 To datact - 1

                If dataEXTM0104.PropDtBairitu.Rows(index)(1) = "2" Then
                    dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 1).Text = "スタジオ"
                End If
                If dataEXTM0104.PropDtBairitu.Rows(index)(1) = "1" Then
                    dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 1).Text = "シアター"
                End If

            Next
            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try
    End Function

    ''' <summary>
    ''' 分類の登録/更新の分岐処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値を登録/更新用の変数に代入し、登録するか更新するかの分岐を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function BtClickBunrui(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        '分類マスタの登録/更新
        Try

            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            Dim datact As Integer = dataEXTM0104.PropVwBunrui.Sheets(0).RowCount
            Dim olddatact As Integer = dataEXTM0104.PropDtBunrui.Rows.Count
            For index = 0 To datact - 1

                'insert用に変数に値をセットする

                '期間FROM
                dataEXTM0104.PropStrKikanFrom = dataEXTM0104.PropTxtKikanFromYear.Text + "/" + dataEXTM0104.PropTxtKikanFromMonth.Text + "/01"
                '期間TO
                Dim t As DateTime = dataEXTM0104.PropTxtKikanToYear.Text + "/" + dataEXTM0104.PropTxtkikanToMonth.Text
                dataEXTM0104.PropStrKikanTo = t.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
                '分類CD
                dataEXTM0104.PropStrBunruicdInsert = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 0).Text
                '施設
                Dim tmp As String = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 1).Text
                If tmp = "スタジオ" Then
                    dataEXTM0104.PropStrShisetu = "2"
                End If
                If tmp = "シアター" Then
                    dataEXTM0104.PropStrShisetu = "1"
                End If
                '料金分類名
                dataEXTM0104.PropStrBunruinm = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 2).Text
                '並び順

                dataEXTM0104.PropStrSort = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 3).Text
                'ステータス
                If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 4).Value Then
                    '「無効」がチェックされていたら、利用不可
                    dataEXTM0104.PropStrSts = "1"
                Else
                    dataEXTM0104.PropStrSts = "0"
                End If

                If dataEXTM0104.PropRdoshinki.Checked Then
                    '新規の場合、入力された行は全てインサート
                    If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 0).Text <> Nothing Then
                        If InsertBunrui(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    End If
                ElseIf dataEXTM0104.PropRdosumi.Checked Then
                    'DB問い合わせたデータの個数分はアップデート
                    If index <= olddatact - 1 Then
                        If UpdataBunrui(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    End If
                    'DB問い合わせたデータの個数以上のデータはインサート
                    If index > olddatact - 1 Then
                        If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 0).Text <> Nothing Then
                            If InsertBunrui(dataEXTM0104, Cn) = False Then
                                'ロールバック
                                If Tsx IsNot Nothing Then
                                    Tsx.Rollback()
                                End If
                                Return False
                            End If
                        End If
                    End If
                End If
            Next

            'コミット
            Tsx.Commit()

            If GetBunruiData(dataEXTM0104) = False Then
                Return False
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金の登録/更新の分岐処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値を登録/更新用の変数に代入し、登録するか更新するかの分岐を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function BtClickRyokin(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        '料金マスタの登録/更新
        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            Dim datact As Integer = dataEXTM0104.PropVwRyokin.Sheets(0).RowCount
            Dim olddatact As Integer = dataEXTM0104.PropDtRyoukin.Rows.Count
            'For index = 0 To datact - 1

            '    'insert用に変数に値をセットする

            '    '期間FROM
            '    dataEXTM0104.PropStrKikanFrom = dataEXTM0104.PropTxtKikanFromYear.Text + "/" + dataEXTM0104.PropTxtKikanFromMonth.Text + "/01"
            '    '期間TO
            '    Dim t As DateTime = dataEXTM0104.PropTxtKikanToYear.Text + "/" + dataEXTM0104.PropTxtkikanToMonth.Text
            '    dataEXTM0104.PropStrKikanTo = t.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

            '    '施設
            '    Dim shisetu As String = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(dataEXTM0104.PropVwBunrui.ActiveSheet.ActiveRowIndex, 1).Text
            '    If shisetu = "シアター" Then
            '        dataEXTM0104.PropStrShisetu = "1"
            '    End If
            '    If shisetu = "スタジオ" Then
            '        dataEXTM0104.PropStrShisetu = "2"
            '    End If

            '    '料金分類CD
            '    dataEXTM0104.PropStrBunruicdInsert = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 0).Text
            '    '料金CD
            '    dataEXTM0104.PropStrRyokincd = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 1).Text
            '    '料金名
            '    dataEXTM0104.PropStrRyokinnm = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 2).Text
            '    '料金時間貸し
            '    dataEXTM0104.PropStrRyokinhour = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 3).Text
            '    '料金
            '    dataEXTM0104.PropStrRyokin = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 4).Text
            '    '並び順
            '    dataEXTM0104.PropStrSort = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 5).Text
            '    'ステータス
            '    If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 6).Value Then
            '        '「無効」がチェックされていたら、利用不可
            '        dataEXTM0104.PropStrSts = "1"
            '    Else
            '        dataEXTM0104.PropStrSts = "0"
            '    End If
            '    If dataEXTM0104.PropRdoshinki.Checked Then
            '        '新規の場合、入力された行は全てインサート
            '        If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 0).Text <> Nothing Then
            '            If InsertRyokin(dataEXTM0104, Cn) = False Then
            '                'ロールバック
            '                If Tsx IsNot Nothing Then
            '                    Tsx.Rollback()
            '                End If
            '                Return False
            '            End If
            '        End If
            '    ElseIf dataEXTM0104.PropRdosumi.Checked Then
            '        'DB問い合わせたデータの個数分はアップデート
            '        If index <= olddatact - 1 Then
            '            If UpdataRyokin(dataEXTM0104, Cn) = False Then
            '                'ロールバック
            '                If Tsx IsNot Nothing Then
            '                    Tsx.Rollback()
            '                End If
            '                Return False
            '            End If
            '        End If
            '        'DB問い合わせたデータの個数以上のデータはインサート
            '        If index > olddatact - 1 Then
            '            If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 0).Text <> Nothing Then
            '                If InsertRyokin(dataEXTM0104, Cn) = False Then
            '                    'ロールバック
            '                    If Tsx IsNot Nothing Then
            '                        Tsx.Rollback()
            '                    End If
            '                    Return False
            '                End If
            '            End If
            '        End If
            '    End If
            'Next
            For index = 0 To dataEXTM0104.PropDtRyoukin.Rows.Count - 1

                'insert用に変数に値をセットする

                '期間FROM
                dataEXTM0104.PropStrKikanFrom = dataEXTM0104.PropTxtKikanFromYear.Text + "/" + dataEXTM0104.PropTxtKikanFromMonth.Text + "/01"
                '期間TO
                Dim t As DateTime = dataEXTM0104.PropTxtKikanToYear.Text + "/" + dataEXTM0104.PropTxtkikanToMonth.Text
                dataEXTM0104.PropStrKikanTo = t.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

                '施設
                'Dim shisetu As String = dataEXTM0104.PropVwBunrui.Sheets(0).Cells(dataEXTM0104.PropVwBunrui.ActiveSheet.ActiveRowIndex, 1).Text
                'If shisetu = "シアター" Then
                '    dataEXTM0104.PropStrShisetu = "1"
                'End If
                'If shisetu = "スタジオ" Then
                '    dataEXTM0104.PropStrShisetu = "2"
                'End If
                dataEXTM0104.PropStrShisetu = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_SHISETU)

                '料金分類CD
                dataEXTM0104.PropStrBunruicdInsert = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_BUNRUICD)
                '料金CD
                dataEXTM0104.PropStrRyokincd = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_RYOKINCD)
                '料金名
                dataEXTM0104.PropStrRyokinnm = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_RYOKINNM)
                '料金時間貸し
                dataEXTM0104.PropStrRyokinhour = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_RYOKINHOUR)
                '料金
                dataEXTM0104.PropStrRyokin = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_RYOKIN)
                '並び順
                dataEXTM0104.PropStrSort = dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_SORT)
                'ステータス
                If dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_STS) Then
                    '「無効」がチェックされていたら、利用不可
                    dataEXTM0104.PropStrSts = "1"
                Else
                    dataEXTM0104.PropStrSts = "0"
                End If

                If dataEXTM0104.PropRdoshinki.Checked Then
                    '新規の場合、入力された行は全てインサート
                    If InsertRyokin(dataEXTM0104, Cn) = False Then
                        'ロールバック
                        If Tsx IsNot Nothing Then
                            Tsx.Rollback()
                        End If
                        Return False
                    End If
                ElseIf dataEXTM0104.PropRdosumi.Checked Then
                    If dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_UPKBN) = "0" Then
                        If UpdataRyokin(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    ElseIf dataEXTM0104.PropDtRyoukin.Rows(index).Item(M0104_RYOKIN_COL_UPKBN) = "1" Then
                        If InsertRyokin(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    End If

                End If
            Next

            'コミット
            Tsx.Commit()

            If GetRyokinData(dataEXTM0104, 0) = False Then
                Return False
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 倍率の登録/更新の分岐処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値を登録/更新用の変数に代入し、登録するか更新するかの分岐を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function BtClickBairitu(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション
        Dim olddatact As Integer

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '倍率マスタの登録/更新
            Dim datact As Integer = dataEXTM0104.PropVwBairitu.Sheets(0).RowCount

            olddatact = dataEXTM0104.PropDtBairitu.Rows.Count


            For index = 0 To datact - 1

                'insert用に変数に値をセットする

                '期間FROM
                dataEXTM0104.PropStrKikanFrom = dataEXTM0104.PropTxtKikanFromYear.Text + "/" + dataEXTM0104.PropTxtKikanFromMonth.Text + "/01"
                '期間TO
                Dim t As DateTime = dataEXTM0104.PropTxtKikanToYear.Text + "/" + dataEXTM0104.PropTxtkikanToMonth.Text
                dataEXTM0104.PropStrKikanTo = t.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

                '倍率CD
                dataEXTM0104.PropStrBairitucd = dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 0).Text

                '施設
                Dim shisetu As String = dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 1).Text
                If shisetu = "シアター" Then
                    dataEXTM0104.PropStrShisetu = "1"
                End If
                If shisetu = "スタジオ" Then
                    dataEXTM0104.PropStrShisetu = "2"
                End If
                '倍率名
                dataEXTM0104.PropStrbairitunm = dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 2).Text
                '倍率
                dataEXTM0104.propStrBairitu = dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 3).Text
                '並び順
                dataEXTM0104.PropStrSort = dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 4).Text

                'ステータス
                If dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 5).Value Then
                    '「無効」がチェックされていたら、利用不可
                    dataEXTM0104.PropStrSts = "1"
                Else
                    dataEXTM0104.PropStrSts = "0"
                End If

                If dataEXTM0104.PropRdoshinki.Checked Then
                    '新規の場合、入力された行は全てインサート
                    If dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 0).Text <> Nothing Then
                        If InsertBairitu(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    End If
                ElseIf dataEXTM0104.PropRdosumi.Checked Then
                    'DB問い合わせたデータの個数分はアップデート
                    If index <= olddatact - 1 Then
                        If UpdataBairitu(dataEXTM0104, Cn) = False Then
                            'ロールバック
                            If Tsx IsNot Nothing Then
                                Tsx.Rollback()
                            End If
                            Return False
                        End If
                    End If
                    'DB問い合わせたデータの個数以上のデータはインサート
                    If index > olddatact - 1 Then
                        If dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 0).Text <> Nothing Then
                            If InsertBairitu(dataEXTM0104, Cn) = False Then
                                'ロールバック
                                If Tsx IsNot Nothing Then
                                    Tsx.Rollback()
                                End If
                                Return False
                            End If
                        End If
                    End If
                End If
            Next

            'コミット
            Tsx.Commit()

            If GetBairituData(dataEXTM0104) = False Then
                Return False
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 分類の更新処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに分類の更新処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function UpdataBunrui(ByRef dataEXTM0104 As DataEXTM0104, _
                                 ByVal Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.UpdataBunrui(cmd, Cn, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            cmd.ExecuteNonQuery()

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 分類の登録処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに分類の登録処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertBunrui(ByRef dataEXTM0104 As DataEXTM0104, _
                                 ByVal Cn As NpgsqlConnection) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.InsertBunrui(cmd, Cn, dataEXTM0104) = False Then
                Return False

            End If
            'sqlを実行する
            cmd.ExecuteNonQuery()

            '分類データを再取得
            If GetBunruiData(dataEXTM0104) = False Then
                Return False
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金の登録処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに料金の登録処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertRyokin(ByRef dataEXTM0104 As DataEXTM0104, _
                                 ByVal Cn As NpgsqlConnection) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.InsertRyokin(cmd, Cn, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            cmd.ExecuteNonQuery()

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金の更新処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに料金の更新処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function UpdataRyokin(ByRef dataEXTM0104 As DataEXTM0104, _
                                 ByVal Cn As NpgsqlConnection) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.UpdataRyokin(cmd, Cn, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            cmd.ExecuteNonQuery()

            '開始ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 倍率の登録処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに倍率の登録処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertBairitu(ByRef dataEXTM0104 As DataEXTM0104, _
                                  ByVal Cn As NpgsqlConnection) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.InsertBairitu(cmd, Cn, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            cmd.ExecuteNonQuery()

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 倍率の更新処理
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>入力された値をもとに倍率の更新処理を行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    ''' 
    Public Function UpdataBairitu(ByRef dataEXTM0104 As DataEXTM0104, _
                                  ByVal Cn As NpgsqlConnection) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmd As New NpgsqlCommand        'アダプタ

        Try
            'sqlを設定する
            If SqlEXTM0104.UpdataBairitu(cmd, Cn, dataEXTM0104) = False Then
                Return False
            End If

            'sqlを実行する
            cmd.ExecuteNonQuery()

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            cmd.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 期間入力チェック
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>利用料マスタメンテ画面で入力された期間(FROM)年、期間(FROM)月、期間(TO)年、期間(TO)月の入力チェックを行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function KikanInputCheck(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '半角数字のパターンチェック用変数
        Dim number As New System.Text.RegularExpressions.Regex("^[0-9]+$")
        '設定済みの期間を判定するための変数を宣言
        Dim kf As String
        Dim kt As String
        Dim txtkt As String

        Try
            With dataEXTM0104
                '期間(FROM)年の入力チェック
                '期間(FROM)年の必須チェック
                If .PropTxtKikanFromYear.Text = "" Then
                    puErrMsg = String.Format(E0001, "期間(FROM)年")
                    Return False
                End If
                '期間(FROM)年の文字パターン
                If number.IsMatch(.PropTxtKikanFromYear.Text) = False Then
                    puErrMsg = String.Format(E0003, "期間(FROM)年")
                    Return False
                End If

                '期間（FROM)年の桁数チェック
                If .PropTxtKikanFromYear.Text.Length <> 4 Then
                    puErrMsg = String.Format(E0010, "期間(FROM)年", "4")
                    Return False
                End If

                '期間(FROM)月の入力チェック
                '期間(FROM)月の必須チェック
                If .PropTxtKikanFromMonth.Text = "" Then
                    puErrMsg = String.Format(E0001, "期間(FROM)月")
                    Return False
                End If
                '期間(FROM)月の文字パターン
                If number.IsMatch(.PropTxtKikanFromMonth.Text) = False Then
                    puErrMsg = String.Format(E0003, "期間(FROM)月")
                    Return False
                End If

                '期間（FROM)月の桁数チェック
                If .PropTxtKikanFromMonth.Text.Length > 2 Then
                    puErrMsg = String.Format(E0009, "期間(FROM)月", "1", "2")
                    Return False
                End If

                '期間(FROM)月の範囲チェック
                If CInt(.PropTxtKikanFromMonth.Text) < 1 Or CInt(.PropTxtKikanFromMonth.Text) > 12 Then
                    puErrMsg = String.Format(E0018, "期間(FROM)月", "1", "12")
                    Return False
                End If

                '期間(TO)年の入力チェック
                '期間(TO)年の必須チェック
                If .PropTxtKikanToYear.Text = "" Then
                    puErrMsg = String.Format(E0001, "期間(TO)年")
                    Return False
                End If

                '期間(TO)年の文字パターン
                If number.IsMatch(.PropTxtKikanToYear.Text) = False Then
                    puErrMsg = String.Format(E0003, "期間(TO)年")
                    Return False
                End If

                '期間(TO)年の桁数チェック
                If .PropTxtKikanToYear.Text.Length <> 4 Then
                    puErrMsg = String.Format(E0010, "期間(TO)年", "4")
                    Return False
                End If

                '期間(TO)月の入力チェック
                '期間(TO)月の必須チェック
                If .PropTxtkikanToMonth.Text = "" Then
                    puErrMsg = String.Format(E0001, "期間(TO)月")
                    Return False
                End If

                '期間(TO)月の文字パターン
                If number.IsMatch(.PropTxtkikanToMonth.Text) = False Then
                    puErrMsg = String.Format(E0003, "期間(TO)月")
                    Return False
                End If

                '期間(TO)月の桁数チェック
                If .PropTxtkikanToMonth.Text.Length > 2 Then
                    puErrMsg = String.Format(E0009, "期間(TO)月", "1", "2")
                    Return False
                End If

                If CInt(.PropTxtkikanToMonth.Text) < 1 Or CInt(.PropTxtkikanToMonth.Text) > 12 Then
                    puErrMsg = String.Format(E0018, "期間(TO)月", "1", "12")
                    Return False
                End If

                If CInt(.PropTxtKikanFromYear.Text + .PropTxtKikanFromMonth.Text) > CInt(.PropTxtKikanToYear.Text + .PropTxtkikanToMonth.Text) Then
                    puErrMsg = String.Format(E0020, "期間(FROM)", "期間(TO)")
                    Return False
                End If

                For index = 0 To dataEXTM0104.PropDtKikan.Rows.Count - 1
                    If dataEXTM0104.PropDtKikan.Rows(index).Item(1) = Nothing Then
                    Else
                        kf = dataEXTM0104.PropDtKikan.Rows(index)(0).ToString.Substring(0, 10)
                        kf = DateTime.Parse(kf).ToString("yyyyMMdd")
                        kt = dataEXTM0104.PropDtKikan.Rows(index)(0).ToString.Substring(10, 10)
                        kt = DateTime.Parse(kt).ToString("yyyyMMdd")

                        If dataEXTM0104.PropRdoshinki.Checked = True Then
                            '期間FROMが既存している期間に含まれているかの判定
                            If Integer.Parse(dataEXTM0104.PropTxtKikanFromYear.Text + dataEXTM0104.PropTxtKikanFromMonth.Text + "01") >= Integer.Parse(kf) _
                                AndAlso Integer.Parse(dataEXTM0104.PropTxtKikanFromYear.Text + dataEXTM0104.PropTxtKikanFromMonth.Text + "01") <= Integer.Parse(kt) Then
                                puErrMsg = E2032
                                Return False
                            End If
                            '期間TOが既存している期間に含まれているかの判定
                            txtkt = DateTime.Parse(dataEXTM0104.PropTxtKikanToYear.Text + "/" + dataEXTM0104.PropTxtkikanToMonth.Text).AddMonths(1).AddDays(-1).ToString("yyyyMMdd")

                            If Integer.Parse(txtkt) >= Integer.Parse(kf) _
                                AndAlso Integer.Parse(txtkt) <= Integer.Parse(kt) Then
                                puErrMsg = E2032
                                Return False
                            End If
                            '画面期間FROM　< DB期間　<画面期間TOの場合
                            If Integer.Parse(dataEXTM0104.PropTxtKikanFromYear.Text + dataEXTM0104.PropTxtKikanFromMonth.Text + "01") < Integer.Parse(kf) _
                                AndAlso Integer.Parse(txtkt) > Integer.Parse(kt) Then
                                puErrMsg = E2032
                                Return False
                            End If
                        End If
                    End If
                Next

            End With
            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try
    End Function

    ''' <summary>
    ''' 分類入力チェック
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>利用料マスタメンテ画面で入力された分類ビューの値の入力チェックを行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    ''' 
    Public Function BunruiInputCheck(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ビューカウント変数を宣言
        Dim vwct As Integer = dataEXTM0104.PropVwBunrui.Sheets(0).RowCount
        'データテーブルカウント変数を宣言
        Dim datact As Integer = dataEXTM0104.PropDtBunrui.Rows.Count
        '半角数字のパターンチェック用変数
        Dim number As New System.Text.RegularExpressions.Regex("^[0-9]+$")
        '1行に空項目の個数のカウンタ
        Dim cnt As Integer = 0

        Try
            '画面で入力された期間を条件に分類を取得する【既存するデータであるかどうかの判定を行うため】
            If GetBunruiData(dataEXTM0104) = False Then
                Return False
            End If

            'から項目の個数を数える
            For index = 0 To vwct - 1
                cnt = 0
                For j = 0 To 4
                    If dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, j).Text = Nothing Then
                        cnt += 1
                    End If
                Next

                '全部からではない時に以下の処理
                If cnt <> 5 Then

                    With dataEXTM0104.PropVwBunrui.Sheets(0)
                        '料金分類CDの必須チェック
                        If .Cells(index, 0).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(分類)]料金分類コード")
                            Return False
                        End If

                        '料金分類コードの文字パターンチェック
                        If number.IsMatch(.Cells(index, 0).Text) = False Then
                            puErrMsg = String.Format(E0003, "[利用料(分類)]料金分類コード")
                            Return False
                        End If

                        '料金分類CDの桁数チェック
                        If .Cells(index, 0).Text.Length > 2 Then
                            puErrMsg = String.Format(E0009, "[利用料(分類)]料金分類コード", "1", "2")
                            Return False
                        End If

                        '施設の入力必須チェック
                        If .Cells(index, 1).Text = "" Then
                            puErrMsg = String.Format(E0002, "[利用料(分類)]施設")
                            Return False
                        End If

                        '料金分類名の必須チェック
                        If .Cells(index, 2).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(分類)]料金分類名")
                            Return False
                        End If

                        '料金分類名の桁数チェック
                        If .Cells(index, 2).Text.Length > 20 Then
                            puErrMsg = String.Format(E0009, "[利用料(分類)]料金分類名", "1", "20")
                            Return False
                        End If

                        '並び順の必須チェック
                        If .Cells(index, 3).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(分類)]並び順")
                            Return False
                        End If

                        '並び順の文字パターンチェック
                        If number.IsMatch(.Cells(index, 3).Text) = False Then
                            puErrMsg = String.Format(E0003, "[利用料(分類)]並び順")
                            Return False
                        End If

                        '並び順の桁数チェック
                        If .Cells(index, 3).Text.Length > 4 Then
                            puErrMsg = String.Format(E0009, "[利用料(分類)]並び順", "1", "4")
                            Return False
                        End If
                    End With
                End If
            Next
            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 料金入力チェック
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>利用料マスタメンテ画面で入力された料金ビューの値の入力チェックを行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function RyokinInputCheck(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '半角数字のパターンチェック用変数
        Dim number As New System.Text.RegularExpressions.Regex("^[0-9]+$")
        'ビューカウント変数を宣言
        Dim vwct As Integer = dataEXTM0104.PropVwRyokin.Sheets(0).RowCount
        'データテーブルカウント変数を宣言
        Dim datact As Integer = dataEXTM0104.PropDtRyoukin.Rows.Count

        '1行にから項目の個数のカウント変数
        Dim cnt As Integer = 0

        Try

            '画面で入力された期間を条件に料金を取得する【既存するデータであるかどうかの判定を行うため】
            '20151023 If GetRyokinData(dataEXTM0104, dataEXTM0104.PropVwBunrui.Sheets(0).ActiveRowIndex) = False Then
            '20151023     Return False
            '20151023 End If

            For index = 0 To vwct - 1

                cnt = 0
                For j = 0 To 6
                    If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, j).Text = Nothing Then
                        cnt += 1
                    End If
                Next

                '全部空ではない時に以下の処理
                If cnt <> 7 Then
                    With dataEXTM0104.PropVwRyokin.Sheets(0)

                        '料金分類CDの必須チェック
                        If .Cells(index, 0).Text = "" Then
                            puErrMsg = String.Format(E0001, "料金分類コード")
                            Return False
                        End If

                        '料金分類CDの文字パターンチェック
                        If number.IsMatch(.Cells(index, 0).Text) = False Then
                            puErrMsg = String.Format(E0003, "料金分類コード")
                            Return False
                        End If

                        '料金分類CDの桁数チェック
                        If .Cells(index, 0).Text.Length > 2 Then
                            puErrMsg = String.Format(E0009, "料金分類コード", "1", "2")
                            Return False
                        End If

                        '料金CDの必須チェック
                        If .Cells(index, 1).Text = "" Then
                            puErrMsg = String.Format(E0001, "料金コード")
                            Return False
                        End If

                        '料金CDの文字パターンチェック
                        If number.IsMatch(.Cells(index, 1).Text) = False Then
                            puErrMsg = String.Format(E0003, "[利用料(料金)]料金コード")
                            Return False
                        End If

                        '料金CDの桁数チェック
                        If .Cells(index, 1).Text.Length > 3 Then
                            puErrMsg = String.Format(E0009, "[利用料(料金)]料金コード", "1", "3")
                            Return False
                        End If

                        '料金名の必須チェック
                        If .Cells(index, 2).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(料金)]料金名")
                            Return False
                        End If

                        '料金名の桁数チェック
                        If .Cells(index, 2).Text.Length > 20 Then
                            puErrMsg = String.Format(E0009, "[利用料(料金)]料金名", "1", "20")
                            Return False
                        End If

                        '時間貸しの必須チェック
                        If .Cells(index, 3).Text = "" Then
                            'puErrMsg = String.Format(E0001, "[利用料(料金)]時間貸し(1h)")                   ' 2015.12.21 UPD h.hagiwara
                            puErrMsg = String.Format(E0001, "[利用料(料金)]時間貸し")                        ' 2015.12.21 UPD h.hagiwara
                            Return False
                        End If

                        '時間貸しの文字パターンチェック
                        ' 2015.12.01 UPD START↓ h.hagiwara
                        'If number.IsMatch(.Cells(index, 3).Text) = False Then
                        '    puErrMsg = String.Format(E0003, "[利用料(料金)]時間貸し(1h)")
                        '    Return False
                        'End If
                        If IsNumeric(.Cells(index, 3).Value) = False Then
                            'puErrMsg = String.Format(E0003, "[利用料(料金)]時間貸し(1h)")                   ' 2015.12.21 UPD h.hagiwara
                            puErrMsg = String.Format(E0003, "[利用料(料金)]時間貸し")                        ' 2015.12.21 UPD h.hagiwara
                            Return False
                        End If
                        ' 2015.12.01 UPD END↑ h.hagiwara

                        '時間貸しの桁数チェック
                        'If .Cells(index, 3).Text.Length > 2 Then
                        'If .Cells(index, 3).Text.Length > 8 Then
                        If .Cells(index, 3).Text.Replace(",", "").Length > 8 Then
                            'puErrMsg = String.Format(E0009, "[利用料(料金)]時間貸し(1h)", "1", "8")        ' 2015.12.21 UPD h.hagiwara
                            puErrMsg = String.Format(E0009, "[利用料(料金)]時間貸し", "1", "8")             ' 2015.12.21 UPD h.hagiwara
                            Return False
                        End If

                        '利用料金の必須チェック
                        If .Cells(index, 4).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(料金)]利用料金(1日,Lockout)")
                            Return False
                        End If

                        '利用料金の文字パターンチェック
                        ' 2015.12.01 UPD START↓ h.hagiwara
                        'If number.IsMatch(.Cells(index, 4).Text) = False Then
                        '    puErrMsg = String.Format(E0003, "[利用料(料金)]利用料金(1日,Lockout)")
                        '    Return False
                        'End If
                        If IsNumeric(.Cells(index, 4).Value) = False Then
                            puErrMsg = String.Format(E0003, "[利用料(料金)]利用料金(1日,Lockout)")
                            Return False
                        End If
                        ' 2015.12.01 UPD END↑ h.hagiwara

                        '利用料金の桁数チェック
                        'If .Cells(index, 4).Text.Length > 8 Then
                        If .Cells(index, 4).Text.Replace(",", "").Length > 8 Then
                            puErrMsg = String.Format(E0009, "[利用料(料金)]利用料金(1日,Lockout)", "1", "8")
                            Return False
                        End If

                        '並び順の必須チェック
                        If .Cells(index, 5).Text = "" Then
                            puErrMsg = String.Format(E0001, "[利用料(料金)]並び順")
                            Return False
                        End If

                        '並び順の文字パターンチェック
                        If number.IsMatch(.Cells(index, 5).Text) = False Then
                            puErrMsg = String.Format(E0003, "[利用料(料金)]並び順")
                            Return False
                        End If
                        '並び順の桁数チェック
                        If .Cells(index, 5).Text.Length > 4 Then
                            puErrMsg = String.Format(E0009, "[利用料(料金)]並び順", "1", "4")
                            Return False
                        End If
                        .Cells(index, 7).Value = dataEXTM0104.PropStrShisetu
                        If .Cells(index, 8).Value = "" Then
                            .Cells(index, 8).Value = "1"
                        End If
                    End With
                End If
            Next

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 料金入力チェック
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>利用料マスタメンテ画面で入力された料金ビューの値の入力チェックを行う
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function BairituInputCheck(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '半角数字のパターンチェック用変数
        Dim number As New System.Text.RegularExpressions.Regex("^[0-9]+$")
        '小数点以下を含む半角数字のパターンチェック用変数
        Dim dbl As New System.Text.RegularExpressions.Regex("^[0-9]+[.]+[0-9]+$")
        'ビューカウント変数を宣言
        Dim vwct As Integer = dataEXTM0104.PropVwBairitu.Sheets(0).RowCount
        'データテーブルカウント変数を宣言
        Dim datact As Integer = dataEXTM0104.PropDtBairitu.Rows.Count

        'DOUBLE型の長さをチェックするための変数
        Dim intDot As Integer = 0
        Dim intIntegral As Integer = 0
        Dim intDecimal As Integer = 0
        Dim leg As Integer = 0

        '1行にから項目の個数のカウント変数
        Dim cnt As Integer = 0

        Try

            '画面で入力された期間を条件に料金を取得する【既存するデータであるかどうかの判定を行うため】
            If GetBairituData(dataEXTM0104) = False Then
                Return False
            End If

            For index = 0 To vwct - 1

                cnt = 0
                For j = 0 To 5
                    If dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, j).Text = Nothing Then
                        cnt += 1
                    End If
                Next

                '全部空ではない時に以下の処理
                If cnt <> 6 Then

                    With dataEXTM0104.PropVwBairitu.Sheets(0)
                        '倍率CDの必須チェック
                        If .Cells(index, 0).Text = "" Then
                            puErrMsg = String.Format(E0001, "[倍率]倍率コード")
                            Return False
                        End If

                        '倍率CDの文字パターンチェック
                        If number.IsMatch(.Cells(index, 0).Text) = False Then
                            puErrMsg = String.Format(E0003, "[倍率]倍率コード")
                            Return False
                        End If

                        '倍率CDの桁数チェック（２）
                        If .Cells(index, 0).Text.Length > 2 Then
                            puErrMsg = String.Format(E0009, "[倍率]倍率コード", "1", "2")
                            Return False
                        End If

                        '施設の必須チェック
                        If .Cells(index, 1).Text = "" Then
                            puErrMsg = String.Format(E0001, "[倍率]施設")
                            Return False
                        End If

                        '倍率名の必須チェック
                        If .Cells(index, 2).Text = "" Then
                            puErrMsg = String.Format(E0001, "[倍率]倍率名")
                            Return False
                        End If

                        '倍率名の桁数チェック
                        If .Cells(index, 2).Text.Length > 20 Then
                            puErrMsg = String.Format(E0009, "[倍率]倍率名", "1", "20")
                            Return False
                        End If

                        '倍率の必須チェック
                        If .Cells(index, 3).Text = "" Then
                            puErrMsg = String.Format(E0001, "[倍率]倍率")
                            Return False
                        End If

                        '倍率の文字パターンチェック
                        'どうやる？　コンマがあるため数字のみではない
                        'Dim stArrayData As String() = .Cells(index, 4).Text.Split(".")
                        'MsgBox(stArrayData(0))
                        'MsgBox(stArrayData(1))
                        If number.IsMatch(.Cells(index, 3).Text) = False Then
                            If dbl.IsMatch(.Cells(index, 3).Text) = False Then
                                puErrMsg = String.Format(E0003, "[倍率]倍率")
                                Return False
                            End If
                        End If

                        'double型の長さチェック用の変数を初期化

                        intDot = 0
                        intIntegral = 0
                        intDecimal = 0
                        leg = 0
                        intDot = InStr(.Cells(index, 3).Text.ToString(), ".")
                        If intDot = 0 Then
                            intIntegral = .Cells(index, 3).Text.ToString().Length
                        Else
                            intIntegral = intDot - 1
                            intDecimal = .Cells(index, 3).Text.ToString.Length - intDot
                        End If
                        leg = intIntegral + intDecimal

                        If leg > 3 Then
                            puErrMsg = String.Format(E0009, "[倍率]倍率", "1", "3")
                            Return False
                        End If

                        '並び順の必須チェック
                        If .Cells(index, 4).Text = "" Then
                            puErrMsg = String.Format(E0001, "[倍率]並び順")
                            Return False
                        End If

                        '並び順の文字パターンチェック
                        If number.IsMatch(.Cells(index, 4).Text) = False Then
                            puErrMsg = String.Format(E0003, "[倍率]並び順")
                            Return False
                        End If

                        '並び順の桁数チェック
                        If .Cells(index, 4).Text.Length > 4 Then
                            puErrMsg = String.Format(E0009, "[倍率]並び順", "1", "4")
                            Return False
                        End If

                    End With

                End If
            Next

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try
    End Function

    ''' <summary>
    ''' 分類ビューをクリアする
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>分類ビューのすべての値をクリアする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClearVwBunrui(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            For index = 0 To dataEXTM0104.PropVwBunrui.Sheets(0).RowCount - 1
                For i = 0 To dataEXTM0104.PropVwBunrui.Sheets(0).ColumnCount - 1
                    dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, i).Text = ""
                Next
            Next

            For index = 0 To dataEXTM0104.PropVwBunrui.Sheets(0).RowCount - 1
                dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 0).Locked = False
                dataEXTM0104.PropVwBunrui.Sheets(0).Cells(index, 1).Locked = False
            Next

            dataEXTM0104.PropVwBunrui.Sheets(0).RowCount = 5

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金ビューをクリアする
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>料金ビューのすべての値をクリアする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClearVwRyokin(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            For index = 0 To dataEXTM0104.PropVwRyokin.Sheets(0).RowCount - 1
                For i = 0 To dataEXTM0104.PropVwRyokin.Sheets(0).ColumnCount - 1
                    dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, i).Text = ""
                Next

            Next
            For index = 0 To dataEXTM0104.PropVwRyokin.Sheets(0).RowCount - 1
                dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 0).Locked = False
                dataEXTM0104.PropVwRyokin.Sheets(0).Cells(index, 1).Locked = False

            Next

            dataEXTM0104.PropVwRyokin.Sheets(0).RowCount = 5

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 倍率ビューをクリアする
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>倍率ビューの値をすべてクリアする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>

    Public Function ClearVwBairitu(ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            For index = 0 To dataEXTM0104.PropVwBairitu.Sheets(0).RowCount - 1
                For i = 0 To dataEXTM0104.PropVwBairitu.Sheets(0).ColumnCount - 1
                    dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, i).Text = ""
                Next
            Next

            For index = 0 To dataEXTM0104.PropVwBairitu.Sheets(0).RowCount - 1
                dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 0).Locked = False
                dataEXTM0104.PropVwBairitu.Sheets(0).Cells(index, 1).Locked = False
            Next
            dataEXTM0104.PropVwBairitu.Sheets(0).RowCount = 5

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 期間FROM月のtextchange
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>期間FROM月の1桁の数字が入力されたら前に0を追加する
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function TxtChangeKikanFromMonth(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            If dataEXTM0104.PropTxtKikanFromMonth.Text.Length = 1 Then
                dataEXTM0104.PropTxtKikanFromMonth.Text = "0" + dataEXTM0104.PropTxtKikanFromMonth.Text
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 期間TO月のtextchange
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>期間TO月の1桁の数字が入力されたら前に0を追加する
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function TxtChangeKikanToMonth(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            If dataEXTM0104.PropTxtkikanToMonth.Text.Length = 1 Then
                dataEXTM0104.PropTxtkikanToMonth.Text = "0" + dataEXTM0104.PropTxtkikanToMonth.Text
            End If

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try
    End Function

    ''' <summary>
    ''' 期間セット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>期間(FROM)年、期間(FROM)月、期間(TO)年、期間(TO)月をセットする
    ''' <para>作成情報：2015/08/19 mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetTxtKikan(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            dataEXTM0104.PropTxtKikanFromYear.Text = dataEXTM0104.PropCmbKikan.SelectedValue.ToString.Substring(0, 4)
            dataEXTM0104.PropTxtKikanFromMonth.Text = dataEXTM0104.PropCmbKikan.SelectedValue.ToString.Substring(5, 2)
            dataEXTM0104.PropTxtKikanToYear.Text = dataEXTM0104.PropCmbKikan.SelectedValue.ToString.Substring(10, 4)
            dataEXTM0104.PropTxtkikanToMonth.Text = dataEXTM0104.PropCmbKikan.SelectedValue.ToString.Substring(15, 2)

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' サーバ日時取得
    ''' <paramref name="dataEXTM0104">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>サーバ日時取得
    ''' <para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetSelectSysDate(ByRef dataEXTM0104 As DataEXTM0104)

        '開始ログ出力()
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる

        Try
            'コネクションを開く
            Cn.Open()

            '新規Inc番号、システム日付取得（SELECT）
            If SelectSysDate(Cn, dataEXTM0104) = False Then
                Return False
            End If

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = E0000 & ex.Message
            Return False

        Finally
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' サーバ日時取得
    ''' </summary>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTM0104">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>サーバ日時を取得する
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SelectSysDate(ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTM0104 As DataEXTM0104) As Boolean

        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlDataAdapter
        Dim dtResult As New DataTable

        Try
            '取得用SQLの作成・設定
            If SqlEXTM0104.GetSysDate(Adapter, Cn, DataEXTM0104) = False Then
                Return False
            End If

            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.DEBUG_Lv, "サーバ日時", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtResult)

            'データが取得できなかった場合、エラー
            If dtResult.Rows.Count = 0 Then
                Return False
            End If

            '取得データをデータクラスにセット
            dataEXTM0104.PropDtSysDate = dtResult.Rows(0).Item(0)

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金データ再セット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>SPREAD上に表示している情報と退避した情報をマージする
    ''' <para>作成情報：2015.10.23 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function RyokinInfSet(ByRef dataEXTM0104 As DataEXTM0104) As Boolean
        '開始ログ出力
        commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim dtrow1 As DataRow

        Try

            dataEXTM0104.PropDtRyoukin.Clear()

            For i = 0 To dataEXTM0104.PropDtRyokinEscp.Rows.Count - 1
                dtrow1 = dataEXTM0104.PropDtRyoukin.NewRow
                For k = 0 To 8
                    dtrow1(k) = dataEXTM0104.PropDtRyokinEscp.Rows(i).Item(k)
                Next
                dataEXTM0104.PropDtRyoukin.Rows.Add(dtrow1)
            Next

            ' 格納領域を表示用・退避用をマージする。
            For i = 0 To dataEXTM0104.PropVwRyokin.Sheets(0).RowCount - 1
                If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, 0).Text = "" Then
                Else
                    dtrow1 = dataEXTM0104.PropDtRyoukin.NewRow
                    dtrow1(0) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_BUNRUICD).Value
                    dtrow1(1) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_RYOKINCD).Value
                    dtrow1(2) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_RYOKINNM).Value
                    dtrow1(3) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_RYOKINHOUR).Value
                    dtrow1(4) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_RYOKIN).Value
                    dtrow1(5) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_SORT).Value
                    If dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_STS).Value Then
                        dtrow1(6) = True
                    Else
                        dtrow1(6) = False
                    End If
                    dtrow1(7) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_SHISETU).Value
                    dtrow1(8) = dataEXTM0104.PropVwRyokin.Sheets(0).Cells(i, M0104_RYOKIN_COL_UPKBN).Value
                    dataEXTM0104.PropDtRyoukin.Rows.Add(dtrow1)
                End If
            Next

            '終了ログ出力
            commonlogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonlogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

End Class
