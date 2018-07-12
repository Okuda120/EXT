Imports Npgsql
Imports Common
Imports CommonEXT
Imports FarPoint.Win.Spread
Imports System.Text.RegularExpressions

Public Class LogicEXTM0103

    'インスタンス作成
    Public SqlEXTM0103 As New SqlEXTM0103
    Public CommonEXT As New CommonLogicEXT
    Public CommonValidation As New CommonValidation

    ''' <summary>
    ''' 一覧データの取得
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>初期表示時検索を行う
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitSearch(ByRef dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)                    'コネクション
        Dim Adapter As New NpgsqlDataAdapter                        'アダプタ
        Dim Table As New DataTable()                                'テーブル

        Try
            'コネクションを開く
            Cn.Open()

            '検索用SQLの作成・設定
            If SqlEXTM0103.SetSelectInitTaxSearchSql(Adapter, Cn, dataEXTM0103) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXシアター消費税マスタメンテ初期表示時検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            'データテーブルに取得データをセット
            dataEXTM0103.PropDtExtTaxMasta = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function
    ''' <summary>
    ''' 画面表示
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>シートに情報をセットする
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSheet(ByRef dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'スプレッドシート上の行を削除
            dataEXTM0103.PropVwList.Sheets(0).RemoveRows(0, dataEXTM0103.PropVwList.Sheets(0).Rows.Count)

            If dataEXTM0103.PropDtExtTaxMasta Is Nothing Then

                '10件の空白登録枠を表示
                dataEXTM0103.PropVwList.Sheets(0).RowCount = 10

            Else

                With dataEXTM0103

                    '常に10件の登録枠を表示
                    .PropVwList.Sheets(0).RowCount = 10

                    For i As Integer = 0 To dataEXTM0103.PropDtExtTaxMasta.Rows.Count - 1
                        '一覧シートに一行ずつデータをセット
                        .PropVwList.Sheets(0).Cells(i, 0).Text = .PropDtExtTaxMasta.Rows(i).Item("TAXS_DT")
                        If .PropDtExtTaxMasta.Rows(i).Item("TAXE_DT").ToString = "2099/12/31" Then
                            .PropVwList.Sheets(0).Cells(i, 1).Text = ""
                        Else
                            .PropVwList.Sheets(0).Cells(i, 1).Text = .PropDtExtTaxMasta.Rows(i).Item("TAXE_DT")
                        End If
                        .PropVwList.Sheets(0).Cells(i, 2).Text = .PropDtExtTaxMasta.Rows(i).Item("TAX_RITU")
                        .PropVwList.Sheets(0).Cells(i, 3).Text = .PropDtExtTaxMasta.Rows(i).Item("SEQ")
                        .PropVwList.Sheets(0).Cells(i, 4).Text = .PropDtExtTaxMasta.Rows(i).Item(4) '更新区分
                        .PropVwList.Sheets(0).Cells(i, 5).Text = .PropDtExtTaxMasta.Rows(i).Item(5) '修正前開始日
                        .PropVwList.Sheets(0).Cells(i, 6).Text = .PropDtExtTaxMasta.Rows(i).Item(6) '修正前終了日
                        .PropVwList.Sheets(0).Cells(i, 7).Text = .PropDtExtTaxMasta.Rows(i).Item(7) '修正前消費税率
                    Next

                End With


            End If
            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 入力チェック
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の入力チェックを行う
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckInputSheet(ByRef dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '行数カウント・初期化
        Dim intRowCount As Integer = 0

        '空行判断
        For intRow As Integer = 0 To 9
            Dim noData As Integer = 0    '一行あたりの空欄数・初期化
            For intColumn As Integer = 0 To 2
                If dataEXTM0103.PropVwList.Sheets(0).Cells(intRow, intColumn).Value = Nothing Then
                    noData += 1          '一行あたりの空欄数をカウントアップ
                End If
            Next
            If noData = 3 Then           '一行あたりの空欄数が3の場合＝空行
                Exit For
            End If
            intRowCount += 1
        Next

        Try
            '入力行数分 入力チェックを行う
            For i As Integer = 0 To intRowCount - 1

                With dataEXTM0103.PropVwList.Sheets(0)
                    '開始日
                    With .Cells(i, 0)
                        '未入力の場合、エラー
                        If .Text.Trim() = "" Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E0001, "開始日")
                            'エラーを返す
                            Return False
                        End If
                        '入力日付がカレンダー上に存在するかどうか
                        If CommonEXT.IsDate(.Text) = False Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E2005, "開始日")
                            'エラーを返す
                            Return False
                        End If
                        '半角数字（記号有）
                        If Regex.IsMatch(.Text, "^[0-9/]*$") = False Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E0004, "開始日")
                            'エラーを返す
                            Return False
                        End If
                        '（2行目以降）ある行の開始日 > 前行の開始日
                        If i > 0 Then
                            If .Value < dataEXTM0103.PropVwList.Sheets(0).Cells(i - 1, 0).Value Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0103_E2005, "開始日")
                                'エラーを返す
                                Return False
                            End If
                        End If

                    End With

                    '終了日
                    With .Cells(i, 1)
                        '最終行は空白のみ可能
                        If i = intRowCount - 1 Then
                            If .Text.Trim() <> "" Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0103_E2040, "終了日")
                                Return False
                            End If

                        Else
                            '入力日付がカレンダー上に存在するかどうか
                            If CommonEXT.IsDate(.Text) = False Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0103_E2005, "終了日")
                                'エラーを返す
                                Return False
                            End If
                            'ある行の開始日 < 同じ行の終了日
                            If .Value <> Nothing Then
                                If .Value < dataEXTM0103.PropVwList.Sheets(0).Cells(i, 0).Value Then
                                    'エラーメッセージ設定
                                    puErrMsg = String.Format(M0103_E2011, "終了日", "開始日")
                                    'エラーを返す
                                    Return False
                                End If
                            End If
                            '半角数字（記号有）
                            If Regex.IsMatch(.Text, "^[0-9/]*$") = False Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0103_E0004, "終了日")
                                'エラーを返す
                                Return False
                            End If

                        End If
                    End With

                    '消費税率
                    With .Cells(i, 2)
                        '未入力の場合、エラー
                        If .Text.Trim() = "" Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E0001, "消費税率")
                            'エラーを返す
                            Return False
                        End If
                        '桁数チェック(3桁以内)
                        If .Text.Length > 3 Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E0009, "消費税率", "1", "3")
                            'エラーを返す
                            Return False
                        End If
                        '半角数字
                        If CommonValidation.IsHalfNmb(.Text) = False Then
                            'エラーメッセージ設定
                            puErrMsg = String.Format(M0103_E0003, "消費税率")
                            'エラーを返す
                            Return False
                        End If
                        '(2行目以降)消費税率が前行と重複している場合
                        If i > 0 Then
                            If .Value = dataEXTM0103.PropVwList.Sheets(0).Cells(i - 1, 2).Value Then
                                'エラーメッセージ設定
                                puErrMsg = String.Format(M0103_E2039, "消費税率")
                                'エラーを返す
                                Return False
                            End If
                        End If
                    End With

                End With
            Next

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False

        End Try
        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        '確認メッセージ設定
        puErrMsg = String.Format(M0103_C0011, "消費税マスタ")
        Return True

    End Function
    ''' <summary>
    ''' 追加・更新チェック
    ''' </summary>
    ''' <param name="dataEXTM0103">[IN/OUT]消費税マスタメンテDataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>隠し列（＝更新区分）によって条件分岐処理を行う
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ChkTaxmstData(ByRef dataEXTM0103 As DataEXTM0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数を宣言
        Dim Cn As New NpgsqlConnection(DbString)  'コネクション
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション
        Dim Table As DataTable                    'データテーブル

        Table = DirectCast(dataEXTM0103.PropVwList.DataSource, DataTable)

        '行数カウント・初期化
        Dim intRowCount As Integer = 0

        '終了日の自動修正
        Dim EndDt As DateTime            '前行の終了日
        Dim StartDt As DateTime          '修正または追加された開始日
        Dim StrDelflg As String = ""                                                                   ' 2015.12.08 ADD h.hagiwara

        '空行判断
        'For intRow As Integer = 0 To 9                                                                ' 2015.12.08 UPD h.hagiwara
        For intRow As Integer = 0 To dataEXTM0103.PropVwList.Sheets(0).RowCount - 1                    ' 2015.12.08 UPD h.hagiwara
            Dim noData As Integer = 0    '一行あたりの空欄数・初期化
            For intColumn As Integer = 0 To 2
                If dataEXTM0103.PropVwList.Sheets(0).Cells(intRow, intColumn).Value = Nothing Then
                    noData += 1          '一行あたりの空欄数をカウントアップ
                End If
            Next
            If noData = 3 Then           '一行あたりの空欄数が3の場合＝空行
                ' 2015.12.08 UPD START↓ h.hagiwara 既存情報で空白とした場合は処理対象とする
                'Exit For
                If dataEXTM0103.PropVwList.Sheets(0).Cells(intRow, 4).Value = 1 Then
                    dataEXTM0103.PropVwList.Sheets(0).Cells(intRow, 4).Value = 2
                Else
                    Exit For
                End If
                ' 2015.12.08 UPD END↑ h.hagiwara 既存情報で空白とした場合は処理対象とする
            End If
            intRowCount += 1
        Next

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '一覧シートの入力行数分チェック
            For j As Integer = 0 To intRowCount - 1
                '初期化
                EndDt = New DateTime                     '前行の終了日
                StartDt = New DateTime                   '開始日

                '隠し列（＝更新区分）に１が立っているかチェック
                If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 4).Value = 2 Then                                        ' 2015.12.08 ADD h.hagiwara データ削除追加
                    '既存データの削除処理(DELETE)                                                                      ' 2015.12.08 ADD h.hagiwara データ削除追加
                    DeleteDB(dataEXTM0103, j, Tsx, Cn)                                                                 ' 2015.12.08 ADD h.hagiwara データ削除追加
                    StrDelflg = "1"                                                                                    ' 2015.12.08 ADD h.hagiwara データ削除追加

                    'If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 4).Value <> 1 Then                                  ' 2015.12.08 UPD h.hagiwara データ削除追加
                ElseIf dataEXTM0103.PropVwList.Sheets(0).Cells(j, 4).Value <> 1 Then                                   ' 2015.12.08 UPD h.hagiwara データ削除追加
                    '新規登録 = 追加処理(Insert)
                    Insert(dataEXTM0103, j, Tsx, Cn)

                    ' 2015.10.09 UPDATE START↓ h.hagiwara １件目登録時のINDEXエラー対応
                    'If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value = Nothing Then
                    '    EndDt = Date.Parse("2099/12/31")
                    'Else
                    '    EndDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value)
                    'End If
                    If j = 0 Then
                    Else
                        If dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value = Nothing Then
                            EndDt = Date.Parse("2099/12/31")
                        Else
                            EndDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value)
                        End If
                        StartDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Value)
                        '前行の終了日が不適切なとき、自動修正
                        If EndDt.AddDays(1) <> StartDt Then
                            UpdateTaxeDt(dataEXTM0103, j, Tsx, Cn)
                        End If
                    End If
                    ' 2015.10.09 UPDATE END↑ h.hagiwara １件目登録時のINDEXエラー対応
                    StrDelflg = ""                                                                                     ' 2015.12.08 ADD h.hagiwara データ削除追加
                Else
                    '隠し列（修正前データ）と比較・既存データが修正されたかチェック
                    If dataEXTM0103.PropVwList.Sheets(0).Cells(j, 5).Text <> dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Text Or
                       dataEXTM0103.PropVwList.Sheets(0).Cells(j, 6).Text <> dataEXTM0103.PropVwList.Sheets(0).Cells(j, 1).Text Or
                       dataEXTM0103.PropVwList.Sheets(0).Cells(j, 7).Text <> dataEXTM0103.PropVwList.Sheets(0).Cells(j, 2).Text Then

                        '既存データの修正 = 更新処理(Update)
                        Update(dataEXTM0103, j, Tsx, Cn)

                        '前行の終了日が不適切なとき、自動修正
                        If j > 0 Then
                            EndDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value)
                            StartDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Value)
                            If EndDt.AddDays(1) <> StartDt Then
                                UpdateTaxeDt(dataEXTM0103, j, Tsx, Cn)
                            End If
                        End If
                        StrDelflg = ""                                                                                 ' 2015.12.08 ADD h.hagiwara データ削除追加
                    Else                                                                                               ' 2015.12.08 ADD h.hagiwara データ削除追加
                        If StrDelflg = "1" Then                                                                        ' 2015.12.08 ADD h.hagiwara データ削除追加
                            '前行の終了日が不適切なとき、自動修正                                                      ' 2015.12.08 ADD h.hagiwara データ削除追加
                            If j > 0 Then                                                                              ' 2015.12.08 ADD h.hagiwara データ削除追加
                                EndDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j - 1, 1).Value)            ' 2015.12.08 ADD h.hagiwara データ削除追加
                                StartDt = Date.Parse(dataEXTM0103.PropVwList.Sheets(0).Cells(j, 0).Value)              ' 2015.12.08 ADD h.hagiwara データ削除追加
                                If EndDt.AddDays(1) <> StartDt Then                                                    ' 2015.12.08 ADD h.hagiwara データ削除追加
                                    UpdateTaxeDt(dataEXTM0103, j, Tsx, Cn)                                             ' 2015.12.08 ADD h.hagiwara データ削除追加
                                End If                                                                                 ' 2015.12.08 ADD h.hagiwara データ削除追加
                            End If                                                                                     ' 2015.12.08 ADD h.hagiwara データ削除追加
                            StrDelflg = ""                                                                             ' 2015.12.08 ADD h.hagiwara データ削除追加
                        End If
                    End If
                End If
            Next

            'コミット
            If Tsx IsNot Nothing Then
                Tsx.Commit()
            End If
            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '確認メッセージ設定
            puErrMsg = String.Format(M0103_I0002, "登録")
            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False

        End Try

    End Function
    ''' <summary>
    ''' 更新
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の登録処理を行う
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function Update(ByRef dataEXTM0103 As DataEXTM0103, _
                           ByVal j As Integer, _
                           ByRef Tsx As NpgsqlTransaction, _
                           ByRef Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            'ユーザーマスタ新規登録（INSERT）用SQLを作成
            If SqlEXTM0103.SetUpdateTaxmstSql(Cn, Cmd, j, dataEXTM0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "消費税マスタメンテ更新", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function
    ''' <summary>
    ''' 終了日更新
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>ある行の開始日が修正・または更新された際、前行の終了日に開始日-1日を上書きセットする
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function UpdateTaxeDt(ByRef dataEXTM0103 As DataEXTM0103, _
                                 ByRef j As Integer, _
                                 ByRef Tsx As NpgsqlTransaction, _
                                 ByRef Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            'ユーザーマスタ新規登録（UPDATE）用SQLを作成
            If SqlEXTM0103.SetUpdateTaxeDtSql(Cn, Cmd, j, dataEXTM0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "消費税マスタメンテ終了日更新", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function
    ''' <summary>
    ''' 新規登録
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の登録処理を行う
    ''' <para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function Insert(ByRef dataEXTM0103 As DataEXTM0103, _
                           ByVal j As Integer, _
                           ByRef Tsx As NpgsqlTransaction, _
                           ByRef Cn As NpgsqlConnection) As Boolean


        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try

            'ユーザーマスタ新規登録（INSERT）用SQLを作成
            If SqlEXTM0103.SetInsertTaxmstSql(Cn, Cmd, j, dataEXTM0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "消費税マスタメンテ新規登録", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function

    ''' <summary>
    ''' 削除
    ''' <paramref name="dataEXTM0103">[IN/OUT]データクラス</paramref>
    ''' <param name="j">[IN]行番号</param>
    ''' <param name="Tsx">[IN]トランザクション</param>
    ''' <param name="Cn">[IN]コネクション</param>
    ''' </summary>
    ''' <returns>boolean 終了状況    true  正常  false  異常</returns>
    ''' <remarks>シート内入力項目の削除処理を行う
    ''' <para>作成情報： 2015.12.08 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function DeleteDB(ByRef dataEXTM0103 As DataEXTM0103, _
                           ByVal j As Integer, _
                           ByRef Tsx As NpgsqlTransaction, _
                           ByRef Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            'ユーザーマスタ削除(DELETE)用SQLを作成
            If SqlEXTM0103.SetDeleteTaxmstSql(Cn, Cmd, j, dataEXTM0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "消費税マスタメンテ削除", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = String.Format(M0103_E0000)
            Return False
        Finally
            Cmd.Dispose()
        End Try

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        Return True

    End Function
End Class
