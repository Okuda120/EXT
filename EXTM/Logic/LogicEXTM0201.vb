Imports Common
Imports CommonEXT
Imports Npgsql
Imports FarPoint.Win.Spread
'Imports System.DirectoryService

''' <summary>利用者一覧画面Logicクラス
''' </summary>
''' <remarks>利用者一覧画面のロジックを定義する
''' <para>作成情報：2015/08/10 ozawa
''' <p>改訂情報：</p>
''' </para></remarks>

Class LogicEXTM0201
    'インスタンス生成
    Private SqlEXTM0201 As New SqlEXTM0201    'SQLクラス
    Private CommonLogic As New CommonLogic    'CommonLogicクラス
    Public DataEXTM0201 As New DataEXTM0201   'Dataクラス

    '定数宣言
    Public Const COL_SELECT As Integer = 0      '選択

    ''' <summary>初期表示時、スプレッドシート列表示処理
    ''' </summary>
    ''' <param name="dataEXTM0201">DataEXTM0201型オブジェクト</param>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>スプレッドシートに表示する列、行を制御する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function InitView(ByRef dataEXTM0201 As DataEXTM0201) As Boolean

        Try

            ' スプレッドの描画を停止
            dataEXTM0201.PropVwList.SuspendLayout()

            ' 全ての行を非表示にする
            For Each row As FarPoint.Win.Spread.Row In dataEXTM0201.PropVwList.ActiveSheet.Rows
                row.Visible = False
            Next

            Return True

        Catch ex As Exception
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False
        Finally
            ' スプレッドの描画を再開
            dataEXTM0201.PropVwList.ResumeLayout(True)
        End Try

    End Function


    ''' <summary>初期検索を行う
    ''' </summary>
    ''' <param name="dataEXTM0201"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function InitSearch(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)       'コネクション
        Dim Adapter As New NpgsqlDataAdapter           'アダプター
        Dim Table As New DataTable()                   'テーブル

        Try
            'コネクションを開く
            Cn.Open()
            'エラーメッセージ初期化
            puErrMsg = System.String.Empty

            'SELECT用SQLコマンドを作成
            If SqlEXTM0201.SetSelectInit(Adapter, Cn, dataEXTM0201) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "初期検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得データをDataクラスへ保存
            dataEXTM0201.PropVwList.DataSource = Table

            'O件の場合
            If dataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then
                ClearSpreadRow(dataEXTM0201)
            End If

            'セルタイプ変更（確認・編集ボタン表示）
            Dim Btn_Hensyu As New CellType.ButtonCellType
            'ボタン名
            Btn_Hensyu.Text = "確認・編集"
            dataEXTM0201.PropVwList.Sheets(0).Columns(13).CellType = Btn_Hensyu

            '終了ログ
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>検索処理
    ''' </summary>
    ''' <remarks>検索ボタン押下時、データ取得SQLを作成しデータを取得
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    ''' 
    Public Function Search(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)       'コネクション
        Dim Adapter As New NpgsqlDataAdapter           'アダプター
        Dim Table As New DataTable()                   'テーブル

        Try
            'コネクションを開く
            Cn.Open()
            'エラーメッセージ初期化
            puErrMsg = System.String.Empty

            'SELECT用SQLコマンドを作成
            If SqlEXTM0201.SetSelectRiyosha(Adapter, Cn, dataEXTM0201) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用者一覧検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得データをDataクラスへ保存
            dataEXTM0201.PropVwList.DataSource = Table

            'O件の場合
            If dataEXTM0201.PropVwList.ActiveSheet.RowCount = 0 Then
                ClearSpreadRow(dataEXTM0201)
            End If

            'セルタイプ変更（確認・編集ボタン表示）
            Dim Btn_Hensyu As New CellType.ButtonCellType
            'ボタン名
            Btn_Hensyu.Text = "確認・編集"
            dataEXTM0201.PropVwList.Sheets(0).Columns(13).CellType = Btn_Hensyu

            '終了ログ
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>検索結果表示
    ''' </summary>
    ''' <param name="dataEXTM0201"></param>
    ''' <returns></returns>
    ''' <remarks>検索結果を表示する
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function ViewRow(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            ' スプレッドの描画を停止
            dataEXTM0201.PropVwList.SuspendLayout()

            ' 全ての行を非表示にする
            For Each row As FarPoint.Win.Spread.Row In dataEXTM0201.PropVwList.ActiveSheet.Rows
                row.Visible = False
            Next

            '行を表示する
            For Each row As FarPoint.Win.Spread.Row In dataEXTM0201.PropVwList.ActiveSheet.Rows
                row.Visible = True
            Next

            '終了ログ
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False
        Finally
            ' スプレッドの描画を再開
            dataEXTM0201.PropVwList.ResumeLayout(True)
        End Try
    End Function

    ''' <summary>
    ''' 一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTM0201">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/11/30 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickVwCellMain(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスをラジオボタンのように制御する
            If SetCheckAsRadio(dataEXTM0201) = False Then
                Return False
            End If

            '選択フラグON
            'dataEXTM0201.PropBlnSelected = True

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' チェックボックス疑似ラジオボタン制御処理
    ''' <paramref name="dataHBKC0701">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/11/30 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckAsRadio(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTM0201

                'ON状態のチェックボックスを選択した場合、自分自身のチェックを外す
                'OFF状態のチェックボックスを選択した場合、選択行にはチェックをつけ、それ以外にはチェックを外す
                For i As Integer = 0 To .PropVwList.ActiveSheet.RowCount - 1

                    'ON状態のチェックボックスを選択した場合
                    If .PropVwList.ActiveSheet.GetValue(i, COL_SELECT) = True Then
                        'チェックを外す
                        .PropVwList.ActiveSheet.SetValue(i, COL_SELECT, False)

                    Else
                        'チェックを外す
                        .PropVwList.ActiveSheet.SetValue(i, COL_SELECT, False)

                        '選択行にチェックをつける
                        If i = .PropCheckIndex Then
                            .PropVwList.ActiveSheet.SetValue(i, COL_SELECT, True)
                        End If

                    End If

                Next

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' スプレッドシートのすべての行を削除する
    ''' </summary>
    ''' <param name="dataEXTM0201">DataEXTM0201型オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function ClearSpreadRow(ByRef dataEXTM0201 As DataEXTM0201) As Boolean
        ' 開始ログ出力
        Common.CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            If dataEXTM0201.PropVwList.ActiveSheet.RowCount > 0 Then
                dataEXTM0201.PropVwList.ActiveSheet.RemoveRows(0, dataEXTM0201.PropVwList.ActiveSheet.RowCount)
            End If

            ' 終了ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

            '例外処理
        Catch ex As Exception
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False
        End Try

    End Function

End Class
