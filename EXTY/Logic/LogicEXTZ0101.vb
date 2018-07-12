Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTZ0101
''' </summary>
''' <remarks>予約一覧画面で発生する情報の取得処理等を行う
''' <para>作成情報：2015/09/09 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTZ0101

    '変数宣言
    Private sqlEXTZ0101 As New SqlEXTZ0101              'sqlクラス

    '定数宣言
    Public Const COL_SELECT As Integer = 0              '選択

#Region "スプレッドの作成"

    ''' <summary>
    ''' スプレッドの作成
    ''' </summary>
    ''' <param name="dataEXTZ0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドを作成する。
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function MakeSpread(ByRef dataEXTZ0101 As DataEXTZ0101) As Boolean

        '変数宣言
        Dim intlastRow As Integer = 0       '予約一覧最終行

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            With dataEXTZ0101

                'データ取得
                If GetSpreadData(dataEXTZ0101) = False Then
                    Return False
                End If

                '一覧表示
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合

                    '---一覧クリア
                    If .PropvwYoyakuTheatre.ActiveSheet.Rows.Count > 0 Then
                        .PropvwYoyakuTheatre.Sheets(0).RemoveRows(0, .PropvwYoyakuTheatre.ActiveSheet.Rows.Count)
                    End If

                    '---一覧表示
                    intlastRow = .PropvwYoyakuTheatre.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtYoyaku.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwYoyakuTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_YoyakuNo).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_NO").ToString
                        'ｽﾃｰﾀｽ
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_Sts).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_STS").ToString
                        '利用開始日
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_RiyoDtFrom).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_DT_FROM").ToString
                        '利用終了日
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_RiyoDtTo).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_DT_TO").ToString
                        '開始時間
                        If Not .PropDtYoyaku.Rows(intDataRow)("START_TIME") Is DBNull.Value AndAlso .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString <> String.Empty Then
                            .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_StartTime).Value = _
                                .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString.Substring(0, 2) & ":" & _
                                .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString.Substring(2, 2)
                        End If
                        '終了時間
                        If Not .PropDtYoyaku.Rows(intDataRow)("END_TIME") Is DBNull.Value AndAlso .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString <> String.Empty Then
                            .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_EndTime).Value = _
                                .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString.Substring(0, 2) & ":" & _
                                .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString.Substring(2, 2)
                        End If
                        '催事名
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_SaijiNm).Value = .PropDtYoyaku.Rows(intDataRow)("SAIJI_NM").ToString
                        '利用形状
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_RiyoType).Value = .PropDtYoyaku.Rows(intDataRow)("RIYO_TYPE").ToString
                        '利用者名
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_RiyoNm).Value = .PropDtYoyaku.Rows(intDataRow)("RIYO_NM").ToString
                        '責任者名
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_SekininNm).Value = .PropDtYoyaku.Rows(intDataRow)("SEKININ_NM").ToString
                        '承認人数
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_ShoninNinzu).Value = .PropDtYoyaku.Rows(intDataRow)("SHONIN_NINZU").ToString
                        '利用料
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_RiyoKin).Value = .PropDtYoyaku.Rows(intDataRow)("RIYO_KIN").ToString
                        '設備入力ステータス
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_FInputSts).Value = .PropDtYoyaku.Rows(intDataRow)("FINPUT_STS").ToString
                        '付帯設備
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_FutaiSetubi).Value = .PropDtYoyaku.Rows(intDataRow)("FUTAI_SETUBI").ToString
                        'STS_CD
                        .PropvwYoyakuTheatre.ActiveSheet.Cells(intlastRow, SpreadIndex_Theatre_StsCd).Value = .PropDtYoyaku.Rows(intDataRow)("STS_CD").ToString

                        intlastRow += 1
                    Next
                    'スクロールバーを必要な場合のみ表示させます
                    .PropvwYoyakuTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    .PropvwYoyakuTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                Else
                    'スタジオの場合

                    '---一覧クリア
                    If .PropvwYoyakuStudio.ActiveSheet.Rows.Count > 0 Then
                        .PropvwYoyakuStudio.Sheets(0).RemoveRows(0, .PropvwYoyakuStudio.ActiveSheet.Rows.Count)
                    End If

                    '---一覧表示
                    intlastRow = .PropvwYoyakuStudio.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtYoyaku.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwYoyakuStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_YoyakuNo).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_NO").ToString
                        'ｽﾃｰﾀｽ
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_Sts).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_STS").ToString
                        'スタジオ
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_Studio).Value = .PropDtYoyaku.Rows(intDataRow)("STUDIO").ToString
                        '利用開始日
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_RiyoDtFrom).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_DT_FROM").ToString
                        '利用終了日
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_RiyoDtTo).Value = .PropDtYoyaku.Rows(intDataRow)("YOYAKU_DT_TO").ToString
                        '開始時間
                        If Not .PropDtYoyaku.Rows(intDataRow)("START_TIME") Is DBNull.Value AndAlso .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString <> String.Empty Then
                            .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_StartTime).Value = _
                                .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString.Substring(0, 2) & ":" & _
                                .PropDtYoyaku.Rows(intDataRow)("START_TIME").ToString.Substring(2, 2)
                        End If
                        '終了時間
                        If Not .PropDtYoyaku.Rows(intDataRow)("END_TIME") Is DBNull.Value AndAlso .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString <> String.Empty Then
                            .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_EndTime).Value = _
                                .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString.Substring(0, 2) & ":" & _
                                .PropDtYoyaku.Rows(intDataRow)("END_TIME").ToString.Substring(2, 2)
                        End If
                        'アーティスト名
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_ShutsuenNm).Value = .PropDtYoyaku.Rows(intDataRow)("SHUTSUEN_NM").ToString
                        '利用者名
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_RiyoNm).Value = .PropDtYoyaku.Rows(intDataRow)("RIYO_NM").ToString
                        '責任者名
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_SekininNm).Value = .PropDtYoyaku.Rows(intDataRow)("SEKININ_NM").ToString
                        '承認人数
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_ShoninNinzu).Value = .PropDtYoyaku.Rows(intDataRow)("SHONIN_NINZU").ToString
                        '利用料
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_RiyoKin).Value = .PropDtYoyaku.Rows(intDataRow)("RIYO_KIN").ToString
                        '設備入力ステータス
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_FInputSts).Value = .PropDtYoyaku.Rows(intDataRow)("FINPUT_STS").ToString
                        '付帯設備
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_FutaiSetubi).Value = .PropDtYoyaku.Rows(intDataRow)("FUTAI_SETUBI").ToString
                        'STS_CD
                        .PropvwYoyakuStudio.ActiveSheet.Cells(intlastRow, SpreadIndex_Studio_StsCd).Value = .PropDtYoyaku.Rows(intDataRow)("STS_CD").ToString

                        intlastRow += 1
                    Next
                    'スクロールバーを必要な場合のみ表示させます
                    .PropvwYoyakuStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    .PropvwYoyakuStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

                End If

            End With

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

    '*****Privateメソッド*****
#Region "スプレッドへのデータ取得"

    ''' <summary>
    ''' スプレッドへのデータ取得
    ''' </summary>
    ''' <param name="dataEXTZ0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/09/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function GetSpreadData(ByRef dataEXTZ0101 As DataEXTZ0101) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0101.setSelectYoyaku(Adapter, Cn, dataEXTZ0101) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTZ0101.PropDtYoyaku = New DataTable
            Adapter.Fill(dataEXTZ0101.PropDtYoyaku)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        Finally
            If Cn IsNot Nothing Then
                Cn.Close()
            End If

            Adapter.Dispose()
            Cn.Dispose()

        End Try

    End Function

#End Region


    ''' <summary>
    ''' 一覧セルクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0101">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>一覧のチェックボックス状態を制御する
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickVwCellMain(ByRef dataEXTZ0101 As DataEXTZ0101) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスをラジオボタンのように制御する
            If SetCheckAsRadio(dataEXTZ0101) = False Then
                Return False
            End If

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
    ''' <paramref name="DataEXTZ0101">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/12/01 y.ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheckAsRadio(ByRef DataEXTZ0101 As DataEXTZ0101) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With DataEXTZ0101

                'シアター
                'ON状態のチェックボックスを選択した場合、自分自身のチェックを外す
                'OFF状態のチェックボックスを選択した場合、選択行にはチェックをつけ、それ以外にはチェックを外す
                For i As Integer = 0 To .PropvwYoyakuTheatre.ActiveSheet.RowCount - 1
                    'ON状態のチェックボックスを選択した場合
                    If .PropvwYoyakuTheatre.ActiveSheet.GetValue(i, COL_SELECT) = True Then
                        .PropvwYoyakuTheatre.ActiveSheet.SetValue(i, COL_SELECT, False)
                    Else
                        'チェックを外す
                        .PropvwYoyakuTheatre.ActiveSheet.SetValue(i, COL_SELECT, False)
                        '選択行にチェックをつける
                        If i = .PropIntCheckRow Then
                            .PropvwYoyakuTheatre.ActiveSheet.SetValue(i, COL_SELECT, True)
                        End If
                    End If
                Next

                'スタジオ
                'ON状態のチェックボックスを選択した場合、自分自身のチェックを外す
                'OFF状態のチェックボックスを選択した場合、選択行にはチェックをつけ、それ以外にはチェックを外す
                For i As Integer = 0 To .PropvwYoyakuStudio.ActiveSheet.RowCount - 1
                    'ON状態のチェックボックスを選択した場合
                    If .PropvwYoyakuStudio.ActiveSheet.GetValue(i, COL_SELECT) = True Then
                        .PropvwYoyakuStudio.ActiveSheet.SetValue(i, COL_SELECT, False)
                    Else
                        'チェックを外す
                        .PropvwYoyakuStudio.ActiveSheet.SetValue(i, COL_SELECT, False)
                        '選択行にチェックをつける
                        If i = .PropIntCheckRow Then
                            .PropvwYoyakuStudio.ActiveSheet.SetValue(i, COL_SELECT, True)
                        End If
                    End If
                Next

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function



End Class
