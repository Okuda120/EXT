Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTZ0102
''' </summary>
''' <remarks>請求一覧画面で発生する情報の取得処理等を行う
''' <para>作成情報：2015/09/15 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTZ0102

    '変数宣言
    Private sqlEXTZ0102 As New SqlEXTZ0102              'sqlクラス

#Region "スプレッドの作成"

    ''' <summary>
    ''' スプレッドの作成
    ''' </summary>
    ''' <param name="dataEXTZ0102">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドを作成する。
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function MakeSpread(ByRef dataEXTZ0102 As DataEXTZ0102) As Boolean

        '変数宣言
        Dim intlastRow As Integer = 0       '請求一覧最終行

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            With dataEXTZ0102

                'データ取得
                If GetSpreadData(dataEXTZ0102) = False Then
                    Return False
                End If

                '一覧表示
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合

                    '---一覧クリア
                    If .PropvwSeikyuTheatre.ActiveSheet.Rows.Count > 0 Then
                        .PropvwSeikyuTheatre.Sheets(0).RemoveRows(0, .PropvwSeikyuTheatre.ActiveSheet.Rows.Count)
                    End If

                    '---一覧表示
                    intlastRow = .PropvwSeikyuTheatre.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtSeikyu.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwSeikyuTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '請求依頼番号
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_SeikyuIraiNo).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_IRAI_NO").ToString
                        '請求日
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_SeikyuDt).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_DT").ToString
                        '利用開始日
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_RiyoDtFrom).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_DT_FROM").ToString
                        '利用終了日
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_RiyoDtTo).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_DT_TO").ToString
                        '催事名
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_SaijiNm).Value = .PropDtSeikyu.Rows(intDataRow)("SAIJI_NM").ToString
                        '相手先名
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_AitesakiNm).Value = .PropDtSeikyu.Rows(intDataRow)("AITE_NM").ToString
                        '請求金額
                        If Not .PropDtSeikyu.Rows(intDataRow)("SEIKYU_KIN") Is DBNull.Value Then
                            .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_SeikyuKingaku).Value = Integer.Parse(.PropDtSeikyu.Rows(intDataRow)("SEIKYU_KIN")).ToString("#,##0")
                        End If
                        '請求内容
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_SeikyuNaiyo).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_NAIYO").ToString
                        '入金予定日 
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_NyukinYoteiDt).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_YOTEI_DT").ToString
                        '入金日 
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_NyukinDt).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_DT").ToString
                        ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
                        '入金額 
                        If Not .PropDtSeikyu.Rows(intDataRow)("NYUKIN_KIN") Is DBNull.Value Then
                            .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_NyukinKin).Value = Integer.Parse(.PropDtSeikyu.Rows(intDataRow)("NYUKIN_KIN")).ToString("#,##0")
                        End If
                        '入金状況 
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_NyukinSts).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_INPUT_FLG").ToString
                        ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
                        '予約番号
                        .PropvwSeikyuTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Theatre_YoyakuNo).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_NO").ToString

                        intlastRow += 1
                    Next
                    ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
                    'スクロールバーを必要な場合のみ表示させます
                    .PropvwSeikyuTheatre.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    .PropvwSeikyuTheatre.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
                Else
                    'スタジオの場合

                    '---一覧クリア
                    If .PropvwSeikyuStudio.ActiveSheet.Rows.Count > 0 Then
                        .PropvwSeikyuStudio.Sheets(0).RemoveRows(0, .PropvwSeikyuStudio.ActiveSheet.Rows.Count)
                    End If

                    '---一覧表示
                    intlastRow = .PropvwSeikyuStudio.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtSeikyu.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwSeikyuStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '請求依頼番号
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_SeikyuIraiNo).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_IRAI_NO").ToString
                        '請求日
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_SeikyuDt).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_DT").ToString
                        '利用開始日
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_RiyoDtFrom).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_DT_FROM").ToString
                        '利用終了日
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_RiyoDtTo).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_DT_TO").ToString
                        'スタジオ
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_Studio).Value = .PropDtSeikyu.Rows(intDataRow)("STUDIO").ToString
                        'アーティスト名
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_ShutsuenNm).Value = .PropDtSeikyu.Rows(intDataRow)("SHUTSUEN_NM").ToString
                        '相手先名
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_AitesakiNm).Value = .PropDtSeikyu.Rows(intDataRow)("AITE_NM").ToString
                        '請求金額
                        If Not .PropDtSeikyu.Rows(intDataRow)("SEIKYU_KIN") Is DBNull.Value Then
                            .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_SeikyuKingaku).Value = Integer.Parse(.PropDtSeikyu.Rows(intDataRow)("SEIKYU_KIN")).ToString("#,##0")
                        End If
                        '請求内容
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_SeikyuNaiyo).Value = .PropDtSeikyu.Rows(intDataRow)("SEIKYU_NAIYO").ToString
                        '入金予定日 
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_NyukinYoteiDt).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_YOTEI_DT").ToString
                        '入金日 
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_NyukinDt).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_DT").ToString
                        ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
                        '入金額 
                        If Not .PropDtSeikyu.Rows(intDataRow)("NYUKIN_KIN") Is DBNull.Value Then
                            .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_NyukinKin).Value = Integer.Parse(.PropDtSeikyu.Rows(intDataRow)("NYUKIN_KIN")).ToString("#,##0")
                        End If
                        '入金状況
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_NyukinSts).Value = .PropDtSeikyu.Rows(intDataRow)("NYUKIN_INPUT_FLG").ToString
                        ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
                        '予約番号
                        .PropvwSeikyuStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuIndex_Studio_YoyakuNo).Value = .PropDtSeikyu.Rows(intDataRow)("YOYAKU_NO").ToString

                        intlastRow += 1
                    Next
                    ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
                    'スクロールバーを必要な場合のみ表示させます
                    .PropvwSeikyuStudio.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    .PropvwSeikyuStudio.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
                    ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
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
    ''' <param name="dataEXTZ0102">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function GetSpreadData(ByRef dataEXTZ0102 As DataEXTZ0102) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0102.setSelectSeikyu(Adapter, Cn, dataEXTZ0102) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTZ0102.PropDtSeikyu = New DataTable
            Adapter.Fill(dataEXTZ0102.PropDtSeikyu)

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


End Class
