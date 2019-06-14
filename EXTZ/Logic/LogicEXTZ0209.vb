Imports Common
Imports CommonEXT
Imports Npgsql

Public Class LogicEXTZ0209
    ' --- 2019/06/13 軽減税率対応 Start E.Okuda@Compass ---
    ' 列番号定数化
    Private Const COL_FUTAI_SELECTED_FUTAI_SU As Integer = 3
    Private Const COL_FUTAI_SELECTED_FUTAI_CHOSEI As Integer = 6
    Private Const COL_FUTAI_SELECTED_FUTAI_KIN As Integer = 7
    Private Const COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD As Integer = 12
    Private Const COL_FUTAI_SELECTED_FUTAI_CD As Integer = 13
    ' --- 2019/06/13 軽減税率対応 End E.Okuda@Compass ---

    Private sqlEXTZ0209 As New sqlEXTZ0209          'sqlクラス
    Private commonLogicEXT As New CommonLogicEXT

    ''' <summary>
    ''' 分類取得
    ''' </summary>
    ''' <param name="dataEXTZ0209"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBunrui(ByRef dataEXTZ0209 As DataEXTZ0209) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0209.selectBunrui(Adapter, Cn, dataEXTZ0209) = False Then
                '異常終了
                Return False
            End If
            Adapter.Fill(Table)
            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "分類0件", Nothing, Adapter.SelectCommand)
                'Falseを返す
                Return False
            End If
            dataEXTZ0209.PropFutaiBunruiTable = Table
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "分類取得", Nothing, Adapter.SelectCommand)
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 分類取得
    ''' </summary>
    ''' <param name="dataEXTZ0209"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFutai(ByRef dataEXTZ0209 As DataEXTZ0209) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0209.selectFutaiMst(Adapter, Cn, dataEXTZ0209) = False Then
                '異常終了
                Return False
            End If
            Adapter.Fill(Table)
            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯設備マスタ0件", Nothing, Adapter.SelectCommand)
                'Falseを返す
                Return False
            End If
            dataEXTZ0209.PropFutaiMstTable = Table
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "分類取得", Nothing, Adapter.SelectCommand)
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 消費税取得処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>SYSTEM PROPERTY
    ''' <para>作成情報：2015.11.13 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetTax(ByVal Riyobi As String) As Double

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand
        Dim Table As New DataTable()

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0209.selectTax(Adapter, Cn, Riyobi) = False Then
                '異常終了
                Return 0
            End If
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "消費税取得失敗", Nothing, Adapter.SelectCommand)
                Return 0
            End If

            '取得した項目を格納する
            Dim tax As Double = Double.Parse(IIf(DBNull.Value.Equals(Table.Rows(0).Item("TAX_RITU")), Nothing, Table.Rows(0).Item("TAX_RITU"))) / 100
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "TAX = " & tax.ToString, Nothing, Nothing)
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return tax
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return 0
        Finally
            '終了処理
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック処理
    ''' <paramref name="dataEXTZ0209">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上の選択および値入力のチェックを行う
    ''' <para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InputCheck(ByRef dataEXTZ0209 As DataEXTZ0209)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 変数宣言
        Dim intSetCnt As Integer = 0
        Dim intInpCnt As Integer = 0
        Dim lngAllkin As Long = 0                                  ' 2015.12.21 ADD h.hagiwara

        Try

            With dataEXTZ0209.PropVwFutaiSelecSheet
                For i = 0 To .Sheets(0).RowCount - 1
                    ' --- 2019/06/13 軽減税率対応 Start E.Okuda@Compass ---
                    ' 列番号定数化
                    ' 数量チェック
                    If Not IsNumeric(.Sheets(0).Cells(i, COL_FUTAI_SELECTED_FUTAI_SU).Value) = True Then
                        puErrMsg = String.Format(CommonDeclareEXT.E0003, "数量")
                        Return False
                    End If
                    ' 調整額チェック
                    If Not IsNumeric(.Sheets(0).Cells(i, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value) = True Then
                        puErrMsg = String.Format(CommonDeclareEXT.E0003, "調整額")
                        Return False
                    End If

                    '' 数量チェック
                    'If Not IsNumeric(.Sheets(0).Cells(i, 3).Value) = True Then
                    '    puErrMsg = String.Format(CommonDeclareEXT.E0003, "数量")
                    '    Return False
                    'End If
                    '' 調整額チェック
                    'If Not IsNumeric(.Sheets(0).Cells(i, 6).Value) = True Then
                    '    puErrMsg = String.Format(CommonDeclareEXT.E0003, "調整額")
                    '    Return False
                    'End If
                    ' --- 2019/06/13 軽減税率対応 End E.Okuda@Compass ---
                Next

                ' 同一付帯設備情報が選択されている場合エラー
                For j = 0 To .Sheets(0).RowCount - 1
                    For k = j + 1 To .Sheets(0).RowCount - 1
                        ' --- 2019/06/13 軽減税率対応 Start E.Okuda@Compass ---
                        ' 列番号定数化
                        If .Sheets(0).Cells(j, COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD).Value = .Sheets(0).Cells(k, COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD).Value And
                           .Sheets(0).Cells(j, COL_FUTAI_SELECTED_FUTAI_CD).Value = .Sheets(0).Cells(k, COL_FUTAI_SELECTED_FUTAI_CD).Value Then
                            puErrMsg = CommonDeclareEXT.E2042
                            Return False
                        End If

                        '' 列番号定数化
                        'If .Sheets(0).Cells(j, 9).Value = .Sheets(0).Cells(k, 9).Value And
                        '   .Sheets(0).Cells(j, 10).Value = .Sheets(0).Cells(k, 10).Value Then
                        '    puErrMsg = CommonDeclareEXT.E2042
                        '    Return False
                        'End If
                        ' --- 2019/06/13 軽減税率対応 End E.Okuda@Compass ---
                    Next
                Next

                '2015.12.21 ADD START↓ h.hagiwara
                ' --- 2019/06/13 軽減税率対応 Start E.Okuda@Compass ---
                ' 列番号定数化
                .ActiveSheet.ColumnFooter.SetAggregationType(0, COL_FUTAI_SELECTED_FUTAI_KIN, FarPoint.Win.Spread.Model.AggregationType.Sum)
                lngAllkin = commonLogicEXT.convLong(.Sheets(0).ColumnFooter.Cells(0, COL_FUTAI_SELECTED_FUTAI_KIN).Value)
                '.ActiveSheet.ColumnFooter.SetAggregationType(0, 7, FarPoint.Win.Spread.Model.AggregationType.Sum)
                'lngAllkin = commonLogicEXT.convLong(.Sheets(0).ColumnFooter.Cells(0, 7).Value)
                ' --- 2019/06/13 軽減税率対応 End E.Okuda@Compass ---
                If lngAllkin.ToString.Length > 11 Then
                    puErrMsg = String.Format(CommonDeclareEXT.E2049, "付帯設備")
                    Return False
                End If

                lngAllkin += dataEXTZ0209.PropLngTotal
                If lngAllkin.ToString.Length > 11 Then
                    puErrMsg = String.Format(CommonDeclareEXT.E2049, "予約番号の付帯設備")
                    Return False
                End If
                '2015.12.21 ADD END↑ h.hagiwara

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
