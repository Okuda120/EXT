Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTB0101
''' </summary>
''' <remarks>その他施設利用登録（シアター）画面で発生する取得・更新処理等を行う
''' <para>作成情報：2015/08/27 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTB0105

    '変数宣言
    Private sqlEXTB0105 As New SqlEXTB0105              'sqlクラス

#Region "その他利用情報取得処理"

    ''' <summary>
    ''' その他利用情報取得処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>年度・回次データを取得する。
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetSonotaInfo(ByRef dataEXTB0105 As DataEXTB0105) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTB0105.setSelectSonotaInfo(Adapter, Cn, dataEXTB0105) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            '年度回次データ取得
            Adapter.Fill(Table)

            '取得した項目をDataクラスに格納する
            With dataEXTB0105
                .PropIntDataCnt = Table.Rows.Count
                If Table.Rows.Count > 0 Then
                    '開始時間
                    .PropStrStartTime = IIf(DBNull.Value.Equals(Table.Rows(0).Item("Start_Time")), "", Table.Rows(0).Item("Start_Time"))
                    '終了時間
                    .PropStrEndTime = IIf(DBNull.Value.Equals(Table.Rows(0).Item("End_Time")), "", Table.Rows(0).Item("End_Time"))
                    '場所
                    .PropStrRiyoBasho = IIf(DBNull.Value.Equals(Table.Rows(0).Item("Riyo_Basho")), "", Table.Rows(0).Item("Riyo_Basho"))
                    '用途
                    .PropStrRiyoYoto = IIf(DBNull.Value.Equals(Table.Rows(0).Item("Riyo_Yoto")), "", Table.Rows(0).Item("Riyo_Yoto"))
                    '利用者
                    .PropStrRiyosha = IIf(DBNull.Value.Equals(Table.Rows(0).Item("Riyosha")), "", Table.Rows(0).Item("Riyosha"))
                Else
                    .PropStrStartTime = String.Empty
                    .PropStrEndTime = String.Empty
                    .PropStrRiyoBasho = String.Empty
                    .PropStrRiyoYoto = String.Empty
                    .PropStrRiyosha = String.Empty
                End If
            End With

            Cn.Close()

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
            '終了処理
            Cn.Dispose()
            Table.Dispose()
            Adapter.Dispose()
        End Try
    End Function

#End Region

#Region "その他利用情報更新処理"

    ''' <summary>
    ''' その他利用情報更新処理
    ''' </summary>
    ''' <param name="dataEXTB0105">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>その他利用情報更新処理
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function UpdateSonotaInfo(ByRef dataEXTB0105 As DataEXTB0105) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '■シアターその他利用日程表更新
            With dataEXTB0105
                If .PropIntDataCnt > 0 Then
                    '更新処理
                    If sqlEXTB0105.setUpdateSonotaInfo(Cmd, Cn, dataEXTB0105) = False Then
                        '異常終了
                        Return False
                    End If
                Else
                    '新規登録処理
                    If sqlEXTB0105.setInsertSonotaInfo(Cmd, Cn, dataEXTB0105) = False Then
                        '異常終了
                        Return False
                    End If
                End If
            End With

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "祝祭日マスタ更新処理", Nothing, Cmd)

            '更新実行
            Cmd.Transaction = Tsx

            Cmd.ExecuteNonQuery()

            'コミット
            Tsx.Commit()

            Cn.Close()

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

#Region "その他利用情報削除処理"

    ''' <summary>
    ''' その他利用情報削除処理
    ''' </summary>
    ''' <param name="dataEXTB0105">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>その他利用情報削除処理
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DeleteSonotaInfo(ByRef dataEXTB0105 As DataEXTB0105) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '■シアターその他利用日程表削除
            With dataEXTB0105
                '削除処理
                If sqlEXTB0105.setDeleteSonotaInfo(Cmd, Cn, dataEXTB0105) = False Then
                    '異常終了
                    Return False
                End If
            End With

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "祝祭日マスタ更新処理", Nothing, Cmd)

            '更新実行
            Cmd.Transaction = Tsx

            Cmd.ExecuteNonQuery()

            'コミット
            Tsx.Commit()

            Cn.Close()

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

End Class
