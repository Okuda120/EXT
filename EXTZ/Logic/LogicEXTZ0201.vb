Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTZ0201
''' </summary>
''' <remarks>メンテ日登録画面で発生する情報の取得・更新処理等を行う
''' <para>作成情報：2015/08/28 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTZ0201

    '変数宣言
    Private sqlEXTZ0201 As New SqlEXTZ0201              'sqlクラス

#Region "休館メンテ情報取得処理"

    ''' <summary>
    ''' 休館メンテ情報取得処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>休館メンテ情報を取得する。
    ''' <para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetHolmentInfo(ByRef dataEXTZ0201 As DataEXTZ0201) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTZ0201.setSelectHolmentInfo(Adapter, Cn, dataEXTZ0201) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            Adapter.Fill(Table)

            '取得した項目をDataクラスに格納する
            With dataEXTZ0201
                .PropIntDataCnt = Table.Rows.Count
                If Table.Rows.Count > 0 Then
                    'メンテナンス内容
                    .PropStrMnaiyo = IIf(DBNull.Value.Equals(Table.Rows(0).Item("Mnaiyo")), String.Empty, Table.Rows(0).Item("Mnaiyo"))
                Else
                    .PropStrMnaiyo = String.Empty
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

            If Cn IsNot Nothing Then
                Cn.Close()
            End If

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

#Region "休館メンテ情報更新処理"

    ''' <summary>
    ''' 休館メンテ情報更新処理
    ''' </summary>
    ''' <param name="dataEXTZ0201">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>休館メンテ情報更新処理
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function UpdateHolmentInfo(ByRef dataEXTZ0201 As DataEXTZ0201) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション
        Dim intJikkoKaisu As Integer = 0                    '更新処理実行回数

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '■休館メンテマスタ更新
            With dataEXTZ0201
                If .PropIntDataCnt > 0 Then
                    '更新処理
                    If sqlEXTZ0201.setUpdateHolmentInfo(Cmd, Cn, dataEXTZ0201) = False Then
                        '異常終了
                        Return False
                    End If
                Else
                    '新規登録処理
                    If sqlEXTZ0201.setInsertHolmentInfo(Cmd, Cn, dataEXTZ0201) = False Then
                        '異常終了
                        Return False
                    End If
                End If
            End With

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "休館メンテマスタ更新処理", Nothing, Cmd)

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

            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            If Cn IsNot Nothing Then
                Cn.Close()
            End If

            '例外処理
            puErrMsg = ex.Message
            Return False

        Finally
            Cn.Dispose()
            Cmd.Dispose()
            Tsx.Dispose()

        End Try

    End Function

#End Region

#Region "休館メンテ情報更新処理（複数スタジオの場合）"

    ''' <summary>
    ''' 休館メンテ情報更新処理（複数スタジオの場合）
    ''' </summary>
    ''' <param name="dataEXTZ0201">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>休館メンテ情報更新処理（複数スタジオの場合）
    ''' <para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DelInsHolmentInfo(ByRef dataEXTZ0201 As DataEXTZ0201) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション
        Dim intJikkoKaisu As Integer = 0                    '更新処理実行回数

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '休館メンテマスタ更新
            '＜削除処理＞
            If sqlEXTZ0201.setDeleteHolmentInfo(Cmd, Cn, dataEXTZ0201) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "休館メンテマスタ削除処理", Nothing, Cmd)

            '削除処理実行
            Cmd.Transaction = Tsx

            Cmd.ExecuteNonQuery()

            '＜追加処理＞
            For intStudio As Integer = 1 To 2

                dataEXTZ0201.PropStrStudioKbn = intStudio.ToString
                If sqlEXTZ0201.setInsertHolmentInfo(Cmd, Cn, dataEXTZ0201) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "休館メンテマスタ追加処理", Nothing, Cmd)

                '追加処理実行
                Cmd.Transaction = Tsx

                Cmd.ExecuteNonQuery()

            Next

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "休館メンテマスタ更新処理", Nothing, Cmd)

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

            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            If Cn IsNot Nothing Then
                Cn.Close()
            End If

            '例外処理
            puErrMsg = ex.Message
            Return False

        Finally
            Cn.Dispose()
            Cmd.Dispose()
            Tsx.Dispose()

        End Try

    End Function

#End Region

End Class
