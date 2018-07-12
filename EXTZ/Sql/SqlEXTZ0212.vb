Imports Common
Imports CommonEXT
Imports Npgsql

Public Class SqlEXTZ0212
    'SQL文(利用料分類)
    Private strEX45S001 As String = _
                           "select " & vbCrLf & _
                           "    bunrui_cd, " & vbCrLf & _
                           "    bunrui_nm " & vbCrLf & _
                           "FROM " & vbCrLf & _
                           "    rbunrui_mst " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                           "    sts = '0' " & vbCrLf & _
                           "    AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "    AND :Riyobi BETWEEN kikan_from AND kikan_to" & vbCrLf & _
                           "ORDER BY " & vbCrLf & _
                           "    sort "

    'SQL文(利用料)
    Private strEX45S002 As String = _
                           "select " & vbCrLf & _
                           "    bunrui_cd, " & vbCrLf & _
                           "    ryokin_cd, " & vbCrLf & _
                           "    ryokin_nm, " & vbCrLf & _
                           "    ryokin_hour, " & vbCrLf & _
                           "    ryokin " & vbCrLf & _
                           "FROM " & vbCrLf & _
                           "    riyoryo_mst " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                           "    sts = '0' " & vbCrLf & _
                           "    AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "    AND bunrui_cd = :BunruiCd " & vbCrLf & _
                           "    AND :Riyobi BETWEEN kikan_from AND kikan_to" & vbCrLf & _
                           "ORDER BY " & vbCrLf & _
                           "    sort "

    'SQL文(倍率)
    Private strEX45S003 As String = _
                           "select " & vbCrLf & _
                           "    bairitu_cd, " & vbCrLf & _
                           "    bairitu_nm, " & vbCrLf & _
                           "    bairitu " & vbCrLf & _
                           "FROM " & vbCrLf & _
                           "    bairitu_mst " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                           "    sts = '0' " & vbCrLf & _
                           "    AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "    AND :Riyobi BETWEEN kikan_from AND kikan_to" & vbCrLf & _
                           "ORDER BY " & vbCrLf & _
                           "    sort "

    ''' <summary>
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectBunrui(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX45S001
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTZ0212.PropStrShisetu
                .Parameters(1).Value = dataEXTZ0212.PropStrSelectRiyobi
            End With

            Adapter.SelectCommand = Cmd
            Dim ds = dataEXTZ0212.PropDsTanakaBairitu
            Adapter.Fill(ds, "RBUNRUI_MST")
            '取得した件数が0件の場合
            If ds.Tables("RBUNRUI_MST").Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If
            dataEXTZ0212.PropDsTanakaBairitu = ds

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了

            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectRyokin(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTZ0212 As DataEXTZ0212, _
                                       ByVal strBunrui As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX45S002
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":BunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With

            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTZ0212.PropStrShisetu
                .Parameters(1).Value = strBunrui
                .Parameters(2).Value = dataEXTZ0212.PropStrSelectRiyobi
            End With

            Dim ds = dataEXTZ0212.PropDsTanakaBairitu
            Adapter.SelectCommand = Cmd
            Adapter.Fill(ds, "RIYORYO_MST")
            '取得した件数が0件の場合
            If ds.Tables("RIYORYO_MST").Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If
            dataEXTZ0212.PropDsTanakaBairitu = ds

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            Debug.WriteLine(ex.Message)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTZ0212"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectBairitu(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTZ0212 As DataEXTZ0212) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX45S003
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTZ0212.PropStrShisetu
                .Parameters(1).Value = dataEXTZ0212.PropStrSelectRiyobi
            End With

            Adapter.SelectCommand = Cmd
            Dim ds = dataEXTZ0212.PropDsTanakaBairitu
            Adapter.Fill(ds, "BAIRITU_MST")
            '取得した件数が0件の場合
            If ds.Tables("BAIRITU_MST").Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If
            dataEXTZ0212.PropDsTanakaBairitu = ds

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了

            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            '例外処理
            Return False
        End Try
    End Function

End Class
