Imports Common
Imports CommonEXT
Imports Npgsql

Public Class LogicEXTZ0208

    Private sqlEXTZ0208 As New SqlEXTZ0208          'sqlクラス

    '定数宣言
    Public Const COL_SELECT As Integer = 0            '選択 

    ''' <summary>
    ''' プロジェクト検索
    ''' </summary>
    ''' <param name="dataEXTZ0208"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetEXASNyukin(ByRef dataEXTZ0208 As DataEXTZ0208) As Boolean
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
            If sqlEXTZ0208.searchEXASNyukin(Adapter, Cn, dataEXTZ0208) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXAS入金検索", Nothing, Adapter.SelectCommand)
            '予約情報を取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            dataEXTZ0208.PropDtResult = Table
            Cn.Close()
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
    ''' スプレッドシートチェックボックスクリック時メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0208">[IN/OUT]セット選択画面データクラス</param>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッドシートのチェックボックス状態を制御する
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function ClickResultCellMain(ByRef dataEXTZ0208 As DataEXTZ0208) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'チェックボックスを制御する
            If SetCheck(dataEXTZ0208) = False Then
                Return False
            End If


            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' チェックボックス制御処理
    ''' <paramref name="dataEXTZ0208">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>既にチェックの入っている行のチェックを外し、選択行のチェックをつける
    ''' <para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetCheck(ByRef dataEXTZ0208 As DataEXTZ0208) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTZ0208.PropResult.ActiveSheet

                '選択行にチェックをつけ、それ以外はチェックを外す
                For i As Integer = 0 To .RowCount - 1

                    '選択行がすでにチェック済みの場合
                    If .GetValue(i, COL_SELECT) = True Then
                        'チェックを外す
                        .SetValue(i, COL_SELECT, False)
                    Else
                        'チェックを外す
                        .SetValue(i, COL_SELECT, False)
                        'チェックを付ける
                        If i = dataEXTZ0208.PropIntCheckRow Then
                            .SetValue(i, COL_SELECT, True)
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
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 入力チェック処理
    ''' <paramref name="dataEXTZ0208">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>スプレッド上の選択および値入力のチェックを行う
    ''' <para>作成情報：2016.08.16 m.hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InputCheck(ByRef dataEXTZ0208 As DataEXTZ0208)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言

        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()
            ' 登録済の請求入金表に存在するか判定 '2016.08.16 m.hayabuchi add
            If GetDupData(Adapter, Cn, dataEXTZ0208) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function


    ''' <summary>
    ''' 請求入金情報表取得：重複チェック用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="dataEXTZ0208">[IN/OUT]インシデント登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>画面で選択した入金情報が引き当て済みか件数取得を行う
    ''' <para>作成情報：2016.08.16 m.hayabuchi
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetDupData(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTZ0208 As DataEXTZ0208) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTZ0208.searchDup(Adapter, Cn, dataEXTZ0208) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "請求入金情報表取得（重複チェック）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            '選択行の入金リンクNOが請求入金票に一件以上存在する場合
            If dtmst.Rows.Count > 0 Then
                '変数宣言
                Dim y_nm As String
                Dim s_nm As String
                Dim r_nm As String
                Dim r_dy As String
                Dim i_no As String
                Dim i_co As String
                Dim i_mo As String

                y_nm = dtmst.Rows(0).Item("YOYAKU_NO").ToString
                s_nm = dtmst.Rows(0).Item("SAIJI_NM").ToString
                r_nm = dtmst.Rows(0).Item("RIYO_NM").ToString
                r_dy = dtmst.Rows(0).Item("YOYAKU_DAY").ToString
                i_no = dtmst.Rows(0).Item("SEIKYU_IRAI_NO").ToString
                If dtmst.Rows(0).Item("SEIKYU_NAIYO").ToString = "1" Then
                    i_co = "利用料"
                ElseIf dtmst.Rows(0).Item("SEIKYU_NAIYO").ToString = "2" Then
                    i_co = "付帯設備"
                ElseIf dtmst.Rows(0).Item("SEIKYU_NAIYO").ToString = "3" Then
                    i_co = "利用料+付帯設備"
                Else
                    i_co = " "
                End If
                i_mo = dtmst.Rows(0).Item("SEIKYU_KIN").ToString
                'エラーメッセージ設定
                puErrMsg = "選択した入金情報は、既に他の請求に引当られています。" & vbCrLf & vbCrLf &
                           "予約番号       : " & y_nm & vbCrLf &
                           "催事名          : " & s_nm & vbCrLf &
                           "利用者名       : " & r_nm & vbCrLf &
                           "利用日          : " & r_dy & vbCrLf &
                           "請求依頼番号 : " & i_no & vbCrLf &
                           "請求内容       : " & i_co & vbCrLf &
                           "請求金額       : " & i_mo
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            dtmst.Dispose()
        End Try

    End Function
End Class
