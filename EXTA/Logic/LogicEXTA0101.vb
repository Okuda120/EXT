Imports Common
Imports CommonEXT
Imports Npgsql
Imports EXTZ
Imports System.Configuration

Public Class LogicEXTA0101

    '変数宣言
    Private sqlEXTA0101 As New SqlEXTA0101          'sqlクラス
    Private sqlEXTZ0202 As New SqlEXTZ0202          'sqlクラス

    ''' <summary>
    ''' ログインデータ取得処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>検索条件をもとにログインデータを取得する
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetLoginData(ByRef dataEXTA0101 As DataEXTA0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTA0101.setSelectLoginData(Adapter, Cn, dataEXTA0101) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "ログイン情報取得", Nothing, Adapter.SelectCommand)

            'ログインデータを取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            '取得した項目をDataクラスに格納する
            With dataEXTA0101
                'ユーザー名
                .PropStrUserName = IIf(DBNull.Value.Equals(Table.Rows(0).Item("USER_NM")), "", Table.Rows(0).Item("USER_NM"))
                .PropStrMailAddr = IIf(DBNull.Value.Equals(Table.Rows(0).Item("MAIL")), "", Table.Rows(0).Item("MAIL"))
                .PropStrCode_BUSHO = IIf(DBNull.Value.Equals(Table.Rows(0).Item("BUSHO_CD")), "", Table.Rows(0).Item("BUSHO_CD"))
                .PropStrFlg_SHOHIN = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SHONIN_FLG")), "", Table.Rows(0).Item("SHONIN_FLG"))
                .PropStrFlg_MST = IIf(DBNull.Value.Equals(Table.Rows(0).Item("MST_FLG")), "", Table.Rows(0).Item("MST_FLG"))
            End With

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            'ログインログ
            'CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "LOGIN:" & dataEXTA0101.PropTxtUserId, Nothing, Nothing)

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
    End Function

    ''' <summary>
    ''' ログイン成功時処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の削除
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function DelYoyakuCtlData(ByRef dataEXTA0101 As DataEXTA0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '予約制御データの削除処理
            Tsx = Cn.BeginTransaction
            sqlEXTZ0202.deleteYoyakuCtl(Cmd, Cn, CommonDeclareEXT.SHISETU_KBN_BLANK)
            Cmd.ExecuteNonQuery()
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "Login Proc1", Nothing, Cmd)
            sqlEXTZ0202.deleteCancelCtl(Cmd, Cn)
            Cmd.ExecuteNonQuery()
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "Login Proc2", Nothing, Cmd)

            Tsx.Commit()
            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function LoginInputCheck(ByVal dataEXTA0101 As DataEXTA0101) As Boolean

        'トレースログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'ユーザーIDが未入力の場合
            If dataEXTA0101.PropTxtUserId.Text = String.Empty Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "ユーザーID"), MsgBoxStyle.Exclamation, "エラー")

                '異常終了
                Return False
            End If

            'パスワードが未入力の場合
            If dataEXTA0101.PropTxtPassword.Text = String.Empty Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "パスワード"), MsgBoxStyle.Exclamation, "エラー")
                '異常終了
                Return False
            End If

            'トレースログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '異常終了
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ログイン成功時処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>SYSTEM PROPERTY
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetSystemProperties() As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'SELECT用SQLCommandを作成
            If sqlEXTA0101.setSelectSystemProperty(Adapter, Cn) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システムプロパティ取得", Nothing, Adapter.SelectCommand)

            'ログインデータを取得
            Adapter.Fill(Table)

            '取得した件数が0件の場合
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            '取得した項目をDataクラスに格納する
            CommonDeclareEXT.PropIntTeiinA = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TEIIN_A")), Nothing, Table.Rows(0).Item("TEIIN_A"))
            CommonDeclareEXT.PropIntTeiinB = IIf(DBNull.Value.Equals(Table.Rows(0).Item("TEIIN_B")), Nothing, Table.Rows(0).Item("TEIIN_B"))
            CommonDeclareEXT.PropIntKariJizenTuti = IIf(DBNull.Value.Equals(Table.Rows(0).Item("KARI_JIZEN_TUTI")), Nothing, Table.Rows(0).Item("KARI_JIZEN_TUTI"))
            CommonDeclareEXT.PropGSeikyusaki = IIf(DBNull.Value.Equals(Table.Rows(0).Item("GROUP_SEIKYU")), Nothing, Table.Rows(0).Item("GROUP_SEIKYU"))      ' 2016.01.18 ADD y.morooka グループ請求対応
            CommonDeclareEXT.PropStrDaikoTanto = IIf(DBNull.Value.Equals(Table.Rows(0).Item("DAIKO_TANTO_CD")), Nothing, Table.Rows(0).Item("DAIKO_TANTO_CD"))      ' 2016.09.20 m.hayabuchi ADD 請求依頼データ代行処理対応
            CommonDeclareEXT.PropStrDaikoBusho = IIf(DBNull.Value.Equals(Table.Rows(0).Item("DAIKO_BUSHO_CD")), Nothing, Table.Rows(0).Item("DAIKO_BUSHO_CD"))      ' 2016.09.20 m.hayabuchi ADD 請求依頼データ代行処理対応


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
    End Function

    ''' <summary>
    ''' 環境設定フラグ取得
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>入力・桁数の入力チェックを行う
    ''' <para>作成情報：2015.11.18 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetConfigrationFlg(ByVal a As Object) As Boolean

        'トレースログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            CommonEXT.PropConfigrationFlg = ConfigurationManager.AppSettings("configrationflg")

            'トレースログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '異常終了
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ログ出力設定（操作ログ）
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>ログの出力パスとファイル名の設定、及び、出力先パスへの接続を行う
    ''' <para>作成情報：2015.12.10 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetOpLog() As Boolean

        '※ログ出力設定前のためメソッド内でログ出力は行わない

        '変数宣言
        Dim strCmd As String = ""
        Dim OpLogPath As String = ""
        Dim OpLogUserId As String = ""
        Dim OpLogPass As String = ""
        'プロセスクラスの宣言
        Dim p As Process = Nothing                              'プロセスクラス
        Dim psi As New System.Diagnostics.ProcessStartInfo()    'プロセススタートインフォクラス

        Try
            OpLogPath = ConfigurationManager.AppSettings("OpLogPath")
            OpLogUserId = ConfigurationManager.AppSettings("OpLogUserId")
            OpLogPass = ConfigurationManager.AppSettings("OpLogPass")

            '操作ログ出力先パスが未設定の場合、出力先を変更しない
            If OpLogPath Is Nothing OrElse OpLogPath.Equals("") Then
                Return True
            End If

            'ログ出力先パスに接続
            Try
                psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
                psi.RedirectStandardInput = False
                psi.RedirectStandardOutput = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                '出力パスに接続できなければ認証
                If System.IO.Directory.Exists(OpLogPath) = False Then
                    'コマンドの設定・実行
                    If OpLogUserId Is Nothing OrElse OpLogUserId.Equals("") Then
                        OpLogUserId = OPLOG_USE_USERID
                        OpLogPass = OPLOG_USE_PASSWORD
                    End If
                    strCmd = "/C net use " & OpLogPath & " " & OpLogPass & " /user:" & OpLogUserId & " /persistent:no"
                    psi.Arguments = strCmd
                    p = Process.Start(psi)
                    p.WaitForExit()
                End If
            Catch ex As Exception
                puErrMsg = EXT_E001 & ex.Message
                Return False
            End Try

            'ログ出力フォルダに、操作ログ出力パスを設定（上書き）
            SystemLogPath = OpLogPath

            'ログファイル名に端末名を追加
            SystemLogFileName = System.Environment.MachineName & "." & SystemLogFileName

            '正常終了
            Return True
        Catch ex As Exception
            puErrMsg = EXT_E001 & ex.Message
            '異常終了
            Return False
        End Try
    End Function

    ''' <summary>
    ''' キャプチャフォルダ表示処理
    ''' </summary>
    ''' <returns>boolean エラーコード　true 正常終了　false 異常終了</returns>
    ''' <remarks>キャプチャフォルダをエクスプローラで表示する
    ''' <para>作成情報：2015.12.15 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function OpenCaptureDir() As Boolean

        '変数宣言
        Dim strCmd As String = ""
        Dim CapturePath As String = ""
        Dim CaptureUserId As String = ""
        Dim CapturePass As String = ""
        'プロセスクラスの宣言
        Dim p As Process = Nothing                              'プロセスクラス
        Dim psi As New System.Diagnostics.ProcessStartInfo()    'プロセススタートインフォクラス

        Try
            CapturePath = ConfigurationManager.AppSettings("CapturePath")
            CaptureUserId = ConfigurationManager.AppSettings("CaptureUserId")
            CapturePass = ConfigurationManager.AppSettings("CapturePass")

            'キャプチャフォルダに接続
            Try
                psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
                psi.RedirectStandardInput = False
                psi.RedirectStandardOutput = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                'キャプチャフォルダに接続できなければ認証
                If System.IO.Directory.Exists(CapturePath) = False Then
                    'コマンドの設定・実行
                    If CaptureUserId Is Nothing OrElse CaptureUserId.Equals("") Then
                        CaptureUserId = CAPTURE_USE_USERID
                        CapturePass = CAPTURE_USE_PASSWORD
                    End If
                    strCmd = "/C net use " & CapturePath & " " & CapturePass & " /user:" & CaptureUserId & " /persistent:no"
                    psi.Arguments = strCmd
                    p = Process.Start(psi)
                    p.WaitForExit()
                End If
                'エクスプローラで表示
                System.Diagnostics.Process.Start(CapturePath)
            Catch ex As Exception
                puErrMsg = EXT_E001 & ex.Message
                Return False
            End Try

            '正常終了
            Return True
        Catch ex As Exception
            puErrMsg = EXT_E001 & ex.Message
            '異常終了
            Return False
        End Try
    End Function

End Class
