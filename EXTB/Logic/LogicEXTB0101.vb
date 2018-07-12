Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTB0101
''' </summary>
''' <remarks>予約カレンダー（シアター）で発生する取得・更新処理等を行う
''' <para>作成情報：2015/08/05 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTB0101

    '変数宣言
    Private commonLogic As New CommonLogic              '共通ロジッククラス
    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private sqlEXTB0101 As New SqlEXTB0101              'sqlクラス

    '*****Publicメソッド*****
#Region "スプレッドの作成"

    ''' <summary>
    ''' スプレッドの作成
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドを作成する。
    ''' <para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function MakeSpread(ByRef dataEXTB0101 As DataEXTB0101) As Boolean

        '変数宣言
        Dim datNewDate As Date = Nothing        '年月日を個別に指定した新しい日付
        Dim intMonthDayCnt As Integer = 0       '月の日数
        Dim txtDay As String = String.Empty     '日
        Dim txtYobi As String = String.Empty    '曜日

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            With dataEXTB0101
                'カレンダーの初期化
                datNewDate = New Date(.PropTxtYear.Text, .PropTxtMonth.Text, 1)
                .PropStrYear = .PropTxtYear.Text
                .PropStrMonth = .PropTxtMonth.Text
                '---データクリア
                For intColumn As Integer = 0 To .PropvwCalandarFirst.ActiveSheet.Columns.Count - 1
                    .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intColumn).Text = " "
                Next
                For intColumn As Integer = 0 To .PropvwCalandarSecond.ActiveSheet.Columns.Count - 1
                    .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intColumn).Text = " "
                Next
                .PropvwCalandarFirst.ActiveSheet.ClearRange(0, 0, .PropvwCalandarFirst.ActiveSheet.Rows.Count, .PropvwCalandarFirst.ActiveSheet.Columns.Count, True)
                .PropvwCalandarSecond.ActiveSheet.ClearRange(0, 0, .PropvwCalandarSecond.ActiveSheet.Rows.Count, .PropvwCalandarSecond.ActiveSheet.Columns.Count, True)
                '---背景色クリア、ロック設定
                Dim vwcellCarendar As FarPoint.Win.Spread.Cell
                For intRow As Integer = 0 To .PropvwCalandarFirst.ActiveSheet.RowCount - 1
                    For intColumn As Integer = 0 To .PropvwCalandarFirst.ActiveSheet.Columns.Count - 1
                        vwcellCarendar = .PropvwCalandarFirst.ActiveSheet.Cells(intRow, intColumn)
                        If intRow >= 3 And intRow <= 5 Then
                            vwcellCarendar.BackColor = Color.FromArgb(240, 240, 240)
                        Else
                            vwcellCarendar.ResetBackColor()
                        End If
                        vwcellCarendar.Locked = True
                    Next
                Next
                For intRow As Integer = 0 To .PropvwCalandarSecond.ActiveSheet.RowCount - 1
                    For intColumn As Integer = 0 To .PropvwCalandarSecond.ActiveSheet.Columns.Count - 1
                        vwcellCarendar = .PropvwCalandarSecond.ActiveSheet.Cells(intRow, intColumn)
                        If intRow >= 3 And intRow <= 5 Then
                            vwcellCarendar.BackColor = Color.FromArgb(240, 240, 240)
                        Else
                            vwcellCarendar.ResetBackColor()
                        End If
                        vwcellCarendar.Locked = True
                    Next
                Next
                '---日付曜日設定
                '<指定月の日数取得>
                intMonthDayCnt = DateTime.DaysInMonth(datNewDate.Year, datNewDate.Month)
                txtDay = String.Empty
                txtYobi = String.Empty
                '<カレンダーに日付曜日設定、ロック解除>
                For intDay As Integer = 1 To 15
                    txtDay = intDay.ToString
                    txtYobi = WeekdayName(Weekday(DateSerial(datNewDate.Year, datNewDate.Month, intDay)))
                    .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).Text = txtDay & "(" & txtYobi.Substring(0, 1) & ")"
                    .PropvwCalandarFirst.ActiveSheet.Cells(0, intDay - 1).Locked = False
                    If txtYobi.Substring(0, 1) = "土" Then
                        .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).ForeColor = Color.Blue
                    ElseIf txtYobi.Substring(0, 1) = "日" Then
                        .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).ForeColor = Color.Red
                    Else
                        .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).ForeColor = Color.Black
                    End If
                Next
                For intDay As Integer = 16 To intMonthDayCnt
                    txtDay = intDay.ToString
                    txtYobi = WeekdayName(Weekday(DateSerial(datNewDate.Year, datNewDate.Month, intDay)))
                    .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).Text = txtDay & "(" & txtYobi.Substring(0, 1) & ")"
                    .PropvwCalandarSecond.ActiveSheet.Cells(0, intDay - 16).Locked = False
                    If txtYobi.Substring(0, 1) = "土" Then
                        .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).ForeColor = Color.Blue
                    ElseIf txtYobi.Substring(0, 1) = "日" Then
                        .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).ForeColor = Color.Red
                    Else
                        .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).ForeColor = Color.Black
                    End If
                Next

                'キャンセル待ち一覧、キャンセル済一覧の初期化
                '---削除
                If .PropvwCancelWait.ActiveSheet.Rows.Count > 0 Then
                    .PropvwCancelWait.Sheets(0).RemoveRows(0, .PropvwCancelWait.ActiveSheet.Rows.Count)
                End If
                If .PropvwCancelDone.ActiveSheet.Rows.Count > 0 Then
                    .PropvwCancelDone.Sheets(0).RemoveRows(0, .PropvwCancelDone.ActiveSheet.Rows.Count)
                End If
                '---ロック設定
                .PropvwCancelWait.Sheets(0).Columns(0).Locked = True
                .PropvwCancelWait.Sheets(0).Columns(1).Locked = True
                .PropvwCancelWait.Sheets(0).Columns(2).Locked = True
                .PropvwCancelWait.Sheets(0).Columns(3).Locked = True
                .PropvwCancelDone.Sheets(0).Columns(0).Locked = True
                .PropvwCancelDone.Sheets(0).Columns(1).Locked = True
                .PropvwCancelDone.Sheets(0).Columns(2).Locked = True

                'データ取得
                .PropStrYear = .PropTxtYear.Text
                .PropStrMonth = Integer.Parse(.PropTxtMonth.Text).ToString("00")
                If GetSpreadData(dataEXTB0101) = False Then
                    Return False
                End If

                '一覧表示
                If ViewSpreadData(dataEXTB0101) = False Then
                    Return False
                End If

            End With

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "指定年月の入力チェック"

    ''' <summary>
    ''' 指定年月の入力チェック 
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>指定年月の入力チェックを行う。
    ''' <para>作成情報：2015/08/12 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function NengetuInputCheck(ByVal dataEXTB0101 As DataEXTB0101) As Boolean

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTB0101
                '必須チェック
                If .PropTxtYear.Text = String.Empty Then
                    MsgBox(String.Format(CommonEXT.E0001, "年"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If
                If .PropTxtMonth.Text = String.Empty Then
                    MsgBox(String.Format(CommonEXT.E0001, "月"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

                '属性チェック(半角数字)
                If commonValidate.IsHalfNmb(.PropTxtYear.Text) = False Then
                    MsgBox(String.Format(CommonEXT.E0003, "年"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If
                If commonValidate.IsHalfNmb(.PropTxtMonth.Text) = False Then
                    MsgBox(String.Format(CommonEXT.E0003, "月"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

                '桁数チェック
                If .PropTxtYear.Text.Length < 4 Then
                    MsgBox(String.Format(CommonEXT.E0010, "年", "4"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

                '存在しない年月の場合、エラーメッセージを出力
                If IsDate(.PropTxtYear.Text & "/" & .PropTxtMonth.Text & "/" & "01") = False Then
                    MsgBox(String.Format(CommonEXT.E2026), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If
            End With

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "選択日のチェック"

    ''' <summary>
    ''' 選択日のチェック（仮予約登録時）
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="DateLst">日付リスト</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/08/18 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function CheckSelectDateKariyoyaku(ByRef dataEXTB0101 As DataEXTB0101, ByVal DateLst As ArrayList) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ
        Dim intYoyakuCount As Integer = 0               '予約件数
        Dim dtCheck As New DataTable                    '各SQL実行毎の情報

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            For intDateLstCount As Integer = 0 To DateLst.Count - 1

                dataEXTB0101.PropStrYYMMDD = DateLst(intDateLstCount)
                intYoyakuCount = 0

                '■予約一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intYoyakuCount += dtCheck.Rows.Count

                '■予約未確定一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectMikakuYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intYoyakuCount += dtCheck.Rows.Count

                '■予約状況判定
                '※指定日のうち、いずれか1日でも仮予約または正式予約が2件行われている場合はエラー
                If intYoyakuCount >= 2 Then
                    MsgBox(String.Format(CommonEXT.E2017), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

                '■休館メンテ一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectKyukanMainteList(Adapter, Cn, dataEXTB0101, KYUKANMAINTE_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)

                '■休館メンテ状況判定
                If dtCheck.Rows.Count > 0 Then
                    MsgBox(String.Format(CommonEXT.E2027), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

            Next

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 選択日のチェック（休館・メンテ・営業日切り替え時）
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="DateLst">日付リスト</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/08/18 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function CheckSelectDateKyukanMainteEigyo(ByRef dataEXTB0101 As DataEXTB0101, ByVal DateLst As ArrayList, Optional ByVal Mode As String = "") As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim intYoyakuCount As Integer = 0
        Dim dtCheck As New DataTable

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            '■複数選択判定
            If Mode <> String.Empty Then
                If Mode <> SETDATE_EIGYO Then
                    If DateLst.Count >= 2 Then
                        MsgBox(String.Format(CommonEXT.E2029, Mode), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                        Return False
                    End If
                End If
            End If

            For Each strDate In DateLst
                dataEXTB0101.PropStrYYMMDD = strDate

                intYoyakuCount = 0

                '■予約一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intYoyakuCount += dtCheck.Rows.Count

                '■予約未確定一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectMikakuYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intYoyakuCount += dtCheck.Rows.Count

                '■予約状況判定
                '※指定日に仮予約または正式予約がある場合はエラー
                If intYoyakuCount > 0 Then
                    MsgBox(String.Format(CommonEXT.E2030), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

                '■休館メンテ一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectKyukanMainteList(Adapter, Cn, dataEXTB0101, KYUKANMAINTE_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)

                '■設定が可能であるかを判定
                If dtCheck.Rows.Count > 0 Then
                    '休館日・メンテ日のデータあり
                    If dtCheck.Rows(0)("HOLMENT_KBN") = "1" Then
                        '休館日の場合
                        If Mode = SETDATE_MAINTE Then
                            '設定日が休館日、メンテ日の場合はエラー
                            MsgBox(String.Format(CommonEXT.E2031, "休館日"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                            Return False
                        End If
                    Else
                        'メンテ日の場合
                        If Mode = SETDATE_KYUKAN Then
                            '設定日が休館日、メンテ日の場合はエラー
                            MsgBox(String.Format(CommonEXT.E2031, "メンテ日"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                            Return False
                        End If
                    End If
                Else
                    '休館日・メンテ日のデータなし
                    If Mode = SETDATE_EIGYO Then
                        '設定日が営業日の場合はエラー
                        MsgBox(String.Format(CommonEXT.E2031, "営業日"), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                        Return False
                    End If
                End If
            Next

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "祝祭日マスタチェック"

    ''' <summary>
    ''' 祝祭日マスタチェック
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="DateLst">日付リスト</param>
    ''' <param name="Mode">祝日設定モード</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function CheckHoliday(ByRef dataEXTB0101 As DataEXTB0101, ByVal DateLst As ArrayList, ByVal Mode As String) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim dtCheck As New DataTable                'SQL実行毎の情報

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            '■選択数分ループ
            For Each strDate In DateLst
                dataEXTB0101.PropStrYYMMDD = strDate

                '■祝祭日データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setJudgeHoliday(Adapter, Cn, dataEXTB0101) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)

                '■設定が可能であるかを判定
                If dtCheck.Rows.Count > 0 Then
                    '祝日データあり
                    If Mode = HOLIYDAY_ON Then
                        '設定日が祝日の場合はエラー
                        MsgBox(String.Format(CommonEXT.E2031, HOLIYDAY_ON), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                        Return False
                    End If
                Else
                    '祝日データなし
                    If Mode = HOLIYDAY_OFF Then
                        '設定日が平日の場合はエラー
                        MsgBox(String.Format(CommonEXT.E2031, HOLIYDAY_OFF), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                        Return False
                    End If
                End If
            Next

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "キャンセル待ちチェック"

    ''' <summary>
    ''' キャンセル待ちチェック
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="DateLst">日付リスト</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>キャンセル待ちチェックを行う
    ''' <para>作成情報：2015/08/18 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function CheckCancelWait(ByRef dataEXTB0101 As DataEXTB0101, ByVal DateLst As ArrayList) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネヌション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ
        Dim intCancelCount As Integer = 0               'キャンセル件数
        Dim dtCheck As New DataTable                    '各SQL実行毎の情報 

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            For intDateLstCount As Integer = 0 To DateLst.Count - 1

                dataEXTB0101.PropStrYYMMDD = DateLst(intDateLstCount)
                intCancelCount = 0

                '■キャンセル待ちデータ（日付確定のみ）取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectCancelList(Adapter, Cn, dataEXTB0101, CANCEL_WAITDATEONLY_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intCancelCount += dtCheck.Rows.Count


                '■予約未確定一覧データ取得
                'SELECT用SQLCommandを作成
                If sqlEXTB0101.setSelectMikakuCancelWaitList(Adapter, Cn, dataEXTB0101, CANCEL_WAITDATEONLY_NENGAPPI) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtCheck.Clear()
                Adapter.Fill(dtCheck)
                intCancelCount += dtCheck.Rows.Count

                '■キャンセル状況判定
                '※指定日のうち、いずれか１日でも最大キャンセル待ち数を超過時はエラー
                If intCancelCount = 3 Then
                    MsgBox(String.Format(CommonEXT.E2028), MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "エラー")
                    Return False
                End If

            Next

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try


    End Function

#End Region

#Region "営業日の設定"

    ''' <summary>
    ''' 営業日の設定
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="SelectDate">指定日</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>休館日・営業日への直接設定を行う
    ''' <para>作成情報：2015/08/18 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetEigyo(ByRef dataEXTB0101 As DataEXTB0101, ByVal SelectDate As String) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション
        Dim Adapter As New NpgsqlDataAdapter                'アダプタ
        Dim dtCheck As New DataTable                        '各SQL実行時の取得情報
        Dim bcCalendar As Color = Nothing                   'カレンダーセル文字色
        Dim intDay As Integer = 0                           '日

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            '指定日をデータクラスに格納
            dataEXTB0101.PropStrYYMMDD = SelectDate

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '■休館メンテマスタ更新
            '営業日に設定
            If sqlEXTB0101.setDeleteKyukanMainte(Cmd, Cn, dataEXTB0101) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "休館メンテマスタ更新処理", Nothing, Cmd)

            '更新実行
            Cmd.Transaction = Tsx

            Cmd.ExecuteNonQuery()

            'コミット
            Tsx.Commit()

            Cn.Close()

            '■背景色の設定、セルのクリア
            '営業日に設定
            bcCalendar = Color.White
            intDay = Integer.Parse(SelectDate.Substring(8, 2))
            If intDay <= 15 Then
                dataEXTB0101.PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).BackColor = bcCalendar
                dataEXTB0101.PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).Value = String.Empty
                dataEXTB0101.PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).BackColor = bcCalendar
            Else
                dataEXTB0101.PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).BackColor = bcCalendar
                dataEXTB0101.PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).Value = String.Empty
                dataEXTB0101.PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).BackColor = bcCalendar
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "祝日・平日の設定"

    ''' <summary>
    ''' 祝日・平日の設定
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <param name="SelectDate">指定日</param>
    ''' <param name="Mode">祝日設定モード</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>祝日・平日の設定
    ''' <para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetHolidayWeekday(ByRef dataEXTB0101 As DataEXTB0101, ByVal SelectDate As String, ByVal Mode As String) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)            'コネクション
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand         'コマンド
        Dim Tsx As NpgsqlTransaction = Nothing              'トランザクション
        Dim fcCalendar As Color = Nothing                   'カレンダーヘッダセル文字色
        Dim intDay As Integer = 0                           '日
        Dim strYobi As String = String.Empty                '曜日

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            '指定日をデータクラスに格納
            dataEXTB0101.PropStrYYMMDD = SelectDate

            Cn.Open()

            'トランザクションを開始
            Tsx = Cn.BeginTransaction

            '■祝祭日マスタ更新
            If Mode = HOLIYDAY_OFF Then
                '祝日に設定
                If sqlEXTB0101.setDeleteHoliday(Cmd, Cn, dataEXTB0101) = False Then
                    '異常終了
                    Return False
                End If
            Else
                '平日に設定
                If sqlEXTB0101.setInsertHoliday(Cmd, Cn, dataEXTB0101) = False Then
                    '異常終了
                    Return False
                End If
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "祝祭日マスタ更新処理", Nothing, Cmd)

            '更新実行
            Cmd.Transaction = Tsx

            Cmd.ExecuteNonQuery()

            'コミット
            Tsx.Commit()

            Cn.Close()

            '■ヘッダ文字色の設定
            intDay = Integer.Parse(SelectDate.Substring(8, 2))
            If Mode = HOLIYDAY_ON Then
                '祝日に設定
                fcCalendar = Color.Red
            Else
                '平日に設定（月～金、土、日の３パターンに分ける）
                strYobi = WeekdayName(Weekday(DateSerial(dataEXTB0101.PropStrYear, dataEXTB0101.PropStrMonth, intDay)))
                If strYobi.Substring(0, 1) = "土" Then
                    fcCalendar = Color.Blue
                ElseIf strYobi.Substring(0, 1) = "日" Then
                    fcCalendar = Color.Red
                Else
                    fcCalendar = Color.Black
                End If
            End If
            If intDay <= 15 Then
                dataEXTB0101.PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).ForeColor = fcCalendar
            Else
                dataEXTB0101.PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).ForeColor = fcCalendar
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

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
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function GetSpreadData(ByRef dataEXTB0101 As DataEXTB0101) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            '■祝祭日取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectShukusaijitsu(Adapter, Cn, dataEXTB0101) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtShukusaijitsu = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtShukusaijitsu)

            '■予約一覧データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGETU) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtYoyakuList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtYoyakuList)

            '■予約未確定一覧データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectMikakuYoyakuList(Adapter, Cn, dataEXTB0101, YOYAKU_NENGETU) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtMikakuYoyakuList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtMikakuYoyakuList)

            '■キャンセル待ちデータ（日付確定のみ）取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectCancelList(Adapter, Cn, dataEXTB0101, CANCEL_WAITDATEONLY) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtDateKakuOnlyCancelWaitList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtDateKakuOnlyCancelWaitList)

            '■キャンセル待ちデータ（全て）取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectCancelList(Adapter, Cn, dataEXTB0101, CANCEL_WAITALL) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtAllCancelWaitList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtAllCancelWaitList)

            '■キャンセル待ち未確定データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectMikakuCancelWaitList(Adapter, Cn, dataEXTB0101, CANCEL_WAITDATEONLY) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtMikakuCancelWaitList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtMikakuCancelWaitList)

            '■キャンセル済データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectCancelList(Adapter, Cn, dataEXTB0101, CANCEL_SUMI) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtCancelDoneList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtCancelDoneList)

            '■その他利用予約一覧データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectEtcUsedYoyakuList(Adapter, Cn, dataEXTB0101) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtEtcUseYoyakuList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtEtcUseYoyakuList)

            '■休館メンテ一覧データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTB0101.setSelectKyukanMainteList(Adapter, Cn, dataEXTB0101, KYUKANMAINTE_NENGETU) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTB0101.PropDtKyuykanMaiteList = New DataTable
            Adapter.Fill(dataEXTB0101.PropDtKyuykanMaiteList)

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "スプレッドへの一覧表示 "

    ''' <summary>
    ''' スプレッドへの一覧表示 
    ''' </summary>
    ''' <param name="dataEXTB0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへの一覧表示 を行う。
    ''' <para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function ViewSpreadData(ByRef dataEXTB0101 As DataEXTB0101) As Boolean

        '変数宣言
        Dim intDay As Integer = 0       '利用日等の日部分
        Dim strTheaterInfo As String = String.Empty         'シアター予約情報 
        Dim bcCalendar As Color = Nothing                   'カレンダーセル背景色
        Dim fcCalendar As Color = Nothing                   'カレンダーセル文字色
        Dim strCancelInfo As String = String.Empty          'キャンセル情報
        Dim strEtcInfo As String = String.Empty             'その他情報
        Dim strKyukanMainteInfo As String = String.Empty    '休館メンテ情報
        Dim intlastRow As Integer = 0                       'キャンセル一覧最終行
        Dim strYoyaku As String = String.Empty              '利用日
        Dim datAdd4MonthAfter As Date = Nothing             '4ヶ月後の日付
        Dim intYoyakuRow As Integer = 0                     '予約行
        Dim intCancelRow As Integer = 0                     'キャンセル行

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            With dataEXTB0101
                '■カレンダー表示
                '祝祭日表示
                For intDataRow As Integer = 0 To .PropDtShukusaijitsu.Rows.Count - 1
                    intDay = Integer.Parse(.PropDtShukusaijitsu.Rows(intDataRow)("HOLIDAY_DT").ToString.Substring(8, 2))
                    If intDay <= 15 Then
                        .PropvwCalandarFirst.ActiveSheet.ColumnHeader.Cells(0, intDay - 1).ForeColor = Color.Red
                    Else
                        .PropvwCalandarSecond.ActiveSheet.ColumnHeader.Cells(0, intDay - 16).ForeColor = Color.Red
                    End If
                Next

                '予約情報表示
                .PropDicYoyakuInfo = New Dictionary(Of String, String)
                For intDataRow As Integer = 0 To .PropDtYoyakuList.Rows.Count - 1
                    intDay = Integer.Parse(.PropDtYoyakuList.Rows(intDataRow)("YOYAKU_DT").ToString.Substring(8, 2))
                    strYoyaku = .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_DT").ToString
                    '---予約文字列取得
                    '町田修正
                    If .PropDtYoyakuList.Rows(intDataRow)("MITEI_FLG").ToString = "0" Then
                        strTheaterInfo = .PropDtYoyakuList.Rows(intDataRow)("START_TIME").ToString.Substring(0, 2) & ":" & _
                                         .PropDtYoyakuList.Rows(intDataRow)("START_TIME").ToString.Substring(2, 2) & _
                                         "-" & .PropDtYoyakuList.Rows(intDataRow)("END_TIME").ToString.Substring(0, 2) & ":" & _
                                         .PropDtYoyakuList.Rows(intDataRow)("END_TIME").ToString.Substring(2, 2) & vbCrLf & _
                                         .PropDtYoyakuList.Rows(intDataRow)("SAIJI_NM").ToString & vbCrLf & _
                                         .PropDtYoyakuList.Rows(intDataRow)("SHUTSUEN_NM").ToString
                    Else
                        strTheaterInfo = .PropDtYoyakuList.Rows(intDataRow)("SAIJI_NM").ToString & vbCrLf & _
                                         .PropDtYoyakuList.Rows(intDataRow)("SHUTSUEN_NM").ToString
                    End If
                    '---背景色取得
                    '---2016.1.15 START ↓y.ozawa 背景色変更
                    Select Case .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_STS").ToString
                        Case "1"
                            '仮予約
                            bcCalendar = Color.White
                        Case "2"
                            '予約
                            bcCalendar = Color.FromArgb(127, 255, 255)
                        Case "3"
                            '申請受諾済
                            bcCalendar = Color.FromArgb(191, 255, 127)
                        Case "4"
                            '精算完了
                            bcCalendar = Color.FromArgb(255, 255, 127)
                    End Select
                    '---2016.1.15 END ↑y.ozawa 背景色変更

                    '---利用日が現在と比べ4ヶ月未満であるかを確認し、文字色取得
                    datAdd4MonthAfter = DateSerial(Now.Year, Now.Month + 4, Now.Day)
                    ''If DateDiff("d", strYoyaku, datAdd4MonthAfter) > 0 Then
                    ''    '4ヶ月未満の場合
                    ''    fcCalendar = Color.Red
                    ''Else
                    ''    '4ヶ月以上の場合
                    ''    fcCalendar = Color.Black
                    ''End If
                    ' 2015.12.03 UPD START↓ h.hagiwara 仮予約のみ赤文字で表示する
                    If .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_STS").ToString = "1" Or _
                       .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_STS").ToString = "2" Then
                        If DateDiff("d", strYoyaku, datAdd4MonthAfter) > 0 Then
                            '4ヶ月未満の場合
                            fcCalendar = Color.Red
                        Else
                            '4ヶ月以上の場合
                            fcCalendar = Color.Black
                        End If
                    Else
                        fcCalendar = Color.Black
                    End If
                    ' 2015.12.03 UPD END↑ h.hagiwara 仮予約のみ赤文字で表示する

                    '---セルへ予約文字列・背景色・文字色を設定
                    If intDay <= 15 Then
                        If .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).Value = String.Empty Then
                            .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).Value = strTheaterInfo
                            .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).BackColor = bcCalendar
                            .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).ForeColor = fcCalendar
                            intYoyakuRow = 1
                        Else
                            .PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).Value = strTheaterInfo
                            .PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).BackColor = bcCalendar
                            .PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).ForeColor = fcCalendar
                            intYoyakuRow = 2
                        End If
                    Else
                        If .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).Value = String.Empty Then
                            .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).Value = strTheaterInfo
                            .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).BackColor = bcCalendar
                            .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).ForeColor = fcCalendar
                            intYoyakuRow = 1
                        Else
                            .PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).Value = strTheaterInfo
                            .PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).BackColor = bcCalendar
                            .PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).ForeColor = fcCalendar
                            intYoyakuRow = 2
                        End If
                    End If

                    '---日・行・予約番号・ステータスをArrayListに保持
                    .PropDicYoyakuInfo.Add(intDay.ToString & "_" & intYoyakuRow.ToString & "_" & .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_STS").ToString, _
                                           .PropDtYoyakuList.Rows(intDataRow)("YOYAKU_NO").ToString)
                Next

                'キャンセル待ち情報表示
                .PropDicCancelInfo = New Dictionary(Of String, String)
                For intDataRow As Integer = 0 To .PropDtDateKakuOnlyCancelWaitList.Rows.Count - 1
                    If .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow).IsNull("RIYO_DT") = False Then
                        intDay = Integer.Parse(.PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("RIYO_DT").ToString.Substring(8, 2))
                        If .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("MITEI_FLG").ToString = "0" Then
                            strCancelInfo = .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("START_TIME").ToString.Substring(0, 2) & ":" & _
                                             .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("START_TIME").ToString.Substring(2, 2) & _
                                             "-" & .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("END_TIME").ToString.Substring(0, 2) & ":" & _
                                             .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("END_TIME").ToString.Substring(2, 2) & vbCrLf & _
                                             .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("SAIJI_NM").ToString & vbCrLf & _
                                             .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("RIYO_NM").ToString
                        Else
                            strCancelInfo = .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("SAIJI_NM").ToString & vbCrLf & _
                                             .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("RIYO_NM").ToString
                        End If
                        intCancelRow = Integer.Parse(.PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("WAKU_NO"))
                        If intDay <= 15 Then
                            .PropvwCalandarFirst.ActiveSheet.Cells(intCancelRow + 2, intDay - 1).Value = strCancelInfo
                        Else
                            .PropvwCalandarSecond.ActiveSheet.Cells(intCancelRow + 2, intDay - 16).Value = strCancelInfo
                        End If

                        '---日・行・予約番号をArrayListに保持
                        .PropDicCancelInfo.Add(intDay.ToString & "_" & intCancelRow.ToString, _
                                               .PropDtDateKakuOnlyCancelWaitList.Rows(intDataRow)("CANCEL_WAIT_NO").ToString)
                    End If
                Next

                'その他情報表示
                For intDataRow As Integer = 0 To .PropDtEtcUseYoyakuList.Rows.Count - 1
                    intDay = Integer.Parse(.PropDtEtcUseYoyakuList.Rows(intDataRow)("RIYO_DT").ToString.Substring(8, 2))
                    strEtcInfo = .PropDtEtcUseYoyakuList.Rows(intDataRow)("START_TIME").ToString.Substring(0, 2) & ":" & _
                                     .PropDtEtcUseYoyakuList.Rows(intDataRow)("START_TIME").ToString.Substring(2, 2) & _
                                     "-" & .PropDtEtcUseYoyakuList.Rows(intDataRow)("END_TIME").ToString.Substring(0, 2) & ":" & _
                                     .PropDtEtcUseYoyakuList.Rows(intDataRow)("END_TIME").ToString.Substring(2, 2) & vbCrLf & _
                                     .PropDtEtcUseYoyakuList.Rows(intDataRow)("RIYO_BASHO").ToString & vbCrLf & _
                                     .PropDtEtcUseYoyakuList.Rows(intDataRow)("RIYO_YOTO").ToString & vbCrLf & _
                                     .PropDtEtcUseYoyakuList.Rows(intDataRow)("RIYOSHA").ToString
                    If intDay <= 15 Then
                        .PropvwCalandarFirst.ActiveSheet.Cells(6, intDay - 1).Value = strEtcInfo
                    Else
                        .PropvwCalandarSecond.ActiveSheet.Cells(6, intDay - 16).Value = strEtcInfo
                    End If
                Next

                '休館日、メンテ日の表示
                For intDataRow As Integer = 0 To .PropDtKyuykanMaiteList.Rows.Count - 1
                    intDay = Integer.Parse(.PropDtKyuykanMaiteList.Rows(intDataRow)("HOLMENT_DT").ToString.Substring(8, 2))
                    strKyukanMainteInfo = .PropDtKyuykanMaiteList.Rows(intDataRow)("MNAIYO").ToString
                    If .PropDtKyuykanMaiteList.Rows(intDataRow)("HOLMENT_KBN").ToString = "1" Then
                        '休館日の場合
                        bcCalendar = Color.FromArgb(224, 224, 224)
                    Else
                        'メンテ日の場合
                        bcCalendar = Color.Gray
                    End If
                    fcCalendar = Color.Black
                    If intDay <= 15 Then
                        .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).Value = strKyukanMainteInfo
                        .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).BackColor = bcCalendar
                        .PropvwCalandarFirst.ActiveSheet.Cells(1, intDay - 1).ForeColor = fcCalendar
                        .PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).BackColor = bcCalendar
                        .PropvwCalandarFirst.ActiveSheet.Cells(2, intDay - 1).ForeColor = fcCalendar
                    Else
                        .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).Value = strKyukanMainteInfo
                        .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).BackColor = bcCalendar
                        .PropvwCalandarSecond.ActiveSheet.Cells(1, intDay - 16).ForeColor = fcCalendar
                        .PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).BackColor = bcCalendar
                        .PropvwCalandarSecond.ActiveSheet.Cells(2, intDay - 16).ForeColor = fcCalendar
                    End If
                Next

                '■キャンセル待ち一覧表示
                intlastRow = .PropvwCancelWait.ActiveSheet.RowCount
                For intDataRow As Integer = 0 To .PropDtAllCancelWaitList.Rows.Count - 1

                    '最終行に1行追加
                    .PropvwCancelWait.Sheets(0).AddUnboundRows(intlastRow, 1)

                    .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 0).Value = .PropDtAllCancelWaitList.Rows(intDataRow)("SAIJI_NM").ToString
                    .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 1).Value = .PropDtAllCancelWaitList.Rows(intDataRow)("RIYO_NM").ToString
                    If .PropDtAllCancelWaitList.Rows(intDataRow)("RIYO_DT_FLG").ToString = "1" Then
                        Dim dt As Date = .PropDtAllCancelWaitList.Rows(intDataRow)("RIYO_DT").ToString
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 2).Value = dt.ToString("MM/dd")
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 3).Value = "○"
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 5).Locked = True
                    Else
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 2).Value = .PropDtAllCancelWaitList.Rows(intDataRow)("RIYO_MEMO").ToString
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 3).Value = String.Empty
                        .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 5).Locked = False
                    End If
                    .PropvwCancelWait.ActiveSheet.Cells(intlastRow, 6).Value = .PropDtAllCancelWaitList.Rows(intDataRow)("CANCEL_WAIT_NO").ToString

                    intlastRow += 1
                Next

                '■キャンセル済一覧表示
                intlastRow = .PropvwCancelDone.ActiveSheet.RowCount
                For intDataRow As Integer = 0 To .PropDtCancelDoneList.Rows.Count - 1

                    '最終行に1行追加
                    .PropvwCancelDone.Sheets(0).AddUnboundRows(intlastRow, 1)

                    .PropvwCancelDone.ActiveSheet.Cells(intlastRow, 0).Value = .PropDtCancelDoneList.Rows(intDataRow)("SAIJI_NM").ToString
                    .PropvwCancelDone.ActiveSheet.Cells(intlastRow, 1).Value = .PropDtCancelDoneList.Rows(intDataRow)("RIYO_NM").ToString
                    If .PropDtCancelDoneList.Rows(intDataRow)("RIYO_DT_FLG").ToString = "1" Then
                        Dim dt As Date = .PropDtCancelDoneList.Rows(intDataRow)("RIYO_DT").ToString
                        .PropvwCancelDone.ActiveSheet.Cells(intlastRow, 2).Value = dt.ToString("MM/dd")
                    Else
                        .PropvwCancelDone.ActiveSheet.Cells(intlastRow, 2).Value = .PropDtCancelDoneList.Rows(intDataRow)("RIYO_MEMO").ToString
                    End If
                    .PropvwCancelDone.ActiveSheet.Cells(intlastRow, 4).Value = .PropDtCancelDoneList.Rows(intDataRow)("CANCEL_WAIT_NO").ToString

                    intlastRow += 1
                Next

            End With


            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            '例外処理
            puErrMsg = ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "予約制御表の登録"

    ''' <summary>
    ''' 予約受付制御登録
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約制御情報の登録
    ''' <para>作成情報：2015.11.19 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetInsertUketukeSeigyo(ByRef dataEXTB0101 As DataEXTB0101, ByVal aryRiyobiList As ArrayList) As Boolean

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            For i = 0 To aryRiyobiList.Count - 1
                sqlEXTB0101.insertYoyakuCtl(Cmd, Cn, aryRiyobiList(i))
                Cmd.ExecuteNonQuery()
            Next

            'エラーがなければコミットする
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

#End Region

#Region "キャンセル予約制御表の登録"

    ''' <summary>
    ''' キャンセル予約受付制御登録
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>キャンセル予約制御情報の登録
    ''' <para>作成情報：2015.11.19 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetInsertCancelUketukeSeigyo(ByRef dataEXTB0101 As DataEXTB0101, ByVal aryRiyobiList As ArrayList) As Boolean

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable()
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            '登録
            Tsx = Cn.BeginTransaction
            For i = 0 To aryRiyobiList.Count - 1
                sqlEXTB0101.insertCancelYoyakuCtl(Cmd, Cn, aryRiyobiList(i))
                Cmd.ExecuteNonQuery()
            Next

            'エラーがなければコミットする
            commonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "COMMIT!!", Nothing, Nothing)
            Tsx.Commit()

            Cn.Close()

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ロールバック
            Tsx.Rollback()
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

#End Region

End Class
