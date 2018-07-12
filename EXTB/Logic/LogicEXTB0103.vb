Imports Common
Imports CommonEXT
Imports Npgsql
Imports EXTZ
Imports System
Imports FarPoint.Win.Spread.Model
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports System.Text

Public Class LogicEXTB0103

    '変数宣言
    Private sqlEXTB0102 As New SqlEXTB0102          'sqlクラス
    Private sqlEXTB0103 As New SqlEXTB0103          'sqlクラス

    'Private定数宣言
    '利用承認書
    '出力文字列：利用内容
    Private Const EVENT_INFO As String = "{0}利用"
    '出力文字列：利用形状
    Private Const USETYPE_STANDING As String = "スタンディング"
    Private Const USETYPE_SEATING As String = "着席"
    Private Const USETYPE_ANOMALY As String = "変則"
    Private Const USETYPE_EVENT As String = "催事"
    '出力文字列：ドリンク代
    Private Const DRINK_IN As String = "主催者負担"
    Private Const DRINK_OUT As String = "お客様負担"
    Private Const DRINK_NOTHING As String = "無"
    '表示文字列：利用日数
    Private Const USE_DAYS As String = "{0}日間"
    '判定文字列：利用料金
    Private Const FEE As String = "利用料"
    '表示文字列：小計
    Private Const SUBTOTAL_FORMULA As String = "SUM(I18:I21)"
    '表示文字列：消費税
    Private Const TAX_TEXT As String = "消費税{0}%"
    Private Const TAX_FORMULA As String = "I22/100*{0}"
    '表示文字列：合計
    Private Const TOTAL_FORMULA As String = "I22+I23"
    '出力ファイル名：利用承認書
    Private Const CERTIFICATES_FILENAME As String = "利用承認書_{0}.xlsx"
    '付帯設備利用明細書
    '表示文字列：基本会場利用料
    Private Const THEATER_FEE As String = "基本会場利用料"
    '表示文字列：小計（項目別）
    Private Const INCIDENT_SUBFEE As String = "F{0}+G{0}"
    '出力ファイル名：利用明細書（合算）
    Private Const DETAILS_FILENAME_ALL As String = "付帯設備利用明細書_{0}.xlsx"
    '出力ファイル名：利用明細書（日別）
    Private Const DETAILS_FILENAME_ONEDAY As String = "付帯設備利用明細書_{0}_{1}.xlsx"

    '帳票フォーマット名
    Private Const FORMAT_CERTIFICATES As String = "利用承認書_シアター.xlsx"                 ' 利用承認書
    Private Const FORMAT_DETAILS As String = "付帯設備利用明細書_シアター.xlsx"              ' 利用明細書

    Private Property vwAppSheetNote As Object

    ''' <summary>
    ''' 各種明細情報登録／更新
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報登録
    ''' <para>作成情報：2015/09/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegYoyakuDetail(ByRef dataEXTB0102 As DataEXTB0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'トランザクション開始
            Tsx = Cn.BeginTransaction
            '予約情報
            sqlEXTB0103.registerYoyakuInfo(Cmd, Cn, dataEXTB0102)
            '明細更新
            '付帯
            Dim rowCnt As Integer = 0
            Dim row As DataRow
            Dim ht As Hashtable = dataEXTB0102.PropHtFutai
            '削除された曜日データの削除
            sqlEXTB0103.deleteFutai(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropListRiyobi)
            '登録・更新
            For Each Key As String In ht.Keys
                Table = ht(Key)
                row = Table.Rows(0)
                sqlEXTB0103.registerFutai(Cmd, Cn, row, dataEXTB0102.PropStrYoyakuNo)
            Next
            '付帯明細
            sqlEXTB0103.deleteFutaiDetail(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo)
            ht = dataEXTB0102.PropHtFutaiDetail
            For Each Key As String In ht.Keys
                rowCnt = 0
                Table = ht(Key)
                Do While Table.Rows.Count > rowCnt
                    row = Table.Rows(rowCnt)
                    sqlEXTB0103.registerFutaiDetail(Cmd, Cn, row, dataEXTB0102.PropStrYoyakuNo)
                    rowCnt = rowCnt + 1
                Loop
            Next
            '請求/入金/EXAS入金(billpay_tbl)
            sqlEXTB0103.deleteBillPay(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo)
            rowCnt = 0
            Table = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim innerTable As New DataTable
            Dim innerRow As DataRow
            Dim innerRowCnt As Integer = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                '請求依頼番号の発番
                If IsDBNull(row("seikyu_irai_no")) And row("seikyu_input_flg") = "1" Then
                    row("seikyu_irai_no") = sqlEXTB0103.nextValSeikyuIraiNo(Cmd, Cn)
                End If
                '請求依頼更新
                sqlEXTB0103.registerBillPay(Cmd, Cn, row, dataEXTB0102.PropStrYoyakuNo)
                'EXASプロジェクト設定(project_tbl)
                'EXASプロジェクト設定付帯設備(project_tbl)
                'If row("seikyu_input_flg") = "1" Then                                                ' 2016.02.03 UPD h.hagiwara
                If row("seikyu_input_flg") = "1" And row("get_seikyu_input_flg") <> "1" Then          ' 2016.02.03 UPD h.hagiwara
                    Dim seikyuKey As String = row("seikyu_irai_no")
                    If SEIKYU_NAIYOU_RIYO = row("seikyu_naiyo") Or SEIKYU_NAIYOU_RIYOFUTAI = row("seikyu_naiyo") Then
                        If dataEXTB0102.PropHtExasPro.ContainsKey(seikyuKey) = True Then
                            innerTable = dataEXTB0102.PropHtExasPro(seikyuKey)
                        Else
                            innerTable = dataEXTB0102.PropHtExasPro(rowCnt.ToString)
                        End If
                        innerRowCnt = 0
                        sqlEXTB0103.deleteExasProject(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo, row("seikyu_irai_no"), innerTable)          ' 2016.01.15 DEL h.hagiwara
                        Do While innerTable.Rows.Count > innerRowCnt
                            innerRow = innerTable.Rows(innerRowCnt)
                            innerRow("seikyu_irai_no") = seikyuKey
                            innerRow("seikyu_irai_seq") = row("seq")
                            sqlEXTB0103.registerExasProject(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo, innerRow)
                            innerRowCnt = innerRowCnt + 1
                        Loop
                    End If
                    If SEIKYU_NAIYOU_FUTAI = row("seikyu_naiyo") Or SEIKYU_NAIYOU_RIYOFUTAI = row("seikyu_naiyo") Then
                        If dataEXTB0102.PropHtExasProFutai.ContainsKey(seikyuKey) = True Then
                            innerTable = dataEXTB0102.PropHtExasProFutai(seikyuKey)
                        Else
                            innerTable = dataEXTB0102.PropHtExasProFutai(rowCnt.ToString)
                        End If
                        innerRowCnt = 0
                        If SEIKYU_NAIYOU_FUTAI = row("seikyu_naiyo") Then                       ' 2016.01.15 ADD  h.hagiwara
                            sqlEXTB0103.deleteExasProject(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo, row("seikyu_irai_no"), innerTable)
                        End If                                                                  ' 2016.01.15 ADD  h.hagiwara
                        Do While innerTable.Rows.Count > innerRowCnt
                            innerRow = innerTable.Rows(innerRowCnt)
                            innerRow("seikyu_irai_no") = seikyuKey
                            innerRow("seikyu_irai_seq") = row("seq")
                            sqlEXTB0103.registerExasProject(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo, innerRow)
                            innerRowCnt = innerRowCnt + 1
                        Loop
                    End If
                End If
                rowCnt = rowCnt + 1
            Loop
            'タイムスケジュール
            sqlEXTB0103.deleteTimeSchedule(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo)           ' タイムスケジュールの削除  2015.12.03 ADD h.hagiwara
            Table = dataEXTB0102.PropDsTimeSch.Tables("TIME_SCHEDULE_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTB0102.PropStrYoyakuNo
                sqlEXTB0103.registerSchedule(Cmd, Cn, row, rowCnt)
                rowCnt = rowCnt + 1
            Loop
            'Ticket
            sqlEXTB0103.deleteTicket(Cmd, Cn, dataEXTB0102.PropStrYoyakuNo)           ' タイムスケジュールの削除  2015.12.03 ADD h.hagiwara
            Table = dataEXTB0102.PropDsTicket.Tables("TICKET_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTB0102.PropStrYoyakuNo
                sqlEXTB0103.registerTicket(Cmd, Cn, row, rowCnt)
                rowCnt = rowCnt + 1
            Loop

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Complate SQL", Nothing, Cmd)

            'トランザクションCOMMIT
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
            Throw ex
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 請求取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBillReq(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectBillReq(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "請求情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 入金情報
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNyukin(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            Dim dtBillPay As DataTable = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow

            '2016.09.21 m.hayabuchi add start 不具合対応（入金取消後、入金リンクデータ枠のデータが残る）
            'PropDtExasNyukinを初期化 
            If Not dataEXTB0102.PropDtExasNyukin Is Nothing Then
                dataEXTB0102.PropDtExasNyukin.Dispose()
            End If
            dataEXTB0102.PropDtExasNyukin = Nothing
            '2016.09.21 m.hayabuchi add end 不具合対応

            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                If row("nyukin_link_no") Is DBNull.Value = False Then
                    sqlEXTB0103.selectNyukin(Adapter, Cn, dataEXTB0102, row("nyukin_link_no"), "", index)
                ElseIf row("seikyu_irai_no") Is DBNull.Value = False Then
                    sqlEXTB0103.selectNyukin(Adapter, Cn, dataEXTB0102, "", row("seikyu_irai_no"), index)
                Else
                    sqlEXTB0103.selectNyukin(Adapter, Cn, dataEXTB0102, "", "", index)
                End If
                index = index + 1
            Loop
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "請求情報取得", Nothing, Adapter.SelectCommand)
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
    ''' チケット情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTicketInfo(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectTicket(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "チケット情報取得", Nothing, Adapter.SelectCommand)
            ' Dim tst As String = sqlEXTB0103.nextValSeikyuIraiNo(Adapter.SelectCommand, Cn)                          ' 2015.01.25 DEL h.hagiwara
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
    ''' 現金電子マネー情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCashAndElectroMoney(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectCashElectro(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "現金電子マネー情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 確認依頼確認記録情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetApproval(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectApprovalReq(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "確認依頼情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTB0103.selectApprovalRes(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "確認記録情報取得", Nothing, Adapter.SelectCommand)
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
    ''' タイムスケジュール情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTimeSchedule(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectTimeSch(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "タイムスケジュール情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 付帯情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFutai(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectFutaiHeadder(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTB0103.selectFutaiDetail(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 付帯情報取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpFutai(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectFutaiHeadderUp(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTB0103.selectFutaiDetailUp(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 付帯情報削除対象チェック
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DelFutai(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
        Dim htHeader As Hashtable = dataEXTB0102.PropHtFutai
        Dim htDetail As Hashtable = dataEXTB0102.PropHtFutaiDetail
        Dim exist As Boolean = False
        For Each Key As String In htHeader.Keys
            For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
                If Key = dataRiyobi.PropStrYoyakuDt = True Then
                    exist = True
                End If
            Next
            If exist = False Then
                htHeader.Remove(Key)
                htDetail.Remove(Key)
            End If
            exist = False
        Next
        Return True
    End Function

    ''' <summary>
    ''' EXASプロジェクト情報
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExasRiyoryo(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectExasRiyoryo(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            If sqlEXTB0103.selectExasFutai(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXASプロジェクト情報取得", Nothing, Adapter.SelectCommand)
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
    ''' EXASプロジェクト情報
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CalcExasRiyoryo(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
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
            If sqlEXTB0103.selectExasRiyoryoUp(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            If sqlEXTB0103.selectExasFutaiUp(Adapter, Cn, dataEXTB0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "EXASプロジェクト情報取得", Nothing, Adapter.SelectCommand)
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
    ''' 履歴記録登録・メール送信
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報登録
    ''' <para>作成情報：2015/09/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegRirekiKiroku(ByRef dataEXTB0102 As DataEXTB0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'トランザクション開始
            Tsx = Cn.BeginTransaction
            Dim rowCnt As Integer = 0
            Dim row As DataRow
            '承認依頼／記録
            Table = dataEXTB0102.PropDsApproval.Tables("IRAI_RIREKI_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTB0102.PropStrYoyakuNo
                If IsDBNull(row("seq")) = True Then
                    'Dim body As String = String.Format(CommonDeclareEXTB.MAIL_SHONIN_IRAI_BODY, dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropStrShutsuenNm, dataEXTB0102.PropStrRiyoNm, GetRiyobi(dataEXTB0102, True), GetRiyobi(dataEXTB0102, False))    ' 2015.12.15 UPD h.hagiwara
                    'sendMail(CommonDeclareEXTB.MAIL_SHONIN_IRAI_SUBJECT, body, False)    ' 2015.11.20 UPD h.hagiwara
                    Dim body As String = String.Format(CommonDeclareEXTB.MAIL_SHONIN_IRAI_BODY, dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropStrSaijiNm, dataEXTB0102.PropStrRiyoNm, GetRiyobi(dataEXTB0102, True), GetRiyobi(dataEXTB0102, False))        ' 2015.12.15 UPD h.hagiwara
                    sendMail(CommonDeclareEXTB.MAIL_SHONIN_IRAI_SUBJECT, body, True, dataEXTB0102.PropStrYoyakuNo)      ' 2015.11.20 UPD h.hagiwara
                    sqlEXTB0103.registerIraiRireki(Cmd, Cn, row)
                End If
                rowCnt = rowCnt + 1
            Loop
            Table = dataEXTB0102.PropDsApproval.Tables("CHECK_RIREKI_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTB0102.PropStrYoyakuNo
                If IsDBNull(row("seq")) = True Then
                    'Dim body As String = String.Format(CommonDeclareEXTB.MAIL_SHONIN_KAKUNIN_BODY, dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropStrShutsuenNm, dataEXTB0102.PropStrRiyoNm, GetRiyobi(dataEXTB0102, True), GetRiyobi(dataEXTB0102, False))    ' 2015.12.15 UPD h.hagiwara
                    'sendMail(CommonDeclareEXTB.MAIL_SHONIN_KAKUNIN_SUBJECT, body, True)     ' 2015.11.20 UPD h.hagiwara
                    Dim body As String = String.Format(CommonDeclareEXTB.MAIL_SHONIN_KAKUNIN_BODY, dataEXTB0102.PropStrYoyakuNo, dataEXTB0102.PropStrSaijiNm, dataEXTB0102.PropStrRiyoNm, GetRiyobi(dataEXTB0102, True), GetRiyobi(dataEXTB0102, False))        ' 2015.12.15 UPD h.hagiwara
                    sendMail(CommonDeclareEXTB.MAIL_SHONIN_KAKUNIN_SUBJECT, body, False, dataEXTB0102.PropStrYoyakuNo)     ' 2015.11.20 UPD h.hagiwara
                    sqlEXTB0103.registerCheckRireki(Cmd, Cn, row)
                End If
                rowCnt = rowCnt + 1
            Loop

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Complate SQL", Nothing, Cmd)

            'トランザクションCOMMIT
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
            Throw ex
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 日付取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <remarks></remarks>
    Public Function GetRiyobi(ByRef dataEXTB0102 As DataEXTB0102, ByVal isFrom As Boolean) As String
        Dim minDate As DateTime
        Dim maxDate As DateTime
        Dim isFirst As Boolean = True
        For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
            Dim dt As DateTime = dataRiyobi.PropStrYoyakuDt
            If isFirst Then
                minDate = dt
                maxDate = dt
                isFirst = False
            Else
                If DateTime.Compare(minDate, dt) > 0 Then
                    minDate = dt
                End If
                If DateTime.Compare(maxDate, dt) < 0 Then
                    maxDate = dt
                End If
            End If
        Next
        If isFrom Then
            Return minDate.ToLongDateString
        Else
            Return maxDate.ToLongDateString
        End If
    End Function

    ''' <summary>
    ''' メール送信
    ''' </summary>
    ''' <param name="StrSubject"></param>
    ''' <param name="StrBody"></param>
    ''' <param name="IsShonin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function sendMail(ByVal StrSubject As String, ByVal StrBody As String, ByVal IsShonin As Boolean, ByVal Yoyakuno As String) As Boolean
        Dim strMailTo As String = SetSendMailAddr(IsShonin, Yoyakuno)
        Dim strMailCc As String = CommonDeclareEXT.PropStrComMailAddr
        Dim strMailBcc As String = ""
        'If CommonLogic.SendMail(strMailTo, strMailCc, strMailBcc, StrSubject, StrBody, "", "") = False Then
        If Mailsend(strMailTo, strMailCc, strMailBcc, StrSubject, StrBody, "", "") = False Then
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "メールの送信に失敗しました。", Nothing, Nothing)
        End If
        Return True
    End Function

    ''' <summary>
    ''' 利用開始日(0),終了日(1)をArrayで取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Function setStartEndRiyobi(ByRef dataList As ArrayList) As Array
        Dim dtStart As Date
        Dim dtEnd As Date
        Dim dtTemp As Date
        Dim index As New Integer
        index = 0
        For Each dataRiyobi As CommonDataRiyobi In dataList
            If dtStart = Nothing Then
                dtStart = dataRiyobi.PropStrYoyakuDt
                dtEnd = dataRiyobi.PropStrYoyakuDt
            Else
                dtTemp = dataRiyobi.PropStrYoyakuDt
                If Date.Compare(dtTemp, dtStart) < 0 Then
                    dtStart = dataRiyobi.PropStrYoyakuDt
                End If
                If Date.Compare(dtTemp, dtEnd) > 0 Then
                    dtEnd = dataRiyobi.PropStrYoyakuDt
                End If
            End If
        Next
        Return {dtStart.ToString("yyyy/MM/dd"), dtEnd.ToString("yyyy/MM/dd")}
    End Function

    ''' <summary>
    ''' 送付先アドレス取得
    ''' </summary>
    ''' <returns>String アドレス</returns>
    ''' <remarks>メールアドレス
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetSendMailAddr(ByVal IsShonin As Boolean, ByVal Yoyakuno As String) As String

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
            If sqlEXTB0103.selecMailTo(Adapter, Cn, IsShonin, Yoyakuno) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "送付先メールアドレス取得", Nothing, Adapter.SelectCommand)

            Adapter.Fill(Table)
            If Table.Rows.Count = 0 Then
                'Falseを返す
                Return False
            End If

            Dim strAddr As String = ""
            Dim i As Integer
            For i = 0 To Table.Rows.Count - 1
                strAddr = strAddr + Table.Rows(i)(0)
                If i <> (Table.Rows.Count - 1) Then
                    strAddr = strAddr + ","
                End If
            Next i

            Cn.Close()

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return strAddr
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
    ''' 初期表示用の請求額(利用料金)を取得する
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRiyoKinTotal(ByVal dataEXTB0102 As DataEXTB0102) As Long
        Dim riyoKin As Long = 0
        For Each riyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
            riyoKin = riyoKin + riyobi.PropIntRiyoKin
        Next
        Return riyoKin
    End Function
    'Public Function GetRiyoKinTotal(ByVal dataEXTB0102 As DataEXTB0102) As Integer
    '    Dim riyoKin As Integer = 0
    '    For Each riyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
    '        riyoKin = riyoKin + riyobi.PropIntRiyoKin
    '    Next
    '    Return riyoKin
    'End Function

    ' 2015.12.21 UPD START↓ h.hagiwara
    ' ''' <summary>
    ' ''' 初期表示用の請求額(付帯料金)を取得する
    ' ''' </summary>
    ' ''' <param name="dataEXTB0102"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function GetFutairyo(ByVal dataEXTB0102 As DataEXTB0102) As Integer
    '    Dim futaiKin As Integer = 0
    '    Dim table As DataTable
    '    Dim row As DataRow
    '    Dim ht As Hashtable = dataEXTB0102.PropHtFutai
    '    For Each Key As String In ht.Keys
    '        table = ht(Key)
    '        row = table.Rows(0)
    '        futaiKin = futaiKin + row("total_fuzoku_kin")
    '    Next
    '    Return futaiKin
    'End Function
    ''' <summary>
    ''' 初期表示用の請求額(付帯料金)を取得する
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFutairyo(ByVal dataEXTB0102 As DataEXTB0102) As Long
        Dim futaiKin As Long = 0
        Dim table As DataTable
        Dim row As DataRow
        Dim ht As Hashtable = dataEXTB0102.PropHtFutai
        For Each Key As String In ht.Keys
            table = ht(Key)
            row = table.Rows(0)
            futaiKin = futaiKin + row("total_fuzoku_kin")
        Next
        Return futaiKin
    End Function
    ' 2015.12.21 UPD END↑ h.hagiwara

    ' 2015.12.21 UPD START↓ h.hagiwara
    ' ''' <summary>
    ' ''' 初期表示用の請求額(付帯料金)を取得する
    ' ''' </summary>
    ' ''' <param name="dataEXTB0102"></param>
    ' ''' <returns></returns>
    ' ''' <para>作成情報：2015.11.11 h.hagiwara</para>
    ' ''' <remarks></remarks>
    'Public Function GetFutairyoTax(ByVal dataEXTB0102 As DataEXTB0102) As Integer
    '    Dim futaiKintax As Integer = 0
    '    Dim table As DataTable
    '    Dim row As DataRow
    '    Dim ht As Hashtable = dataEXTB0102.PropHtFutai
    '    If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
    '        Return 0
    '    End If
    '    For Each Key As String In ht.Keys
    '        table = ht(Key)
    '        row = table.Rows(0)
    '        futaiKintax = futaiKintax + row("tax_kin")
    '    Next
    '    Return futaiKintax
    'End Function
    ''' <summary>
    ''' 初期表示用の請求額(付帯料金)を取得する
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <para>作成情報：2015.11.11 h.hagiwara</para>
    ''' <remarks></remarks>
    Public Function GetFutairyoTax(ByVal dataEXTB0102 As DataEXTB0102) As Long
        Dim futaiKintax As Long = 0
        Dim table As DataTable
        Dim row As DataRow
        Dim ht As Hashtable = dataEXTB0102.PropHtFutai
        If dataEXTB0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
            Return 0
        End If
        For Each Key As String In ht.Keys
            table = ht(Key)
            row = table.Rows(0)
            futaiKintax = futaiKintax + row("tax_kin")
        Next
        Return futaiKintax
    End Function
    ' 2015.12.21 UPD END↑ h.hagiwara

    ''' <summary>
    ''' 利用日(初日)の消費税を取得
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTaxfirstRiyobi(ByVal dataEXTB0102 As DataEXTB0102) As Double
        Dim tax As Double = 0
        If dataEXTB0102.PropStrKashiKind = KASHI_KIND_HOUSE Then
            Return 0
        End If
        For Each riyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
            tax = GetTax(riyobi.PropStrYoyakuDt)
            Return tax
        Next
        Return 0
    End Function

    ''' <summary>
    ''' 消費税取得処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>SYSTEM PROPERTY
    ''' <para>作成情報：2015/08/28 k.machida
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
            If sqlEXTB0103.selectTax(Adapter, Cn, Riyobi) = False Then
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
    ''' タイトル1を作成し、設定する
    ''' </summary>
    ''' <param name="dataEXTB0102"></param>
    ''' <param name="title1"></param>
    ''' <param name="seikyuNaiyo"></param>
    ''' <remarks></remarks>
    Public Sub SetSeikyuTitle(ByRef dataEXTB0102 As DataEXTB0102, _
                              ByRef title1 As String, _
                              ByVal seikyuNaiyo As String)
        Dim minDate As DateTime
        Dim maxDate As DateTime
        Dim isFirst As Boolean = True
        For Each dataRiyobi As CommonDataRiyobi In dataEXTB0102.PropListRiyobi
            Dim dt As DateTime = dataRiyobi.PropStrYoyakuDt
            If isFirst Then
                minDate = dt
                maxDate = dt
                isFirst = False
            Else
                If DateTime.Compare(minDate, dt) > 0 Then
                    minDate = dt
                End If
                If DateTime.Compare(maxDate, dt) < 0 Then
                    maxDate = dt
                End If
            End If
        Next
        Dim footer As String = ""
        If SEIKYU_NAIYOU_RIYO = seikyuNaiyo Then
            footer = "　利用料"
        ElseIf SEIKYU_NAIYOU_FUTAI = seikyuNaiyo Then
            footer = "　付帯設備"
        ElseIf SEIKYU_NAIYOU_RIYOFUTAI = seikyuNaiyo Then
            'footer = "　利用料＋付帯設備"
            footer = "　利用料＋付帯"
        End If
        'title1 = StrConv(minDate.ToLongDateString + "～" + maxDate.ToLongDateString + footer, VbStrConv.Wide)
        title1 = StrConv(minDate.ToLongDateString & "～" & maxDate.Month & "月" & maxDate.Day & "日" & footer, VbStrConv.Wide)
    End Sub

    ''' <summary>
    ''' 請求取消
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CancellSeikyu(ByRef row As DataRow) As Boolean
        '未登録データの場合処理しない
        If IsDBNull(row("seq")) Then
            Return True
        End If
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand
        Try
            Cn.Open()
            '登録
            sqlEXTB0103.updateCancellSeikyu(Cmd, Cn, row)
            Cmd.ExecuteNonQuery()
            'COMMIT
            Cn.Close()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 入金取消
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CancellNyukin(ByRef row As DataRow) As Boolean
        '未登録データの場合処理しない
        If IsDBNull(row("seq")) Then
            Return True
        End If
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand
        Try
            Cn.Open()
            '登録
            sqlEXTB0103.updateCancellNyukin(Cmd, Cn, row)
            Cmd.ExecuteNonQuery()
            'COMMIT
            Cn.Close()
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        Finally
            '終了処理
            Cn.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 利用承認書印刷メイン処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>利用承認書の印刷を行う
    ''' <para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function PrintCertificatesMain(ByVal dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '利用承認書データを作成
        If GetUseApproval(dataEXTB0103) = False Then
            Return False
        End If

        'TODO  消費税
        If GetTax(dataEXTB0103) = False Then
            Return False
        End If

        '利用承認書を編集し、印刷
        If PrintCertificates(dataEXTB0103) = False Then
            Return False
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常終了
        Return True

    End Function

    ''' <summary>
    ''' 付帯設備利用明細書（合算）出力メイン処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書（合算）の出力を行う
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function PrintUseDetailsMain(ByVal dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '合算フラグ設定
        If dataEXTB0103.PropStrClickedIncident IsNot Nothing And _
            dataEXTB0103.PropStrClickedIncident IsNot "" Then
            'ボタンクリック行の日付がある場合、False（日別）を設定
            dataEXTB0103.PropBlnSumFlg = False
        Else
            '日付がない場合はTrue（合算）を設定
            dataEXTB0103.PropBlnSumFlg = True
        End If

        '付帯設備利用証明書（合算）データ作成
        If MakeUseDetailsData(dataEXTB0103) = False Then
            Return False
        End If

        '付帯設備利用証明書（合算）を編集して出力
        If OutputUseDetails(dataEXTB0103) = False Then
            Return False
        End If

        If FildClear(dataEXTB0103) = False Then
            Return False
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常終了
        Return True

    End Function

    ''' <summary>
    ''' 付帯設備利用明細データ作成処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>画面情報から付帯設備利用明細（合算）のデータを作成する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function MakeUseDetailsData(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '変数宣言
        ''利用日時スプレッドシート
        'Dim vwUseDaysSheet As FarPoint.Win.Spread.SheetView
        ''請求／入金スプレッドシート
        'Dim vwClaimSheet As FarPoint.Win.Spread.SheetView
        ''スプレッドシート全体の行数
        'Dim intSpreadRows As Integer = 0
        ''スプレッドシートのデータ行数
        'Dim intCountRows As Integer = 0
        ''利用開始日
        'Dim strUseStartDay As String = ""
        ''利用終了日
        'Dim strUseEndDay As String = ""
        ''利用料合計
        'Dim lngSumFee As Long = 0
        ''調整額合計
        'Dim lngSumAdjustFee As Long = 0

        Dim cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ
        Dim table As New DataTable                     'データテーブル





        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            ' With dataEXTB0103

            '利用日時スプレッドシートを設定
            '    vwUseDaysSheet = .PropVwUseDays.ActiveSheet
            '請求／入金スプレッドシートを設定
            '   vwClaimSheet = .PropVwClaimPayment.ActiveSheet

            '貸出種別（一般or社内）
            'For Each rdoRentalClass As RadioButton In .PropPnlRentalClass.Controls
            '    If rdoRentalClass.Checked Then
            '        .PropStrRentalClass_Output = rdoRentalClass.Text
            '        Exit For
            '    End If
            'Next
            ''承認番号
            '.PropStrReserveNo_Output = .PropLblReserveNo.Text
            ''会社・団体名
            '.PropStrUserName_Output = .PropTxtUserName.Text
            ''催事名（催事名＋出演者名）
            '.PropStrEventName_Output = .PropTxtEventName.Text + " " + .PropTxtPerformerName.Text
            ''出力日
            '.PropStrOutputDay_Output = DateTime.Today.ToString("yyyy年M月d日")

            'If .PropBlnSumFlg Then
            '    '出力対象が合算の場合
            '    '利用日時スプレッドシートの行数を設定
            '    intSpreadRows = vwUseDaysSheet.RowCount
            '    '利用日時スプレッドシートのデータ行数を取得
            '    For i As Integer = 0 To intSpreadRows - 1
            '        If vwUseDaysSheet.Cells(i, 1).Value IsNot Nothing And _
            '             vwUseDaysSheet.Cells(i, 1).Value IsNot "" Then
            '            intCountRows = intCountRows + 1
            '        End If
            '    Next
            '    '利用日時
            '    If intCountRows = 1 Then
            '        '利用日時が１日の場合'yyyy年MM月dd日（ddd）'の表記にする
            '        .PropStrUseDays_Output = DateTime.Parse(vwUseDaysSheet.Cells(0, 2).Value).ToString("yyyy年M月d日")
            '    ElseIf intCountRows >= 2 Then
            '        '利用日時が２日以上の場合'yyyy年MM月dd日（ddd）～yyyy年MM月dd日（ddd）'の表記にする
            '        strUseStartDay = DateTime.Parse(vwUseDaysSheet.Cells(0, 2).Value).ToString("yyyy年M月d日")
            '        strUseEndDay = DateTime.Parse(vwUseDaysSheet.Cells(intCountRows - 1, 2).Value).ToString("yyyy年M月d日")
            '        .PropStrUseDays_Output = strUseStartDay + "～" + strUseEndDay
            '    End If
            '    '利用日数
            '    .PropStrCountUseDays_Output = intCountRows.ToString
            '    '請求／入金スプレッドシートの行数を設定
            '    intSpreadRows = vwClaimSheet.RowCount
            '    For i As Integer = 0 To intSpreadRows - 1
            '        If vwClaimSheet.Cells(i, 1).Value = "○" And _
            '            vwClaimSheet.Cells(i, 12).Value = FEE Then
            '            lngSumFee = lngSumFee + vwClaimSheet.Cells(i, 5).Value
            '            lngSumAdjustFee = lngSumAdjustFee + vwClaimSheet.Cells(i, 6).Value
            '        End If
            '    Next
            '利用料金の合計を設定
            '.PropLngTheaterFee_Output = lngSumFee
            '.PropLngAdjustFee_Output = lngSumAdjustFee
            '税率計算日
            '消費税変更日を跨ぐことはないため、利用開始日を取得する
            '  .PropStrCalculateDay_Output = vwUseDaysSheet.Cells(0, 2).Value

            'ご利用明細、立替経費等明細
            'DBから明細を取得
            'If GetUseDetailsAll(dataEXTB0103) = False Then
            '    Return False
            'End If

            '    Else
            '    '出力対象が付帯設備スプレッドシートの「印刷」ボタンの日付である場合
            '    '利用日時
            '    .PropStrUseDays_Output = DateTime.Parse(.PropStrClickedIncident).ToString("yyyy年M月d日")
            '    '利用日数
            '    .PropStrCountUseDays_Output = "1"
            '    '利用日時スプレッドシートのデータ行数を取得
            '    intSpreadRows = vwUseDaysSheet.RowCount
            '    For i As Integer = 0 To intSpreadRows - 1
            '        If vwUseDaysSheet.Cells(i, 1).Value IsNot Nothing And _
            '             vwUseDaysSheet.Cells(i, 1).Value IsNot "" Then
            '            intCountRows = intCountRows + 1
            '        End If
            '    Next
            '    '利用日時スプレッドシートより利用料を取得
            '    For i As Integer = 0 To intCountRows - 1
            '        If DateTime.Parse(vwUseDaysSheet.Cells(i, 2).Value) = DateTime.Parse(.PropStrClickedIncident) Then
            '            lngSumFee = lngSumFee + vwUseDaysSheet.Cells(i, 9).Value
            '        End If
            '    Next

            '    '利用料金の合計を設定
            '    .PropLngTheaterFee_Output = lngSumFee
            '    .PropLngAdjustFee_Output = lngSumAdjustFee
            '    '税率計算日
            '    '利用日時を設定する
            '    .PropStrCalculateDay_Output = .PropStrClickedIncident

            '    'ご利用明細、立替経費等明細
            '    'DBから明細を取得
            '    If GetUseDetailsOneDay(dataEXTB0103) = False Then
            '        Return False
            '    End If
            '    End If

            '    '消費税マスタより消費税を取得 TODO 取得方法変更
            '    If GetTax(dataEXTB0103) = False Then
            '        Return False
            '    End If

            'End With


            If dataEXTB0103.PropBlnSumFlg Then
                '付帯明細(合算)の場合
                dataEXTB0103.PropBlnNoTaxFlg = False
                '付帯利用明細取得SQLを設定
                If sqlEXTB0103.SetSelectUseDetailsAllSql(Adapter, cn, dataEXTB0103) = False Then
                    Return False
                End If

                'sql実行
                Adapter.Fill(table)
                'DBの値をDATAクラスに代入
                dataEXTB0103.PropDtUseDetailsNoTax_Output = table

                'TODO  消費税
                'dataEXTB0103.PropStrTax_Output = 10
                If GetTax(dataEXTB0103) = False Then
                    Return False
                End If

                ''アダプタとデータテーブルを解放
                'Adapter.Dispose()
                table.Dispose()

                '税なしのデータを取得
                table = New DataTable
                dataEXTB0103.PropBlnNoTaxFlg = True
                '付帯利用明細取得SQLを設定
                If sqlEXTB0103.SetSelectUseDetailsAllSql(Adapter, cn, dataEXTB0103) = False Then
                    Return False
                End If

                'sql実行
                Adapter.Fill(table)
                'DBの値をDATAクラスに代入
                dataEXTB0103.PropDtUseDetails_Output = table
            Else
                '付帯利用明細(日別)の場合

                '税ありのデータを取得
                dataEXTB0103.PropBlnNoTaxFlg = False
                '付帯利用明細取得SQLを設定
                If sqlEXTB0103.SetSelectUseDetailsOneDaySql(Adapter, cn, dataEXTB0103) = False Then
                    Return False
                End If

                'sql実行
                Adapter.Fill(table)
                'DBの値をDATAクラスに代入
                dataEXTB0103.PropDtUseDetailsNoTax_Output = table

                'TODO  消費税
                'dataEXTB0103.PropStrTax_Output = 10
                If GetTax(dataEXTB0103) = False Then
                    Return False
                End If

                ''アダプタとデータテーブルを解放
                'Adapter.Dispose()
                table.Dispose()


                '税なしのデータを取得
                table = New DataTable
                dataEXTB0103.PropBlnNoTaxFlg = True
                '付帯利用明細取得SQLを設定
                If sqlEXTB0103.SetSelectUseDetailsOneDaySql(Adapter, cn, dataEXTB0103) = False Then
                    Return False
                End If

                'sql実行
                Adapter.Fill(table)
                'DBの値をDATAクラスに代入
                dataEXTB0103.PropDtUseDetails_Output = table

            End If
            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            Adapter.Dispose()
            cn.Dispose()
            cn.Close()
            table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 利用承認書出力処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>利用承認書を編集し、印刷する
    ''' <para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function PrintCertificates(ByVal dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数
        Dim downloadPath As String      'ダウンロードパス

        Try
            With dataEXTB0103
                downloadPath = ConfigurationManager.AppSettings("formatPath") & FORMAT_CERTIFICATES
                'Excel読み込み
                .PropVwPrintSheet = New FarPoint.Win.Spread.FpSpread
                '本番環境でフォーマットをどこに置くか要確認
                '.PropVwPrintSheet.OpenExcel("..\..\Format\利用承認書_シアター.xlsx")
                .PropVwPrintSheet.OpenExcel(downloadPath)

                '通貨型セルの設定
                Dim currcell As New FarPoint.Win.Spread.CellType.CurrencyCellType()
                currcell.CurrencySymbol = "\"
                currcell.ShowSeparator = True

                '結合されている場合左上のセルを参照する
                ' シート設定
                Dim vwAppSheet As FarPoint.Win.Spread.SheetView = .PropVwPrintSheet.Sheets(0)      '利用承認書
                vwAppSheet.ColumnCount = 11 '列数
                vwAppSheet.RowCount = 31    '行数
                Dim vwAppSheetNote As FarPoint.Win.Spread.SheetView = .PropVwPrintSheet.Sheets(1)  '利用承認書（シアター控え）
                vwAppSheetNote.ColumnCount = 11 '列数
                vwAppSheetNote.RowCount = 31    '行数

                '付帯設備利用料
                If .PropDtUseApproval_Output IsNot Nothing Then
                    'DataTableから一行取得
                    Dim dtSelectRow As DataRow = .PropDtUseApproval_Output.Rows(0)

                    '利用承認書に値を設定
                    vwAppSheet.Cells("I5").Value = dtSelectRow.Item(0)   '承認番号
                    vwAppSheet.Cells("C8").Value = dtSelectRow.Item(1)   '会社・団体名
                    vwAppSheet.Cells("C9").Value = dtSelectRow.Item(2)   '代表者名
                    vwAppSheet.Cells("C10").Value = dtSelectRow.Item(3)  '利用責任者
                    vwAppSheet.Cells("C12").Value = dtSelectRow.Item(4)  '利用内容
                    vwAppSheet.Cells("C13").Value = dtSelectRow.Item(5)  '催事名
                    vwAppSheet.Cells("C14").Value = dtSelectRow.Item(6)  '利用形状
                    vwAppSheet.Cells("J14").Value = dtSelectRow.Item(7)  '定員
                    vwAppSheet.Cells("C15").Value = dtSelectRow.Item(8)  'ドリンク代徴収
                    vwAppSheet.Cells("A18").Value = dtSelectRow.Item(9)  '利用日時
                    vwAppSheet.Cells("G18").Value = dtSelectRow.Item(10) '利用日数
                    vwAppSheet.Cells("I18").Value = dtSelectRow.Item(11) '利用料金
                    vwAppSheet.Cells("I18").CellType = currcell
                    vwAppSheet.Cells("I22").CellType = currcell

                    '消費税セルタイプ
                    '貸出種別が一般利用／特例の場合￥を表示
                    '貸出種別が社内利用の場合は非表示
                    If dtSelectRow.Item(12) = 1 Or dtSelectRow.Item(12) = 3 Then
                        vwAppSheet.Cells("I23").CellType = currcell
                    ElseIf dtSelectRow.Item(12) = 2 Then
                        vwAppSheet.Cells("I23").Value = Nothing
                    End If

                    vwAppSheet.Cells("I24").CellType = currcell
                    vwAppSheet.Cells("I22").Formula = SUBTOTAL_FORMULA                           '小計

                    'TODO 税率取得方法
                    'vwAppSheet.Cells("G23").Value = String.Format(TAX_TEXT, "10")
                    'vwAppSheet.Cells("I23").Formula = String.Format(TAX_FORMULA, "10")

                    '貸出種別が一般利用／特例の場合、消費税を表示
                    '貸出種別が社内利用の場合、消費税は非表示
                    If dtSelectRow.Item(12) = 1 Or dtSelectRow.Item(12) = 3 Then
                        'vwAppSheet.Cells("G23").Value = String.Format(TAX_TEXT, "10")           '消費税
                        'vwAppSheet.Cells("I23").Formula = String.Format(TAX_FORMULA, "10")
                        vwAppSheet.Cells("G23").Value = String.Format(TAX_TEXT, dataEXTB0103.PropStrTax_Output)           '消費税
                        vwAppSheet.Cells("I23").Formula = String.Format(TAX_FORMULA, dataEXTB0103.PropStrTax_Output)
                    ElseIf dtSelectRow.Item(12) = 2 Then
                        vwAppSheet.Cells("G23").Value = Nothing
                        vwAppSheet.Cells("I23").Formula = Nothing
                    End If

                    vwAppSheet.Cells("I24").Formula = TOTAL_FORMULA                             '合計

                    '控えに値を設定
                    vwAppSheetNote.Cells("I5").Value = dtSelectRow.Item(0)   '承認番号
                    vwAppSheetNote.Cells("C8").Value = dtSelectRow.Item(1)   '会社・団体名
                    vwAppSheetNote.Cells("C9").Value = dtSelectRow.Item(2)   '代表者名
                    vwAppSheetNote.Cells("C10").Value = dtSelectRow.Item(3)  '利用責任者
                    vwAppSheetNote.Cells("C12").Value = dtSelectRow.Item(4)  '利用内容
                    vwAppSheetNote.Cells("C13").Value = dtSelectRow.Item(5)  '催事名
                    vwAppSheetNote.Cells("C14").Value = dtSelectRow.Item(6)  '利用形状
                    vwAppSheetNote.Cells("J14").Value = dtSelectRow.Item(7)  '定員
                    vwAppSheetNote.Cells("C15").Value = dtSelectRow.Item(8)  'ドリンク代徴収
                    vwAppSheetNote.Cells("A18").Value = dtSelectRow.Item(9)  '利用日時
                    vwAppSheetNote.Cells("G18").Value = dtSelectRow.Item(10) '利用日数
                    vwAppSheetNote.Cells("I18").Value = dtSelectRow.Item(11) '利用料金
                    vwAppSheetNote.Cells("I18").CellType = currcell
                    vwAppSheetNote.Cells("I22").CellType = currcell
                    '消費税セルタイプ
                    '貸出種別が一般利用／特例の場合￥を表示
                    '貸出種別が社内利用の場合は非表示
                    If dtSelectRow.Item(12) = 1 Or dtSelectRow.Item(12) = 3 Then
                        vwAppSheetNote.Cells("I23").CellType = currcell
                    ElseIf dtSelectRow.Item(12) = 2 Then
                        vwAppSheetNote.Cells("I23").CellType = Nothing
                    End If

                    vwAppSheetNote.Cells("I24").CellType = currcell
                    vwAppSheetNote.Cells("I22").Formula = SUBTOTAL_FORMULA                  '小計

                    'TODO 税率取得方法
                    'vwAppSheetNote.Cells("G23").Value = String.Format(TAX_TEXT, "10")
                    'vwAppSheetNote.Cells("I23").Formula = String.Format(TAX_FORMULA, "10")  '消費税

                    '貸出種別が一般利用／特例の場合、消費税を表示
                    '貸出種別が社内利用の場合、消費税は非表示
                    If dtSelectRow.Item(12) = 1 Or dtSelectRow.Item(12) = 3 Then
                        'vwAppSheetNote.Cells("G23").Value = String.Format(TAX_TEXT, "10")
                        'vwAppSheetNote.Cells("I23").Formula = String.Format(TAX_FORMULA, "10")  '消費税
                        vwAppSheetNote.Cells("G23").Value = String.Format(TAX_TEXT, dataEXTB0103.PropStrTax_Output)
                        vwAppSheetNote.Cells("I23").Formula = String.Format(TAX_FORMULA, dataEXTB0103.PropStrTax_Output)  '消費税
                    ElseIf dtSelectRow.Item(12) = 2 Then
                        vwAppSheetNote.Cells("G23").Value = Nothing
                        vwAppSheetNote.Cells("I23").Formula = Nothing  '消費税
                    End If

                    vwAppSheetNote.Cells("I24").Formula = TOTAL_FORMULA                     '合計
                End If

                '利用承認書を初期設定
                .PropVwPrintSheet.ActiveSheetIndex = 0

                ' ファイル保存ダイアログ起動
                Dim fn As String = ""
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "利用承認書(*.xlsx)|*.xlsx"
                    sfd.FileName = String.Format(CERTIFICATES_FILENAME, .PropStrReserveNo_Output)
                    If sfd.ShowDialog() = DialogResult.OK Then
                        fn = sfd.FileName

                        ' Excelファイルに保存
                        .PropVwPrintSheet.SaveExcel(fn, FarPoint.Excel.ExcelSaveFlags.UseOOXMLFormat)
                        MsgBox("Excel出力が完了しました。")
                    End If
                End Using

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細出力処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細（合算）を編集し、出力する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function OutputUseDetails(ByVal dataEXTB0103 As DataEXTB0103) As Boolean

        Dim downloadPath As String      'ダウンロードパス

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            downloadPath = ConfigurationManager.AppSettings("formatPath") & FORMAT_DETAILS
            dataEXTB0103.PropVwPrintSheet = New FarPoint.Win.Spread.FpSpread
            'Excel読み込み
            dataEXTB0103.PropVwPrintSheet.OpenExcel(downloadPath)
            'dataEXTB0103.PropVwPrintSheet.OpenExcel("..\..\Format\付帯設備利用明細書_フォーマット.xlsx")
            dataEXTB0103.PropVwPrintSheet.ActiveSheet = dataEXTB0103.PropVwPrintSheet.Sheets(0)
            '付帯設備利用明細編集、出力
            If SetUseDetailsGeneral(dataEXTB0103) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細書出力処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書（一般貸出）を編集し、出力する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetUseDetailsGeneral(ByVal dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            '結合されている場合左上のセルを参照する
            ' シート設定
            Dim vwAppSheet As FarPoint.Win.Spread.SheetView = dataEXTB0103.PropVwPrintSheet.Sheets(0)      '利用明細書
            vwAppSheet.ColumnCount = 11 '列数
            vwAppSheet.RowCount = 100    '行数

            With dataEXTB0103
                'ヘッダ
                'If dataEXTB0103.PropStrRentalClass_Output = "一般貸出" Then
                If dataEXTB0103.PropStrRentalClass_Output = "1" Then
                    If .PropBlnSumFlg Then
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "利用明細書（合算）"
                    Else
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "利用明細書"
                    End If
                    'ElseIf dataEXTB0103.PropStrRentalClass_Output = "特例" Then
                ElseIf dataEXTB0103.PropStrRentalClass_Output = "3" Then
                    If .PropBlnSumFlg Then
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "利用明細書（合算）"
                    Else
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "利用明細書"
                    End If
                Else
                    If .PropBlnSumFlg Then
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "社 内 利 用 明 細 書（合算）"
                    Else
                        .PropVwPrintSheet.ActiveSheet.Cells("A2").Value = "社 内 利 用 明 細 書"
                    End If
                End If

                .PropVwPrintSheet.ActiveSheet.Cells("B1").Value = .PropStrReserveNo_Output '予約番号
                .PropVwPrintSheet.ActiveSheet.Cells("I1").Value = Now.ToString("yyyy年MM月dd日")  '出力日

                ' 2015.11.27 UPD START↓ h.hagiwara 税有・税込どちらか未存在時の判定追加
                ''.PropVwPrintSheet.ActiveSheet.Cells("B3").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(1) '利用者名
                ''.PropVwPrintSheet.ActiveSheet.Cells("B4").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(2) '利用日時
                ''.PropVwPrintSheet.ActiveSheet.Cells("B5").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(3) '催事名＋出演者名
                ' ''ご利用明細
                ' ''基本会場利用料
                ''.PropVwPrintSheet.ActiveSheet.Cells("A8").Value = "1"
                ''.PropVwPrintSheet.ActiveSheet.Cells("B8").Value = THEATER_FEE
                ''.PropVwPrintSheet.ActiveSheet.Cells("C8").Value = "1日"


                ''If .PropBlnSumFlg Then
                ''    .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = .PropDtUseDetailsNoTax_Output(0)(4) '延べ利用日数
                ''Else
                ''    .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = 1 '延べ利用日数
                ''End If
                ''If .PropBlnSumFlg Then
                ''    .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(5) '金額
                ''    .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(6) '調整額
                ''Else
                ''    .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(4) '金額
                ''    .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = 0 '調整額
                ''End If
                If .PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then
                    .PropVwPrintSheet.ActiveSheet.Cells("B3").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(1) '利用者名
                    .PropVwPrintSheet.ActiveSheet.Cells("B4").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(2) '利用日時
                    .PropVwPrintSheet.ActiveSheet.Cells("B5").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(3) '催事名＋出演者名
                    'ご利用明細
                    '基本会場利用料
                    .PropVwPrintSheet.ActiveSheet.Cells("A8").Value = "1"
                    .PropVwPrintSheet.ActiveSheet.Cells("B8").Value = THEATER_FEE
                    .PropVwPrintSheet.ActiveSheet.Cells("C8").Value = "1日"


                    If .PropBlnSumFlg Then
                        .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = .PropDtUseDetailsNoTax_Output(0)(4) '延べ利用日数
                        .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(5) '金額
                        .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(6) '調整額
                    Else
                        If .PropDtUseDetailsNoTax_Output.Rows(0)(4) = 0 Then
                            .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = "" '延べ利用日数
                        Else
                            .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = 1 '延べ利用日数
                        End If
                        .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(4) '金額
                        .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = .PropDtUseDetailsNoTax_Output.Rows(0)(15) '調整額
                    End If
                ElseIf .PropDtUseDetails_Output.Rows.Count > 0 Then
                    .PropVwPrintSheet.ActiveSheet.Cells("B3").Value = .PropDtUseDetails_Output.Rows(0)(1) '利用者名
                    .PropVwPrintSheet.ActiveSheet.Cells("B4").Value = .PropDtUseDetails_Output.Rows(0)(2) '利用日時
                    .PropVwPrintSheet.ActiveSheet.Cells("B5").Value = .PropDtUseDetails_Output.Rows(0)(3) '催事名＋出演者名
                    'ご利用明細
                    '基本会場利用料
                    .PropVwPrintSheet.ActiveSheet.Cells("A8").Value = "1"
                    .PropVwPrintSheet.ActiveSheet.Cells("B8").Value = THEATER_FEE
                    .PropVwPrintSheet.ActiveSheet.Cells("C8").Value = "1日"

                    If .PropBlnSumFlg Then
                        .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = .PropDtUseDetails_Output(0)(4) '延べ利用日数
                        .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetails_Output.Rows(0)(5) '金額
                        .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = .PropDtUseDetails_Output.Rows(0)(6) '調整額
                    Else
                        If .PropDtUseDetails_Output.Rows(0)(4) = 0 Then
                            .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = "" '延べ利用日数
                        Else
                            .PropVwPrintSheet.ActiveSheet.Cells("E8").Value = 1  '延べ利用日数
                        End If
                        .PropVwPrintSheet.ActiveSheet.Cells("F8").Value = .PropDtUseDetails_Output.Rows(0)(4)  '金額
                        .PropVwPrintSheet.ActiveSheet.Cells("G8").Value = .PropDtUseDetails_Output.Rows(0)(15) '調整額
                    End If
                End If
                ' 2015.11.27 UPD END↑ h.hagiwara 税有・税込どちらか未存在時の判定追加

                '小計
                .PropVwPrintSheet.ActiveSheet.Cells("H8").Formula = String.Format(INCIDENT_SUBFEE, 8)
                '付帯設備利用料

                '消費税
                If dataEXTB0103.PropStrRentalClass_Output = "1" Or dataEXTB0103.PropStrRentalClass_Output = "3" Then
                    .PropVwPrintSheet.ActiveSheet.Cells("E48").Formula = CInt(.PropStrTax_Output) / 100
                End If

                '利用料＋全付帯設備の小計
                .PropVwPrintSheet.ActiveSheet.Cells("F47").Formula = "SUM(H8:H46)"
                '.PropVwPrintSheet.ActiveSheet.Cells("G47").Formula = "SUM(G8:G46)"
                '.PropVwPrintSheet.ActiveSheet.Cells("H47").Formula = "SUM(H8:H46)"
                '消費税額
                .PropVwPrintSheet.ActiveSheet.Cells("F48").Formula = "F47*E48"
                '.PropVwPrintSheet.ActiveSheet.Cells("G48").Formula = "G47*E48"
                '.PropVwPrintSheet.ActiveSheet.Cells("H48").Formula = "H47*E48"
                '小計＋消費税
                .PropVwPrintSheet.ActiveSheet.Cells("F49").Formula = "F47+F48"
                '.PropVwPrintSheet.ActiveSheet.Cells("G49").Formula = "G47+G48"
                '.PropVwPrintSheet.ActiveSheet.Cells("H49").Formula = "H47+H48"

                If .PropBlnSumFlg Then
                    '合算の場合
                    'If .PropDtUseDetails_Output IsNot Nothing Then                   ' 2015.11.27 UPD h.hagiwara
                    If .PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then              ' 2015.11.27 UPD h.hagiwara
                        For i As Integer = 0 To .PropDtUseDetailsNoTax_Output.Rows.Count - 1
                            'データを設定するスプレッド行数
                            Dim j As Integer = i + 8
                            'DataTableから一行取得
                            Dim dtSelectRow As DataRow = .PropDtUseDetailsNoTax_Output.Rows(i)
                            'スプレッドシートに値設定
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 0).Value = i + 2 'インデックス
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 1).Value = dtSelectRow.Item(7) '項目名
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 2).Value = dtSelectRow.Item(8) '単位
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 3).Value = dtSelectRow.Item(9) '単価
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 4).Value = dtSelectRow.Item(10) '数量
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 5).Value = dtSelectRow.Item(11) '金額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 6).Value = dtSelectRow.Item(12) '調整額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 7).Formula = String.Format(INCIDENT_SUBFEE, j + 1) '小計
                        Next
                    End If
                    '楽屋ケイタリングサービス・お立替経費・追加人件費他
                    'If .PropDtUseDetailsNoTax_Output IsNot Nothing Then               ' 2015.11.27 UPD h.hagiwara
                    If .PropDtUseDetails_Output.Rows.Count > 0 Then                    ' 2015.11.27 UPD h.hagiwara
                        For i As Integer = 0 To .PropDtUseDetails_Output.Rows.Count - 1
                            'データを設定するスプレッド行数
                            Dim j As Integer = i + 51
                            'DataTableから一行取得
                            Dim dtSelectRow As DataRow = .PropDtUseDetails_Output.Rows(i)
                            'スプレッドシートに値設定
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 0).Value = i + 1
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 1).Value = dtSelectRow.Item(7) '項目名
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 2).Value = dtSelectRow.Item(8) '単位
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 3).Value = dtSelectRow.Item(9) '単価
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 4).Value = dtSelectRow.Item(10) '数量
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 5).Value = dtSelectRow.Item(11) '金額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 6).Value = dtSelectRow.Item(12) '調整額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 7).Formula = String.Format(INCIDENT_SUBFEE, j + 1) '小計
                        Next
                    End If
                Else
                    '日別の場合
                    'If .PropDtUseDetails_Output IsNot Nothing Then                   ' 2015.11.27 UPD h.hagiwara
                    If .PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then              ' 2015.11.27 UPD h.hagiwara
                        For i As Integer = 0 To .PropDtUseDetailsNoTax_Output.Rows.Count - 1
                            'データを設定するスプレッド行数
                            Dim j As Integer = i + 8
                            'DataTableから一行取得
                            Dim dtSelectRow As DataRow = .PropDtUseDetailsNoTax_Output.Rows(i)
                            'スプレッドシートに値設定
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 0).Value = i + 2 'インデックス
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 1).Value = dtSelectRow.Item(5) '項目名
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 2).Value = dtSelectRow.Item(6) '単位
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 3).Value = dtSelectRow.Item(7) '単価
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 4).Value = dtSelectRow.Item(8) '数量
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 5).Value = dtSelectRow.Item(9) '金額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 6).Value = dtSelectRow.Item(10) '調整額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 7).Formula = String.Format(INCIDENT_SUBFEE, j + 1) '小計
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 8).Value = dtSelectRow.Item(12) '備考
                        Next
                    End If
                    '楽屋ケイタリングサービス・お立替経費・追加人件費他
                    'If .PropDtUseDetailsNoTax_Output IsNot Nothing Then               ' 2015.11.27 UPD h.hagiwara
                    If .PropDtUseDetails_Output.Rows.Count > 0 Then                    ' 2015.11.27 UPD h.hagiwara
                        For i As Integer = 0 To .PropDtUseDetails_Output.Rows.Count - 1
                            'データを設定するスプレッド行数
                            Dim j As Integer = i + 51
                            'DataTableから一行取得
                            Dim dtSelectRow As DataRow = .PropDtUseDetails_Output.Rows(i)
                            'スプレッドシートに値設定
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 0).Value = i + 1
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 1).Value = dtSelectRow.Item(5) '項目名
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 2).Value = dtSelectRow.Item(6) '単位
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 3).Value = dtSelectRow.Item(7) '単価
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 4).Value = dtSelectRow.Item(8) '数量
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 5).Value = dtSelectRow.Item(9) '金額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 6).Value = dtSelectRow.Item(10) '調整額
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 7).Formula = String.Format(INCIDENT_SUBFEE, j + 1) '小計
                            .PropVwPrintSheet.ActiveSheet.Cells(j, 8).Value = dtSelectRow.Item(12) '備考
                        Next
                    End If
                End If
                '楽屋ケイタリングサービス・お立替経費・追加人件費他合計
                .PropVwPrintSheet.ActiveSheet.Cells("F57").Formula = "SUM(H52:H56)"
                '.PropVwPrintSheet.ActiveSheet.Cells("G57").Formula = "SUM(G52:G56)"
                '.PropVwPrintSheet.ActiveSheet.Cells("H57").Formula = "SUM(H52:H56)"

                'ご請求額
                .PropVwPrintSheet.ActiveSheet.Cells("F59").Formula = "F49+F57"


                ' ファイル保存ダイアログ起動
                Dim fn As String = ""
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "利用明細書(*.xlsx)|*.xlsx"
                    If .PropBlnSumFlg Then
                        '合算の場合、予約番号がファイル名に入るようにする
                        sfd.FileName = String.Format(DETAILS_FILENAME_ALL, .PropStrReserveNo_Output)
                    Else
                        '日別の場合、予約番号、出力対象日がファイル名に入るようにする
                        sfd.FileName = String.Format(DETAILS_FILENAME_ONEDAY, .PropStrReserveNo_Output, _
                                                     DateTime.Parse(.PropStrUseDays_Output).ToString("yyyyMMdd"))
                    End If
                    If sfd.ShowDialog() = DialogResult.OK Then
                        fn = sfd.FileName

                        ' Excelファイルに保存
                        .PropVwPrintSheet.SaveExcel(fn, FarPoint.Excel.ExcelSaveFlags.UseOOXMLFormat)
                        MsgBox("Excel出力が完了しました。")
                    End If
                End Using

            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            dataEXTB0103.PropStrClickedIncident = Nothing
            dataEXTB0103.PropDtUseDetails_Output.Dispose()
            dataEXTB0103.PropDtUseDetailsNoTax_Output.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 利用承認書取得処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより利用承認書を取得する
    ''' <para>作成情報：2015/08/28 y.naganuma
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetUseApproval(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtUseApproval As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            '利用承認書取得SQLを設定
            If sqlEXTB0103.SetSelectUseApprovalSql(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '利用承認書を取得
            Adapter.Fill(dtUseApproval)
            dataEXTB0103.PropDtUseApproval_Output = dtUseApproval

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtUseApproval.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細（合算）取得処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより付帯設備利用明細（合算）を取得する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetUseDetailsAll(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtUseDetailsAll As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            '利用明細（税有）を取得
            dataEXTB0103.PropBlnNoTaxFlg = False
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTB0103.SetSelectUseDetailsAllSql(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsAll)
            dataEXTB0103.PropDtUseDetails_Output = dtUseDetailsAll

            dtUseDetailsAll = New DataTable

            '利用明細（税無）を取得
            dataEXTB0103.PropBlnNoTaxFlg = True
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTB0103.SetSelectUseDetailsAllSql(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '付帯設備利用明細（税無）を取得
            Adapter.Fill(dtUseDetailsAll)
            dataEXTB0103.PropDtUseDetailsNoTax_Output = dtUseDetailsAll

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtUseDetailsAll.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細（日別）取得処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより付帯設備利用明細（日別）を取得する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetUseDetailsOneDay(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtUseDetailsOneDay As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            '利用明細（税有）を取得
            dataEXTB0103.PropBlnNoTaxFlg = False
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTB0103.SetSelectUseDetailsOneDaySql(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsOneDay)
            dataEXTB0103.PropDtUseDetails_Output = dtUseDetailsOneDay

            dtUseDetailsOneDay = New DataTable

            '利用明細（税無）を取得
            dataEXTB0103.PropBlnNoTaxFlg = True
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTB0103.SetSelectUseDetailsOneDaySql(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '付帯設備利用明細（税無）を取得
            Adapter.Fill(dtUseDetailsOneDay)
            dataEXTB0103.PropDtUseDetailsNoTax_Output = dtUseDetailsOneDay

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            dtUseDetailsOneDay.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 消費税取得処理
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>消費税マスタより、請求日に対する消費税率を取得する
    ''' <para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetTax(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtTax As New DataTable
        Dim strTax As String = ""
        Dim Cn1 As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn1.Open()

            '消費税取得SQLを設定
            If sqlEXTB0103.SetSelectTaxSql(Adapter, Cn1, dataEXTB0103) = False Then
                Return False
            End If

            '消費税マスタより消費税を取得
            Adapter.Fill(dtTax)
            dataEXTB0103.PropStrTax_Output = dtTax.Rows(0).Item(0).ToString

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        Finally
            'コネクションが閉じられていない場合、コネクションを閉じる
            If Cn1 IsNot Nothing Then
                Cn1.Close()
            End If
            Adapter.Dispose()
            Cn1.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' フィールドクリア
    ''' </summary>
    ''' <param name="dataEXTB0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>SQLに渡す変数を初期化
    ''' <para>作成情報：2015/09/02　h.mori
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function FildClear(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            '予約番号
            dataEXTB0103.PropStrReserveNo_Output = ""
            '日付
            dataEXTB0103.PropStrCalculateDay_Output = ""
            'クリックされた日付
            dataEXTB0103.PropStrUseDays_Output = ""
            '貸出区分
            dataEXTB0103.PropStrRentalClass_Output = ""

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            Common.CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 入金予定日の曜日判定
    ''' </summary>
    ''' <param name="CehckDay">[IN/OUT]算出元日付</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>請求日から自動算出した入金予定日が土日祝祭日の場合前営業日にする
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function WeekDayCheck(ByRef CehckDay As DateTime) As DateTime

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim dtNyukin As DateTime
        Dim strStopflg As String = ""
        Dim intweek As Integer

        dtNyukin = CehckDay

        While strStopflg = ""
            intweek = Weekday(dtNyukin)
            If intweek = 1 Or intweek = 7 Then                   ' 土日の場合は前日にし再度確認する
                dtNyukin = dtNyukin.AddDays(-1)
            Else
                If GetSelectHolidayMst(dtNyukin) = True Then        ' 祝祭日マスタに日付が存在するか判定する
                    dtNyukin = dtNyukin.AddDays(-1)              ' 祝祭日マスタに日付が存在した場合は前日にし再度確認する
                Else
                    strStopflg = "1"
                End If
            End If
        End While

        WeekDayCheck = dtNyukin

    End Function

    ''' <summary>
    ''' 祝祭日マスタの取得準備
    ''' <paramref name="CehckDay">対象日付</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>サーバ日時取得
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetSelectHolidayMst(ByRef CehckDay As DateTime) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる

        Try
            'コネクションを開く
            Cn.Open()

            '新規Inc番号、システム日付取得（SELECT）
            If GetSelectHoliday(Cn, CehckDay) = False Then
                Return False
            End If

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            Return False
        Finally
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 祝祭日マスタの取得
    ''' </summary>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="CehckDay">[IN/OUT]判定対象日付</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>祝祭日を取得する
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetSelectHoliday(ByVal Cn As NpgsqlConnection, _
                               ByRef CehckDay As DateTime) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlDataAdapter
        Dim dtResult As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GetHolidayMst(Adapter, Cn, CehckDay) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "祝祭日", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtResult)

            'データが取得できなかった場合は
            If dtResult.Rows.Count = 0 Then
                Return False
            Else
                If dtResult.Rows(0).Item(0) = 0 Then
                    Return False
                End If
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' メール送信処理
    ''' </summary>
    ''' <param name="strMailTo">宛先To</param>
    ''' <param name="strMailCc">宛先Cc</param>
    ''' <param name="strMailBCc">宛先Bcc</param>
    ''' <param name="StrSubject">件名</param>
    ''' <param name="StrBody">本文</param>
    ''' <param name="strFile">添付ファイル</param>
    ''' <param name="strFlg">非同期</param>
    ''' <returns>
    ''' </returns>
    ''' <remarks>メールの編集・送信を行なう
    ''' <para>作成情報：2012/08/20 kudo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function Mailsend(ByVal strMailTo As String, ByVal strMailCc As String, ByVal strMailBcc As String, _
                             ByVal StrSubject As String, ByVal StrBody As String, ByVal strFile As String, ByVal strFlg As String) As Boolean
        '開始システムログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim smtp As New SmtpClient()
        Dim msg As New MailMessage()
        Dim enc As Encoding = Encoding.GetEncoding("iso-2022-jp")
        Dim aryMailTo As String()
        Dim aryMailCc As String()
        Dim aryMailBcc As String()
        Dim aryFile As String()

        Try

            '差出人
            msg.From = New MailAddress(MailFrom, MailFromName, enc)

            '宛先(To)
            If strMailTo <> "" Then
                aryMailTo = Split(strMailTo, ",")
                For i As Integer = 0 To aryMailTo.Length - 1
                    msg.To.Add(New MailAddress(aryMailTo(i)))
                Next
            End If

            '宛先(Cc)
            If strMailCc <> "" Then
                aryMailCc = Split(strMailCc, ",")
                For i As Integer = 0 To aryMailCc.Length - 1
                    msg.CC.Add(New MailAddress(aryMailCc(i)))
                Next
            End If

            '宛先(Bcc)
            If strMailBcc <> "" Then
                aryMailBcc = Split(strMailBcc, ",")
                For i As Integer = 0 To aryMailBcc.Length - 1
                    msg.Bcc.Add(New MailAddress(aryMailBcc(i)))
                Next
            End If

            '添付ファイル
            If strFile <> "" Then
                aryFile = Split(strFile, ",")
                For i As Integer = 0 To aryFile.Length - 1
                    msg.Attachments.Add(New Attachment(aryFile(i)))
                Next
            End If

            '件名
            msg.Subject = StrSubject
            msg.SubjectEncoding = enc

            '本文
            msg.Body = StrBody
            msg.BodyEncoding = enc

            'ＳＭＴＰサーバー
            smtp.Host = MailSmtp

            'SMTP認証
            If SmtpAuth = "1" Then
                smtp.Credentials = New NetworkCredential(SmtpUserId, SmtpPass)
            End If

            ''メール送信
            'If AsyncFlg = "1" Then
            '    smtp.SendAsync(msg, Nothing)
            'Else
            '    smtp.Send(msg)
            'End If

            'メール送信
            smtp.Send(msg)


            '終了システムログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'エラーシステムログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 消費税率設定判定情報取得処理
    ''' <paramref name="dataEXTB0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>EXAS連携時に消費税率の設定
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetNotaxflg(ByRef dataEXTB0103 As DataEXTB0103) As String

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ
        Dim strTaxlfg As String = "0"

        Try
            'コネクションを開く
            Cn.Open()

            ' 税抜きフラグの取得
            If GetNotaxflgInf(Adapter, Cn, dataEXTB0103) = False Then
                strTaxlfg = "0"
            Else
                strTaxlfg = dataEXTB0103.PropStrNoTaxFlg
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return strTaxlfg

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 消費税率設定判定情報取得処理用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTB0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS連携時に消費税率の設定を取得する
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetNotaxflgInf(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GeSqlNotaxflg(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得（ＣＳＶ取り込み処理）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst Is Nothing Then
                Return False
            Else
                If dtmst.Rows.Count > 0 Then
                    dataEXTB0103.PropStrNoTaxFlg = dtmst.Rows(0).Item(0)
                Else
                    Return False
                End If
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

    ''' <summary>
    ''' 相手先情報取得処理
    ''' <paramref name="dataEXTB0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>相手先情報の取得
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetAitesakiInf(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ
        Dim strTaxlfg As String = "0"

        Try
            'コネクションを開く
            Cn.Open()

            ' 税抜きフラグの取得
            If GetAitesaki(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            Else
                strTaxlfg = dataEXTB0103.PropStrNoTaxFlg
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 相手先情報取得処理用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTB0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS連携時に消費税率の設定を取得する
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetAitesaki(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GeSqlAitesaki(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "予約情報取得（ＣＳＶ取り込み処理）", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst Is Nothing Then
                Return False
            Else
                If dtmst.Rows.Count > 0 Then
                    dataEXTB0103.PropStrPostno = dtmst.Rows(0).Item(0)
                    dataEXTB0103.PropStrAddr1 = dtmst.Rows(0).Item(1)
                    dataEXTB0103.PropStrAddr2 = dtmst.Rows(0).Item(2)
                    dataEXTB0103.PropStrAitenm = dtmst.Rows(0).Item(3)
                Else
                    Return False
                End If
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

    ''' <summary>
    ''' 予約情報の論理削除更新処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報の論理削除更新
    ''' <para>作成情報： 2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegYoyakuInfo(ByRef dataEXTB0102 As DataEXTB0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Cmd As NpgsqlCommand = Cn.CreateCommand

        Dim Table As New DataTable
        Dim Tsx As NpgsqlTransaction = Nothing

        Try
            Cn.Open()

            'トランザクション開始
            Tsx = Cn.BeginTransaction
            '予約情報
            sqlEXTB0103.registerYoyakuInfo(Cmd, Cn, dataEXTB0102)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "Complate SQL", Nothing, Cmd)

            'トランザクションCOMMIT
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
            Throw ex
            Return False
        Finally
            '終了処理
            Cn.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 請求情報入力済判定処理
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>請求情報入力済の情報が存在するか判定する
    ''' <para>作成情報： 2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetBillpayFlg(ByRef dataEXTB0102 As DataEXTB0102) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtBillPay As DataTable = dataEXTB0102.PropDsBillReq.Tables("BILLPAY_TBL")
        Dim intCnt As Integer = 0
        Dim row As DataRow

        For i = 0 To dtBillPay.Rows.Count - 1
            row = dtBillPay.Rows(i)
            If row("seikyu_input_flg") Is DBNull.Value = False Then
                If row("seikyu_input_flg") = "1" Then
                    intCnt += 1
                End If
            End If
        Next

        If intCnt > 0 Then
            ' 請求入力した情報が存在
            Return True
        Else
            ' 請求入力した情報が未存在
            Return False
        End If

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常終了
        Return True

    End Function

    ''' <summary>
    ''' 料金マスタ・倍率マスタ情報取得処理
    ''' <paramref name="dataEXTB0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>料金マスタ・倍率マスタ情報の取得
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetRyokinBairituMst(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            ' 料金マスタの取得
            If GetRyokinMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            ' 倍率マスタの取得
            If GeBairituMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 料金マスタ情報取得処理用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTB0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>料金・倍率選択画面表示前に料金マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetRyokinMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GeSqlRyokinMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "遷移前料金マスタ取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count > 0 Then
                If dtmst.Rows(0).Item(0) > 0 Then
                Else
                    Return False
                End If
            Else
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

    ''' <summary>
    ''' 倍率マスタ情報取得処理用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTB0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>料金・倍率選択画面表示前に倍率マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GeBairituMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GeSqlBairituMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "遷移前倍率マスタ取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count > 0 Then
                If dtmst.Rows(0).Item(0) > 0 Then
                Else
                    Return False
                End If
            Else
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

    ''' <summary>
    ''' 付帯設備マスタ情報取得処理
    ''' <paramref name="dataEXTB0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>付帯設備マスタ情報の取得
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetFutaiSetubiInfMst(ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            ' 付帯設備マスタの取得
            If GetFutaiSetubiMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = E0000 & ex.Message
            Return False
        Finally
            Cn.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタ情報取得処理用
    ''' </summary>
    ''' <param name="Adapter">[IN]NpgsqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgsqlConnectionクラス</param>
    ''' <param name="DataEXTB0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備設定画面表示前に付帯設備マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetFutaiSetubiMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTB0103 As DataEXTB0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTB0103.GeSqlFutaiSetubiMst(Adapter, Cn, dataEXTB0103) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "遷移前付帯設備マスタ取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(dtmst)

            'データが取得できなかった場合、エラー
            If dtmst.Rows.Count > 0 Then
                If dtmst.Rows(0).Item(0) > 0 Then
                Else
                    Return False
                End If
            Else
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

    ''' <summary>
    ''' 責任者メールアドレス・携帯電話番号取得
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>責任者名のコンボボックス選択時メールアドレス・携帯電話番号取得
    ''' <para>作成情報：2016/11/04 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetSekininshaMailTel(ByVal dataEXTB0102 As DataEXTB0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable()

        Try
            'コネクションを開く
            Cn.Open()

            'SQLの作成・設定
            If sqlEXTB0103.GetSqlSekininshaMailTelKakuData(Adapter, Cn, dataEXTB0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "責任者メールアドレス・携帯電話番号情報取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得した項目を変数に格納
            With dataEXTB0102
                'メールアドレス
                .PropStrSekininMail = IIf(DBNull.Value.Equals(Table.Rows(0).Item("SEKININ_MAIL")), "", Table.Rows(0).Item("SEKININ_MAIL"))
                '携帯電話番号
                .PropStrRiyoTel21 = IIf(DBNull.Value.Equals(Table.Rows(0).Item("RIYO_TEL21")), "", Table.Rows(0).Item("RIYO_TEL21"))
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception

            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)

            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

End Class
